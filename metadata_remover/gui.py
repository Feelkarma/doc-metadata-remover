"""Tkinter desktop GUI for doc-metadata-remover.

Features:
    * Drag-and-drop files onto the window (requires the optional
      ``tkinterdnd2`` package; the app degrades gracefully without it).
    * "Add files" picker button and a "Clear" button.
    * Batch processing of many files at once.
    * Determinate progress bar plus a live log.
    * Optional output-folder selection (defaults to each file's own folder).
    * A clean, modern light theme.

The heavy lifting happens in :mod:`metadata_remover.core`; the GUI only wires
it up on a background thread so the interface stays responsive.
"""

from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional

from .core import SUPPORTED_EXTENSIONS, clean_document

# Optional drag-and-drop support.
try:  # pragma: no cover - depends on optional dependency
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _DND_AVAILABLE = True
except Exception:  # pragma: no cover
    _DND_AVAILABLE = False


# ---------------------------------------------------------------------------
# Colour palette (clean, modern light theme)
# ---------------------------------------------------------------------------
BG = "#f4f6fb"
CARD = "#ffffff"
ACCENT = "#3b6ef5"
ACCENT_DARK = "#2f58c8"
TEXT = "#1f2430"
MUTED = "#6b7280"
BORDER = "#dfe3ec"
OK = "#1a7f37"
ERR = "#c0362c"


class MetadataRemoverApp:
    """The main application window."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.files: List[str] = []
        self.output_dir: Optional[str] = None
        self._queue: "queue.Queue" = queue.Queue()
        self._worker: Optional[threading.Thread] = None

        root.title("Doc Metadata Remover")
        root.geometry("720x620")
        root.minsize(620, 560)
        root.configure(bg=BG)

        self._build_styles()
        self._build_ui()
        self._poll_queue()

    # -- styling ------------------------------------------------------------
    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("Card.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 18, "bold"))
        style.configure("Sub.TLabel", background=BG, foreground=MUTED, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9))

        style.configure(
            "Accent.TButton",
            background=ACCENT,
            foreground="#ffffff",
            font=("Segoe UI", 10, "bold"),
            borderwidth=0,
            focusthickness=0,
            padding=(16, 9),
        )
        style.map(
            "Accent.TButton",
            background=[("active", ACCENT_DARK), ("disabled", "#a9b7e8")],
        )
        style.configure(
            "Ghost.TButton",
            background=CARD,
            foreground=TEXT,
            font=("Segoe UI", 10),
            borderwidth=1,
            padding=(14, 8),
        )
        style.map("Ghost.TButton", background=[("active", "#eef1f8")])

        style.configure(
            "Bar.Horizontal.TProgressbar",
            troughcolor="#e6e9f2",
            background=ACCENT,
            borderwidth=0,
            thickness=10,
        )

    # -- UI layout ----------------------------------------------------------
    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, style="TFrame")
        outer.pack(fill="both", expand=True, padx=20, pady=18)

        ttk.Label(outer, text="Doc Metadata Remover", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Strip author, dates, comments & hidden properties from DOCX, "
            "PPTX and XLSX \u2014 content stays untouched.",
            style="Sub.TLabel",
            wraplength=680,
            justify="left",
        ).pack(anchor="w", pady=(2, 14))

        # Drop zone card.
        drop = tk.Frame(outer, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        drop.pack(fill="both", expand=False, ipady=6)
        self.drop_zone = drop

        hint = "Drag & drop files here" if _DND_AVAILABLE else "Add files to get started"
        self.drop_label = tk.Label(
            drop,
            text=f"\U0001F4C4  {hint}",
            bg=CARD,
            fg=MUTED,
            font=("Segoe UI", 12),
            pady=18,
        )
        self.drop_label.pack(fill="x")

        btn_row = tk.Frame(drop, bg=CARD)
        btn_row.pack(pady=(0, 14))
        ttk.Button(btn_row, text="Add files", style="Accent.TButton", command=self.add_files).pack(side="left", padx=6)
        ttk.Button(btn_row, text="Clear", style="Ghost.TButton", command=self.clear_files).pack(side="left", padx=6)

        if _DND_AVAILABLE:
            for w in (drop, self.drop_label):
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self._on_drop)

        # File list card.
        list_card = tk.Frame(outer, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        list_card.pack(fill="both", expand=True, pady=(14, 0))

        head = tk.Frame(list_card, bg=CARD)
        head.pack(fill="x", padx=12, pady=(10, 4))
        self.count_label = tk.Label(head, text="No files selected", bg=CARD, fg=TEXT, font=("Segoe UI", 10, "bold"))
        self.count_label.pack(side="left")

        list_wrap = tk.Frame(list_card, bg=CARD)
        list_wrap.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        scroll = tk.Scrollbar(list_wrap)
        scroll.pack(side="right", fill="y")
        self.listbox = tk.Listbox(
            list_wrap,
            yscrollcommand=scroll.set,
            activestyle="none",
            bg="#fbfcfe",
            fg=TEXT,
            highlightthickness=1,
            highlightbackground=BORDER,
            selectbackground="#e6ecff",
            selectforeground=TEXT,
            borderwidth=0,
            font=("Segoe UI", 10),
        )
        self.listbox.pack(side="left", fill="both", expand=True)
        scroll.config(command=self.listbox.yview)

        # Output folder row.
        out_row = tk.Frame(outer, bg=BG)
        out_row.pack(fill="x", pady=(14, 6))
        self.output_var = tk.StringVar(value="Output: same folder as each file")
        ttk.Button(out_row, text="Output folder\u2026", style="Ghost.TButton", command=self.choose_output).pack(side="left")
        ttk.Button(out_row, text="Reset", style="Ghost.TButton", command=self.reset_output).pack(side="left", padx=(8, 0))
        tk.Label(out_row, textvariable=self.output_var, bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(side="left", padx=12)

        # Progress + log.
        self.progress = ttk.Progressbar(outer, style="Bar.Horizontal.TProgressbar", maximum=100)
        self.progress.pack(fill="x", pady=(8, 4))

        self.log = tk.Text(
            outer,
            height=6,
            bg="#0f1524",
            fg="#d5dcec",
            insertbackground="#d5dcec",
            font=("Consolas", 9),
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            state="disabled",
            wrap="word",
        )
        self.log.pack(fill="both", expand=False, pady=(4, 8))
        self.log.tag_configure("ok", foreground="#79e08a")
        self.log.tag_configure("err", foreground="#ff8a80")
        self.log.tag_configure("info", foreground="#9db4ff")

        # Action row.
        action = tk.Frame(outer, bg=BG)
        action.pack(fill="x")
        self.run_btn = ttk.Button(action, text="Remove metadata", style="Accent.TButton", command=self.run)
        self.run_btn.pack(side="right")
        self.status_label = tk.Label(action, text="Ready", bg=BG, fg=MUTED, font=("Segoe UI", 9))
        self.status_label.pack(side="left")

    # -- file management ----------------------------------------------------
    def _add_paths(self, paths: List[str]) -> None:
        added = 0
        for p in paths:
            p = p.strip()
            if not p:
                continue
            if os.path.splitext(p)[1].lower() in SUPPORTED_EXTENSIONS and os.path.isfile(p):
                if p not in self.files:
                    self.files.append(p)
                    self.listbox.insert("end", os.path.basename(p))
                    added += 1
        self._refresh_count()
        if added:
            self._log(f"Added {added} file(s).", "info")

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Select Office documents",
            filetypes=[
                ("Office documents", "*.docx *.pptx *.xlsx"),
                ("Word", "*.docx"),
                ("PowerPoint", "*.pptx"),
                ("Excel", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
        if paths:
            self._add_paths(list(paths))

    def clear_files(self) -> None:
        self.files.clear()
        self.listbox.delete(0, "end")
        self._refresh_count()

    def _on_drop(self, event) -> None:  # pragma: no cover - GUI event
        # tkinterdnd2 returns a brace-wrapped, space-separated list.
        paths = self.root.tk.splitlist(event.data)
        self._add_paths(list(paths))

    def _refresh_count(self) -> None:
        n = len(self.files)
        self.count_label.config(text="No files selected" if n == 0 else f"{n} file(s) selected")

    def choose_output(self) -> None:
        d = filedialog.askdirectory(title="Choose output folder")
        if d:
            self.output_dir = d
            self.output_var.set(f"Output: {d}")

    def reset_output(self) -> None:
        self.output_dir = None
        self.output_var.set("Output: same folder as each file")

    # -- processing ---------------------------------------------------------
    def run(self) -> None:
        if self._worker and self._worker.is_alive():
            return
        if not self.files:
            messagebox.showinfo("No files", "Please add one or more files first.")
            return
        self.run_btn.config(state="disabled")
        self.progress["value"] = 0
        self._set_status("Processing\u2026")
        files = list(self.files)
        out_dir = self.output_dir
        self._worker = threading.Thread(target=self._work, args=(files, out_dir), daemon=True)
        self._worker.start()

    def _work(self, files: List[str], out_dir: Optional[str]) -> None:
        total = len(files)
        ok = 0
        for i, path in enumerate(files, start=1):
            result = clean_document(path, output_dir=out_dir)
            if result.success:
                ok += 1
                self._queue.put(("log", f"\u2713 {os.path.basename(path)}  \u2192  {os.path.basename(result.output_path)}", "ok"))
            else:
                self._queue.put(("log", f"\u2717 {os.path.basename(path)}  \u2014  {result.error}", "err"))
            self._queue.put(("progress", (i / total) * 100, None))
        self._queue.put(("done", (ok, total), None))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload, tag = self._queue.get_nowait()
                if kind == "log":
                    self._log(payload, tag or "info")
                elif kind == "progress":
                    self.progress["value"] = payload
                elif kind == "done":
                    ok, total = payload
                    self.run_btn.config(state="normal")
                    self._set_status(f"Done \u2014 {ok}/{total} cleaned")
                    self._log(f"Finished: {ok} of {total} file(s) cleaned.", "info")
                    if ok:
                        messagebox.showinfo("Complete", f"Cleaned {ok} of {total} file(s).")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # -- helpers ------------------------------------------------------------
    def _log(self, msg: str, tag: str = "info") -> None:
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

    def _set_status(self, text: str) -> None:
        self.status_label.config(text=text)


def main() -> None:
    """Launch the GUI application."""
    if _DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    MetadataRemoverApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
