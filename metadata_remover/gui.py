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

import json
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List, Optional

from .core import (
    SUPPORTED_EXTENSIONS,
    capabilities,
    clean_document,
)

# Small config file used to remember the "don't show intro again" choice.
_CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".doc_metadata_remover.json")

# Optional drag-and-drop support.
try:  # pragma: no cover - depends on optional dependency
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _DND_AVAILABLE = True
except Exception:  # pragma: no cover
    _DND_AVAILABLE = False


# ---------------------------------------------------------------------------
# Colour palette (clean, modern light theme)
# ---------------------------------------------------------------------------
BG = "#eef1f8"
CARD = "#ffffff"
HEADER = "#1b2540"
HEADER_SUB = "#9aa7c7"
ACCENT = "#3b6ef5"
ACCENT_DARK = "#2f58c8"
TEXT = "#1f2430"
MUTED = "#6b7280"
BORDER = "#e2e6f0"
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
        root.geometry("760x720")
        root.minsize(680, 640)
        root.configure(bg=BG)

        self._build_styles()
        self._build_ui()
        self._poll_queue()

        # Show the "how it works" dialog on startup (unless disabled).
        if not self._intro_disabled():
            self.root.after(250, self._show_intro_dialog)

    # -- styling ------------------------------------------------------------
    def _build_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Header.TFrame", background=HEADER)
        style.configure("Bottom.TFrame", background=CARD)
        style.configure("TLabel", background=BG, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("Card.TLabel", background=CARD, foreground=TEXT, font=("Segoe UI", 10))
        style.configure("Title.TLabel", background=HEADER, foreground="#ffffff", font=("Segoe UI", 20, "bold"))
        style.configure("Sub.TLabel", background=HEADER, foreground=HEADER_SUB, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=CARD, foreground=MUTED, font=("Segoe UI", 9))

        style.configure(
            "Accent.TButton",
            background=ACCENT,
            foreground="#ffffff",
            font=("Segoe UI", 11, "bold"),
            borderwidth=0,
            focusthickness=0,
            padding=(22, 12),
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
        # === Header bar (full-width, dark) =================================
        header = tk.Frame(self.root, bg=HEADER)
        header.pack(side="top", fill="x")
        header_inner = tk.Frame(header, bg=HEADER)
        header_inner.pack(fill="x", padx=24, pady=18)
        tk.Label(
            header_inner,
            text="\U0001F512  Doc Metadata Remover",
            bg=HEADER,
            fg="#ffffff",
            font=("Segoe UI", 20, "bold"),
        ).pack(anchor="w")
        tk.Label(
            header_inner,
            text="Strip author, dates, comments & hidden properties from Word, "
            "PowerPoint, Excel (modern & legacy) and PDF \u2014 content stays "
            "untouched.",
            bg=HEADER,
            fg=HEADER_SUB,
            font=("Segoe UI", 10),
            wraplength=700,
            justify="left",
        ).pack(anchor="w", pady=(4, 0))

        # === Bottom zone (packed FIRST so it is ALWAYS visible) ============
        # Holds progress bar, log and the action row so none of them can ever
        # be pushed off-screen, regardless of window height.
        bottom = tk.Frame(self.root, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        bottom.pack(side="bottom", fill="x")

        bottom_inner = tk.Frame(bottom, bg=CARD)
        bottom_inner.pack(fill="x", padx=24, pady=(14, 16))

        # Progress bar.
        self.progress = ttk.Progressbar(bottom_inner, style="Bar.Horizontal.TProgressbar", maximum=100)
        self.progress.pack(fill="x", pady=(0, 8))

        # Live log.
        self.log = tk.Text(
            bottom_inner,
            height=4,
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
        self.log.pack(fill="x", expand=False, pady=(0, 12))
        self.log.tag_configure("ok", foreground="#79e08a")
        self.log.tag_configure("err", foreground="#ff8a80")
        self.log.tag_configure("info", foreground="#9db4ff")

        # Action row.
        action = tk.Frame(bottom_inner, bg=CARD)
        action.pack(fill="x")
        self.run_btn = ttk.Button(action, text="\u2728  Remove metadata", style="Accent.TButton", command=self.run)
        self.run_btn.pack(side="right")
        self.status_label = tk.Label(action, text="Ready", bg=CARD, fg=MUTED, font=("Segoe UI", 10))
        self.status_label.pack(side="left")

        # === Middle content area (expands to fill remaining space) =========
        outer = ttk.Frame(self.root, style="TFrame")
        outer.pack(side="top", fill="both", expand=True, padx=24, pady=18)

        # Drop zone card.
        drop = tk.Frame(outer, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        drop.pack(fill="x", expand=False, ipady=8)
        self.drop_zone = drop

        hint = "Drag & drop files here" if _DND_AVAILABLE else "Add files to get started"
        self.drop_label = tk.Label(
            drop,
            text=f"\U0001F4C4  {hint}",
            bg=CARD,
            fg=MUTED,
            font=("Segoe UI", 13),
            pady=16,
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

        # Output folder row.
        out_row = tk.Frame(outer, bg=BG)
        out_row.pack(fill="x", pady=(14, 0))
        self.output_var = tk.StringVar(value="Output: same folder as each file")
        ttk.Button(out_row, text="Output folder\u2026", style="Ghost.TButton", command=self.choose_output).pack(side="left")
        ttk.Button(out_row, text="Reset", style="Ghost.TButton", command=self.reset_output).pack(side="left", padx=(8, 0))
        tk.Label(out_row, textvariable=self.output_var, bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(side="left", padx=12)

        # File list card.
        list_card = tk.Frame(outer, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        list_card.pack(fill="both", expand=True, pady=(14, 0))

        head = tk.Frame(list_card, bg=CARD)
        head.pack(fill="x", padx=14, pady=(12, 6))
        self.count_label = tk.Label(head, text="No files selected", bg=CARD, fg=TEXT, font=("Segoe UI", 10, "bold"))
        self.count_label.pack(side="left")

        list_wrap = tk.Frame(list_card, bg=CARD)
        list_wrap.pack(fill="both", expand=True, padx=14, pady=(0, 12))
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

    # -- intro / help dialog ------------------------------------------------
    @staticmethod
    def _load_config() -> dict:
        try:
            with open(_CONFIG_PATH, "r", encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return {}

    @staticmethod
    def _save_config(cfg: dict) -> None:
        try:
            with open(_CONFIG_PATH, "w", encoding="utf-8") as fh:
                json.dump(cfg, fh)
        except Exception:
            pass

    def _intro_disabled(self) -> bool:
        return bool(self._load_config().get("hide_intro", False))

    def _show_intro_dialog(self) -> None:
        """A friendly, plain-language explanation of how the app works."""
        caps = capabilities()

        def status(ok: bool) -> str:
            return "available" if ok else "not installed"

        win = tk.Toplevel(self.root)
        win.title("How this app works")
        win.configure(bg=CARD)
        win.resizable(False, False)
        win.transient(self.root)

        # Header strip.
        head = tk.Frame(win, bg=HEADER)
        head.pack(fill="x")
        tk.Label(
            head,
            text="\U0001F512  How this app works",
            bg=HEADER,
            fg="#ffffff",
            font=("Segoe UI", 15, "bold"),
        ).pack(anchor="w", padx=22, pady=(16, 4))
        tk.Label(
            head,
            text="A quick 30-second guide \u2014 in plain English.",
            bg=HEADER,
            fg=HEADER_SUB,
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=22, pady=(0, 16))

        body = tk.Frame(win, bg=CARD)
        body.pack(fill="both", expand=True, padx=22, pady=(16, 6))

        sections = [
            ("\U0001F4C4  What is metadata?",
             "Hidden information saved inside your files \u2014 things like the "
             "author's name, your company, when the file was created or edited, "
             "comments and tracked-change history. You can't see it in the page, "
             "but anyone who receives the file can."),
            ("\u2702\ufe0f  What this app does",
             "It removes that hidden information and saves a clean copy. Your "
             "original file is never changed, and the visible content and "
             "formatting stay exactly the same."),
            ("\u2699\ufe0f  How it does it",
             "\u2022 Modern files (.docx, .pptx, .xlsx): the hidden metadata "
             "parts are edited out directly.\n"
             "\u2022 Older files (.doc, .ppt, .xls): the hidden property "
             "sections are wiped out in place.\n"
             "\u2022 PDFs (.pdf): the document info and hidden metadata are "
             "removed.\n"
             "Every file keeps its original format \u2014 and it's all instant, "
             "with no other programs involved."),
            ("\U0001F5A5\ufe0f  What's ready on this computer",
             f"\u2022 Modern Office (.docx/.pptx/.xlsx): always available\n"
             f"\u2022 Legacy Office (.doc/.ppt/.xls): {status(caps['legacy'])} "
             f"(needs the olefile package)\n"
             f"\u2022 PDF (.pdf): {status(caps['pdf'])} (needs the pikepdf package)"),
            ("\u2705  How to use it",
             "1. Add files (drag & drop or the \u201cAdd files\u201d button).\n"
             "2. Optionally pick an output folder.\n"
             "3. Click \u201cRemove metadata\u201d. Clean copies are saved with a "
             "\u201c_clean\u201d name next to the originals."),
        ]
        for title, text in sections:
            tk.Label(
                body, text=title, bg=CARD, fg=TEXT,
                font=("Segoe UI", 11, "bold"), justify="left", anchor="w",
            ).pack(fill="x", pady=(8, 2))
            tk.Label(
                body, text=text, bg=CARD, fg=MUTED,
                font=("Segoe UI", 10), justify="left", anchor="w",
                wraplength=560,
            ).pack(fill="x")

        # Footer with "don't show again" + close button.
        footer = tk.Frame(win, bg=CARD)
        footer.pack(fill="x", padx=22, pady=(10, 18))

        hide_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            footer,
            text="Don't show this again",
            variable=hide_var,
            bg=CARD, fg=MUTED, font=("Segoe UI", 9),
            activebackground=CARD, selectcolor=CARD,
            highlightthickness=0, borderwidth=0,
        ).pack(side="left")

        def close() -> None:
            if hide_var.get():
                cfg = self._load_config()
                cfg["hide_intro"] = True
                self._save_config(cfg)
            win.destroy()

        ttk.Button(footer, text="Got it", style="Accent.TButton", command=close).pack(side="right")

        win.protocol("WM_DELETE_WINDOW", close)
        # Centre over the main window.
        win.update_idletasks()
        try:
            x = self.root.winfo_x() + (self.root.winfo_width() - win.winfo_width()) // 2
            y = self.root.winfo_y() + (self.root.winfo_height() - win.winfo_height()) // 2
            win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        except Exception:
            pass
        win.grab_set()

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
            title="Select documents",
            filetypes=[
                ("All supported", "*.docx *.pptx *.xlsx *.doc *.ppt *.xls *.pdf"),
                ("Modern Office", "*.docx *.pptx *.xlsx"),
                ("Legacy Office", "*.doc *.ppt *.xls"),
                ("PDF", "*.pdf"),
                ("Word", "*.docx *.doc"),
                ("PowerPoint", "*.pptx *.ppt"),
                ("Excel", "*.xlsx *.xls"),
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
        dependency_errors = set()
        for i, path in enumerate(files, start=1):
            result = clean_document(path, output_dir=out_dir)
            if result.success:
                ok += 1
                self._queue.put(("log", f"\u2713 {os.path.basename(path)}  \u2192  {os.path.basename(result.output_path)}", "ok"))
            else:
                self._queue.put(("log", f"\u2717 {os.path.basename(path)}  \u2014  {result.error}", "err"))
                # Detect missing dependency errors and collect them.
                if result.error and "require" in result.error and "package" in result.error:
                    dependency_errors.add(result.error)
            self._queue.put(("progress", (i / total) * 100, None))
        # Show dependency error dialog(s) if any were encountered.
        for err in dependency_errors:
            self._queue.put(("dependency_error", err, None))
        self._queue.put(("done", (ok, total), None))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload, tag = self._queue.get_nowait()
                if kind == "log":
                    self._log(payload, tag or "info")
                elif kind == "progress":
                    self.progress["value"] = payload
                elif kind == "dependency_error":
                    messagebox.showerror("Missing Dependency", payload)
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
