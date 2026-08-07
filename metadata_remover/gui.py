"""Tkinter desktop GUI for doc-metadata-remover.

A polished, dark "privacy vault" interface built entirely on top of
:mod:`tkinter` — no extra runtime dependencies.  The look and feel is provided
by two pure-presentation helper modules:

    * :mod:`metadata_remover.theme`       — colours, spacing, radii, fonts.
    * :mod:`metadata_remover.components`   — custom-drawn widgets (rounded
      cards, canvas buttons, chips, progress bar, vector icons).

Behaviour is unchanged from earlier versions:

    * Drag-and-drop files onto the window (optional ``tkinterdnd2`` package;
      the app degrades gracefully without it).
    * "Choose files" picker and per-file / clear-all removal.
    * Batch processing on a background thread with a live progress bar and log.
    * Optional output-folder selection (defaults to each file's own folder).

The heavy lifting still happens in :mod:`metadata_remover.core`; this module
only wires it up so the interface stays responsive.
"""

from __future__ import annotations

import json
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import List, Optional

from .core import (
    SUPPORTED_EXTENSIONS,
    capabilities,
    clean_document,
)
from .theme import (
    Color,
    Radius,
    Space,
    font,
    font_body,
    font_body_strong,
    font_caption,
    font_heading,
    font_subtitle,
    font_title,
    mono,
    type_for_extension,
)
from .components import (
    CanvasButton,
    Chip,
    ProgressBar,
    RoundedCard,
    Tooltip,
    draw_check,
    draw_document,
    draw_folder,
    draw_lock,
    draw_round_rect,
    draw_shield,
    draw_type_chip,
    draw_x,
)

# Small config file used to remember the "don't show intro again" choice.
_CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".doc_metadata_remover.json")

# Optional drag-and-drop support.
try:  # pragma: no cover - depends on optional dependency
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _DND_AVAILABLE = True
except Exception:  # pragma: no cover
    _DND_AVAILABLE = False


_FORMATS_LINE = "DOCX, PPTX, XLSX, DOC, PPT, XLS and PDF"


# ---------------------------------------------------------------------------
# Icon adapters (match the ``(canvas, cx, cy, size, color)`` button contract)
# ---------------------------------------------------------------------------
def _icon_shield(canvas, cx, cy, size, color):
    draw_shield(canvas, cx, cy, size, color, check_color=Color.ACCENT)


def _human_size(num: float) -> str:
    for unit in ("B", "KB", "MB", "GB"):
        if num < 1024 or unit == "GB":
            if unit == "B":
                return f"{int(num)} {unit}"
            return f"{num:.1f} {unit}"
        num /= 1024.0
    return f"{num:.1f} GB"


# ---------------------------------------------------------------------------
# Drop zone — a custom canvas with a dashed rounded outline and drag feedback
# ---------------------------------------------------------------------------
class DropZone(tk.Canvas):
    """The primary upload target: dashed card, document glyph, CTA button."""

    def __init__(self, master, on_choose, *, height=150, page_bg=Color.BG):
        super().__init__(master, height=height, bg=page_bg,
                         highlightthickness=0, bd=0)
        self.on_choose = on_choose
        self._hover = False
        self._drag = False
        self.button = CanvasButton(
            self, "Choose files", command=on_choose, kind="primary",
            icon=draw_folder, page_bg=Color.SURFACE, height=40,
        )
        self._btn_win = None
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<Button-1>", lambda e: self.on_choose())

    def set_drag(self, on: bool):
        self._drag = bool(on)
        self._draw()

    def _on_enter(self, _):
        self._hover = True
        self.configure(cursor="hand2")
        self._draw()

    def _on_leave(self, _):
        self._hover = False
        self.configure(cursor="")
        self._draw()

    def _draw(self):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= 1:
            return

        if self._drag:
            border = Color.ACCENT
            head_color = Color.ACCENT_HOVER
            line_color = Color.ACCENT_HOVER
        elif self._hover:
            border = Color.BORDER_HI
            head_color = Color.TEXT
            line_color = Color.ACCENT
        else:
            border = Color.BORDER
            head_color = Color.TEXT
            line_color = Color.ACCENT

        draw_round_rect(self, 2, 2, w - 2, h - 2, Radius.LG,
                        fill=Color.SURFACE, outline=border, width=2,
                        dash=(6, 5))

        cx = w / 2
        # Document glyph in a soft rounded tile.
        tile = 58
        ty = h * 0.30
        draw_round_rect(self, cx - tile / 2, ty - tile / 2,
                        cx + tile / 2, ty + tile / 2, Radius.MD,
                        fill=Color.SURFACE_ALT, outline="")
        draw_document(self, cx, ty, 34, Color.SURFACE_HI, line_color)

        headline = "Drop documents here" if _DND_AVAILABLE else "Add your documents"
        self.create_text(cx, h * 0.60, text=headline, fill=head_color,
                         font=font_heading())
        self.create_text(cx, h * 0.72, text=_FORMATS_LINE,
                         fill=Color.TEXT_MUTED, font=font_caption())

        # (Re)place the CTA button.
        by = h * 0.87
        if self._btn_win is None:
            self._btn_win = self.create_window(cx, by, window=self.button)
        else:
            self.create_window(cx, by, window=self.button)


# ---------------------------------------------------------------------------
# Scrollable custom file list (replaces the native Listbox)
# ---------------------------------------------------------------------------
class FileList(tk.Frame):
    """A scrollable column of file rows with an attractive empty state."""

    def __init__(self, master, on_remove, *, page_bg=Color.SURFACE):
        super().__init__(master, bg=page_bg)
        self.on_remove = on_remove
        self.page_bg = page_bg

        self.canvas = tk.Canvas(self, bg=page_bg, highlightthickness=0, bd=0,
                                height=96)
        self.scroll = tk.Scrollbar(
            self, orient="vertical", command=self.canvas.yview,
            width=10, borderwidth=0, highlightthickness=0,
            troughcolor=Color.SURFACE, bg=Color.BORDER_HI,
            activebackground=Color.TEXT_FAINT, relief="flat",
        )
        self.inner = tk.Frame(self.canvas, bg=page_bg)
        self._win = self.canvas.create_window(0, 0, anchor="nw",
                                               window=self.inner)
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scroll.pack(side="right", fill="y")

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfigure(self._win,
                                                             width=e.width))
        # Mouse-wheel scrolling (cross-platform).
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.canvas.bind_all(seq, self._on_wheel)

    def _on_inner_configure(self, _):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        # Hide the scrollbar when everything fits.
        if self.inner.winfo_reqheight() <= self.canvas.winfo_height():
            self.scroll.pack_forget()
        else:
            self.scroll.pack(side="right", fill="y")

    def _on_wheel(self, event):
        if not str(self.canvas.winfo_containing(event.x_root, event.y_root)
                   ).startswith(str(self.canvas)):
            return
        if getattr(event, "num", None) == 4:
            self.canvas.yview_scroll(-1, "units")
        elif getattr(event, "num", None) == 5:
            self.canvas.yview_scroll(1, "units")
        else:
            self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def render(self, files: List[str]):
        for child in self.inner.winfo_children():
            child.destroy()

        if not files:
            self._render_empty()
            return

        for path in files:
            self._render_row(path)

    def _render_empty(self):
        wrap = tk.Frame(self.inner, bg=self.page_bg)
        wrap.pack(fill="both", expand=True, pady=Space.SM)
        icon = tk.Canvas(wrap, width=34, height=34, bg=self.page_bg,
                         highlightthickness=0, bd=0)
        draw_document(icon, 17, 17, 26, Color.SURFACE_HI, Color.TEXT_FAINT)
        icon.pack()
        tk.Label(wrap, text="No files added yet", bg=self.page_bg,
                 fg=Color.TEXT_MUTED, font=font_body_strong()).pack(pady=(6, 2))
        tk.Label(wrap, text="Choose files or drag them onto the panel above.",
                 bg=self.page_bg, fg=Color.TEXT_FAINT,
                 font=font_caption()).pack()

    def _render_row(self, path: str):
        ext = os.path.splitext(path)[1].lower()
        label, key = type_for_extension(ext)

        row = tk.Frame(self.inner, bg=Color.SURFACE_ALT)
        row.pack(fill="x", pady=3, padx=1)

        chip = tk.Canvas(row, width=34, height=34, bg=Color.SURFACE_ALT,
                         highlightthickness=0, bd=0)
        draw_type_chip(chip, 17, 17, 28, label, key)
        chip.pack(side="left", padx=(10, 10), pady=8)

        remove = CanvasButton(
            row, "", command=lambda p=path: self.on_remove(p), kind="ghost",
            icon=draw_x, width=34, height=34, radius=Radius.SM,
            page_bg=Color.SURFACE_ALT, tooltip="Remove this file",
        )
        remove.pack(side="right", padx=(6, 10))

        text = tk.Frame(row, bg=Color.SURFACE_ALT)
        text.pack(side="left", fill="x", expand=True, pady=6)
        tk.Label(text, text=os.path.basename(path), bg=Color.SURFACE_ALT,
                 fg=Color.TEXT, font=font_body(), anchor="w",
                 justify="left").pack(fill="x", anchor="w")
        try:
            size = _human_size(os.path.getsize(path))
        except OSError:
            size = ""
        sub = f"{label} document" + (f"  \u00b7  {size}" if size else "")
        tk.Label(text, text=sub, bg=Color.SURFACE_ALT, fg=Color.TEXT_FAINT,
                 font=font_caption(), anchor="w").pack(fill="x", anchor="w")


class MetadataRemoverApp:
    """The main application window."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.files: List[str] = []
        self.output_dir: Optional[str] = None
        self._queue: "queue.Queue" = queue.Queue()
        self._worker: Optional[threading.Thread] = None
        self._details_open = False
        self._busy = False

        root.title("Doc Metadata Remover")
        root.geometry("860x780")
        root.minsize(720, 600)
        root.configure(bg=Color.BG)

        self._build_ui()
        self._sync_state()
        self._poll_queue()

        if not self._intro_disabled():
            self.root.after(250, self._show_intro_dialog)

    # -- UI layout ----------------------------------------------------------
    def _build_ui(self) -> None:
        # === Sticky bottom action bar (packed first → always visible) ======
        action = tk.Frame(self.root, bg=Color.BG)
        action.pack(side="bottom", fill="x")
        divider = tk.Frame(action, bg=Color.BORDER, height=1)
        divider.pack(fill="x")
        action_inner = tk.Frame(action, bg=Color.BG)
        action_inner.pack(fill="x", padx=Space.XL, pady=(Space.MD, Space.LG))

        trust = tk.Frame(action_inner, bg=Color.BG)
        trust.pack(side="left", anchor="w")
        shield_icon = tk.Canvas(trust, width=18, height=20, bg=Color.BG,
                                highlightthickness=0, bd=0)
        draw_shield(shield_icon, 9, 10, 16, Color.SUCCESS_SOFT,
                    check_color=Color.SUCCESS)
        shield_icon.pack(side="left", padx=(0, 8))
        tk.Label(trust, text="Your originals are never modified.",
                 bg=Color.BG, fg=Color.TEXT_MUTED,
                 font=font_caption()).pack(side="left")

        self.run_btn = CanvasButton(
            action_inner, "Remove metadata", command=self.run, kind="primary",
            icon=_icon_shield, height=48, width=228, radius=Radius.MD,
            page_bg=Color.BG, font=font(12, "bold"),
        )
        self.run_btn.pack(side="right")

        # === Lower sticky group (output + activity, above the action bar) ==
        # Packed to the bottom so the progress/status and output controls are
        # never pushed off-screen; only the file list flexes/scrolls.
        lower = tk.Frame(self.root, bg=Color.BG)
        lower.pack(side="bottom", fill="x", padx=Space.XL)
        self._build_output_card(lower)
        self._build_activity_card(lower)

        # === Main content area (header, drop zone, flexible file list) =====
        main = tk.Frame(self.root, bg=Color.BG)
        main.pack(side="top", fill="both", expand=True)
        content = tk.Frame(main, bg=Color.BG)
        content.pack(fill="both", expand=True, padx=Space.XL)

        self._build_header(content)
        self._build_dropzone(content)
        self._build_files_card(content)

        # Keyboard shortcuts.
        self.root.bind("<Control-o>", lambda e: self.add_files())
        self.root.bind("<Control-O>", lambda e: self.add_files())
        self.root.bind("<Delete>", lambda e: self.remove_selected())
        self.root.bind("<BackSpace>", lambda e: self.remove_selected())

    def _build_header(self, parent) -> None:
        card = RoundedCard(parent, expand=False, pad=Space.LG,
                           surface=Color.SURFACE, page_bg=Color.BG)
        card.pack(fill="x", pady=(Space.LG, Space.MD))
        body = card.body

        badge = tk.Canvas(body, width=54, height=54, bg=Color.SURFACE,
                          highlightthickness=0, bd=0)
        draw_round_rect(badge, 3, 3, 51, 51, Radius.MD,
                        fill=Color.ACCENT_SOFT, outline=Color.ACCENT)
        draw_lock(badge, 27, 26, 26, Color.ACCENT_HOVER, accent=Color.ACCENT_GLOW)
        badge.pack(side="left", padx=(2, Space.LG))

        titles = tk.Frame(body, bg=Color.SURFACE)
        titles.pack(side="left", fill="x", expand=True)
        tk.Label(titles, text="Doc Metadata Remover", bg=Color.SURFACE,
                 fg=Color.TEXT, font=font_title(), anchor="w").pack(anchor="w")
        tk.Label(titles,
                 text="Clean document identity data. Preserve everything else.",
                 bg=Color.SURFACE, fg=Color.TEXT_MUTED, font=font_subtitle(),
                 anchor="w").pack(anchor="w", pady=(3, 0))

        self.status_chip = Chip(
            body, "Content-safe", fg=Color.SUCCESS, bg=Color.SUCCESS_SOFT,
            page_bg=Color.SURFACE, icon=draw_check, icon_color=Color.SUCCESS,
            height=26, font=font_caption(),
        )
        self.status_chip.pack(side="right", anchor="n", padx=(Space.MD, 2))

    def _build_dropzone(self, parent) -> None:
        self.drop_zone = DropZone(parent, self.add_files, page_bg=Color.BG)
        self.drop_zone.pack(fill="x", pady=(0, Space.MD))

        if _DND_AVAILABLE:
            self.drop_zone.drop_target_register(DND_FILES)
            self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)
            self.drop_zone.dnd_bind("<<DropEnter>>",
                                    lambda e: self.drop_zone.set_drag(True))
            self.drop_zone.dnd_bind("<<DropLeave>>",
                                    lambda e: self.drop_zone.set_drag(False))

    def _build_files_card(self, parent) -> None:
        card = RoundedCard(parent, expand=True, pad=Space.LG,
                           surface=Color.SURFACE, page_bg=Color.BG)
        card.pack(fill="both", expand=True, pady=(0, Space.MD))
        body = card.body

        head = tk.Frame(body, bg=Color.SURFACE)
        head.pack(fill="x", pady=(0, Space.SM))
        tk.Label(head, text="Selected files", bg=Color.SURFACE, fg=Color.TEXT,
                 font=font_heading()).pack(side="left")
        self.count_chip = Chip(head, "0 files", fg=Color.TEXT_MUTED,
                               bg=Color.SURFACE_HI, page_bg=Color.SURFACE,
                               height=24, font=font_caption())
        self.count_chip.pack(side="left", padx=Space.SM)

        self.clear_btn = CanvasButton(
            head, "Clear all", command=self.clear_files, kind="ghost",
            icon=draw_x, height=30, radius=Radius.SM, page_bg=Color.SURFACE,
            font=font_caption(),
        )
        self.clear_btn.pack(side="right")

        self.file_list = FileList(body, self.remove_path, page_bg=Color.SURFACE)
        self.file_list.pack(fill="both", expand=True)

    def _build_output_card(self, parent) -> None:
        card = RoundedCard(parent, expand=False, pad=Space.MD,
                           surface=Color.SURFACE, page_bg=Color.BG)
        card.pack(fill="x", pady=(0, Space.MD))
        body = card.body

        top = tk.Frame(body, bg=Color.SURFACE)
        top.pack(fill="x")

        folder_icon = tk.Canvas(top, width=30, height=30, bg=Color.SURFACE,
                                highlightthickness=0, bd=0)
        draw_folder(folder_icon, 15, 16, 22, Color.TEXT_MUTED)
        folder_icon.pack(side="left", padx=(0, Space.MD))

        labels = tk.Frame(top, bg=Color.SURFACE)
        labels.pack(side="left", fill="x", expand=True)
        tk.Label(labels, text="Save cleaned copies to", bg=Color.SURFACE,
                 fg=Color.TEXT_MUTED, font=font_caption(),
                 anchor="w").pack(anchor="w")
        self.output_var = tk.StringVar(value="Same folder as originals")
        tk.Label(labels, textvariable=self.output_var, bg=Color.SURFACE,
                 fg=Color.TEXT, font=font_body_strong(),
                 anchor="w").pack(anchor="w")

        self.output_btn = CanvasButton(
            top, "Choose folder", command=self.choose_output, kind="ghost",
            icon=draw_folder, height=34, radius=Radius.SM,
            page_bg=Color.SURFACE, font=font_caption(),
        )
        self.output_btn.pack(side="right")
        self.reset_output_btn = CanvasButton(
            top, "", command=self.reset_output, kind="ghost", icon=draw_x,
            width=34, height=34, radius=Radius.SM, page_bg=Color.SURFACE,
            tooltip="Reset to the originals' folder",
        )
        # Shown only when a custom folder is chosen.

        tk.Label(body,
                 text="Copies are saved with a \u201c_clean\u201d name \u2014 "
                      "your original files are never overwritten.",
                 bg=Color.SURFACE, fg=Color.TEXT_FAINT, font=font_caption(),
                 anchor="w", justify="left").pack(anchor="w", pady=(Space.SM, 0))

    def _build_activity_card(self, parent) -> None:
        card = RoundedCard(parent, expand=False, pad=Space.MD,
                           surface=Color.SURFACE, page_bg=Color.BG)
        card.pack(fill="x", pady=(0, Space.LG))
        body = card.body

        row = tk.Frame(body, bg=Color.SURFACE)
        row.pack(fill="x")
        self.status_icon = tk.Canvas(row, width=20, height=20, bg=Color.SURFACE,
                                     highlightthickness=0, bd=0)
        self.status_icon.pack(side="left", padx=(0, Space.SM))
        self.status_label = tk.Label(row, text="Ready", bg=Color.SURFACE,
                                     fg=Color.TEXT_MUTED, font=font_body())
        self.status_label.pack(side="left")

        self.details_btn = CanvasButton(
            row, "View details", command=self.toggle_details, kind="ghost",
            height=28, radius=Radius.SM, page_bg=Color.SURFACE,
            font=font_caption(),
        )
        self.details_btn.pack(side="right")

        self.progress = ProgressBar(body, height=10, page_bg=Color.SURFACE,
                                    track=Color.SURFACE_ALT, fill=Color.ACCENT)
        self.progress.pack(fill="x", pady=(Space.MD, 0))

        # Collapsible log (hidden by default).
        self.details_frame = tk.Frame(body, bg=Color.SURFACE)
        self.log = tk.Text(
            self.details_frame, height=6, bg=Color.ELEVATED, fg=Color.TEXT,
            insertbackground=Color.TEXT, font=mono(9), borderwidth=0,
            highlightthickness=1, highlightbackground=Color.BORDER,
            highlightcolor=Color.BORDER, state="disabled", wrap="word",
            padx=12, pady=10,
        )
        self.log.pack(fill="both", expand=True, pady=(Space.MD, 0))
        self.log.tag_configure("ok", foreground=Color.SUCCESS)
        self.log.tag_configure("err", foreground=Color.ERROR)
        self.log.tag_configure("info", foreground=Color.ACCENT_GLOW)

        self._draw_status_icon("idle")

    # -- status icon --------------------------------------------------------
    def _draw_status_icon(self, state: str) -> None:
        c = self.status_icon
        c.delete("all")
        if state == "idle":
            c.create_oval(6, 6, 14, 14, fill=Color.TEXT_FAINT, outline="")
        elif state == "running":
            c.create_oval(5, 5, 15, 15, outline=Color.ACCENT, width=2)
            c.create_oval(8, 8, 12, 12, fill=Color.ACCENT, outline="")
        elif state == "success":
            c.create_oval(2, 2, 18, 18, fill=Color.SUCCESS_SOFT, outline="")
            draw_check(c, 10, 10, 11, Color.SUCCESS)
        elif state == "error":
            c.create_oval(2, 2, 18, 18, fill=Color.ERROR_SOFT, outline="")
            draw_x(c, 10, 10, 9, Color.ERROR)

    def toggle_details(self) -> None:
        self._details_open = not self._details_open
        if self._details_open:
            self.details_frame.pack(fill="both", expand=True)
            self.details_btn.set_text("Hide details")
        else:
            self.details_frame.pack_forget()
            self.details_btn.set_text("View details")

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
        win.configure(bg=Color.SURFACE)
        win.resizable(False, False)
        win.transient(self.root)

        # Header strip.
        head = tk.Frame(win, bg=Color.ELEVATED)
        head.pack(fill="x")
        head_inner = tk.Frame(head, bg=Color.ELEVATED)
        head_inner.pack(fill="x", padx=24, pady=18)
        badge = tk.Canvas(head_inner, width=42, height=42, bg=Color.ELEVATED,
                          highlightthickness=0, bd=0)
        draw_round_rect(badge, 2, 2, 40, 40, Radius.MD, fill=Color.ACCENT_SOFT,
                        outline=Color.ACCENT)
        draw_lock(badge, 21, 20, 20, Color.ACCENT_HOVER, accent=Color.ACCENT_GLOW)
        badge.pack(side="left", padx=(0, 14))
        head_titles = tk.Frame(head_inner, bg=Color.ELEVATED)
        head_titles.pack(side="left", anchor="w")
        tk.Label(head_titles, text="How this app works", bg=Color.ELEVATED,
                 fg=Color.TEXT, font=font(15, "bold")).pack(anchor="w")
        tk.Label(head_titles, text="A quick 30-second guide \u2014 in plain English.",
                 bg=Color.ELEVATED, fg=Color.TEXT_MUTED,
                 font=font_subtitle()).pack(anchor="w", pady=(2, 0))

        body = tk.Frame(win, bg=Color.SURFACE)
        body.pack(fill="both", expand=True, padx=24, pady=(18, 6))

        sections = [
            ("What is metadata?",
             "Hidden information saved inside your files \u2014 things like the "
             "author's name, your company, when the file was created or edited, "
             "and the names attached to comments and tracked changes. Anyone who "
             "receives the file can dig it out."),
            ("What this app does",
             "It removes that hidden information and saves a clean copy. Your "
             "original file is never changed, and the visible content and "
             "formatting stay exactly the same.\n"
             "Comments and tracked changes are kept \u2014 but the names, "
             "initials and dates attached to them are wiped, so the feedback "
             "stays while the person behind it becomes anonymous."),
            ("How it does it",
             "\u2022 Modern files (.docx, .pptx, .xlsx): the hidden metadata is "
             "edited out directly, and names on comments / tracked changes are "
             "blanked.\n"
             "\u2022 Older files (.doc, .ppt, .xls): the hidden document "
             "properties (author, company, dates, etc.) are wiped in place. "
             "Note: for these older files, names inside comments or tracked "
             "changes in the body are NOT removed \u2014 save them as the modern "
             "format (.docx etc.) first if you need that.\n"
             "\u2022 PDFs (.pdf): the document info and hidden metadata are "
             "removed.\n"
             "Every file keeps its original format \u2014 and it's all instant, "
             "with no other programs involved."),
            ("What's ready on this computer",
             f"\u2022 Modern Office (.docx/.pptx/.xlsx): always available\n"
             f"\u2022 Legacy Office (.doc/.ppt/.xls): {status(caps['legacy'])} "
             f"(needs the olefile package)\n"
             f"\u2022 PDF (.pdf): {status(caps['pdf'])} (needs the pikepdf package)"),
            ("How to use it",
             "1. Add files (drag & drop or the \u201cChoose files\u201d button).\n"
             "2. Optionally pick an output folder.\n"
             "3. Click \u201cRemove metadata\u201d. Clean copies are saved with a "
             "\u201c_clean\u201d name next to the originals."),
        ]
        for title, text in sections:
            tk.Label(body, text=title, bg=Color.SURFACE, fg=Color.TEXT,
                     font=font_body_strong(), justify="left",
                     anchor="w").pack(fill="x", pady=(8, 2))
            tk.Label(body, text=text, bg=Color.SURFACE, fg=Color.TEXT_MUTED,
                     font=font_body(), justify="left", anchor="w",
                     wraplength=560).pack(fill="x")

        footer = tk.Frame(win, bg=Color.SURFACE)
        footer.pack(fill="x", padx=24, pady=(10, 20))

        hide_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            footer, text="Don't show this again", variable=hide_var,
            bg=Color.SURFACE, fg=Color.TEXT_MUTED, font=font_caption(),
            activebackground=Color.SURFACE, activeforeground=Color.TEXT,
            selectcolor=Color.SURFACE_ALT, highlightthickness=0, borderwidth=0,
        ).pack(side="left")

        def close() -> None:
            if hide_var.get():
                cfg = self._load_config()
                cfg["hide_intro"] = True
                self._save_config(cfg)
            win.destroy()

        CanvasButton(footer, "Got it", command=close, kind="primary",
                     height=40, page_bg=Color.SURFACE).pack(side="right")

        win.protocol("WM_DELETE_WINDOW", close)
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
                    added += 1
        self._sync_state()
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
        self._sync_state()

    def remove_path(self, path: str) -> None:
        if path in self.files:
            self.files.remove(path)
            self._sync_state()

    def remove_selected(self) -> None:
        """Keyboard fallback: remove the most recently added file."""
        if self.files:
            self.files.pop()
            self._sync_state()

    def _on_drop(self, event) -> None:  # pragma: no cover - GUI event
        self.drop_zone.set_drag(False)
        paths = self.root.tk.splitlist(event.data)
        self._add_paths(list(paths))

    def _sync_state(self) -> None:
        n = len(self.files)
        self.count_chip.set_text("1 file" if n == 1 else f"{n} files")
        self.file_list.render(self.files)
        if not self._busy:
            self.run_btn.set_enabled(n > 0)
        self.clear_btn.set_enabled(n > 0)

    def choose_output(self) -> None:
        d = filedialog.askdirectory(title="Choose output folder")
        if d:
            self.output_dir = d
            folder_name = os.path.basename(d) or d
            self.output_var.set(folder_name)
            self.reset_output_btn.pack(side="right", padx=(6, 0))
        # Cancelling leaves the current choice untouched.

    def reset_output(self) -> None:
        self.output_dir = None
        self.output_var.set("Same folder as originals")
        self.reset_output_btn.pack_forget()

    # -- processing ---------------------------------------------------------
    def run(self) -> None:
        if self._worker and self._worker.is_alive():
            return
        if not self.files:
            messagebox.showinfo("No files", "Please add one or more files first.")
            return
        self._busy = True
        self.run_btn.set_enabled(False)
        self.run_btn.set_text("Cleaning\u2026")
        self.progress.set_value(0)
        self._draw_status_icon("running")
        self._set_status("Preparing\u2026", Color.TEXT_MUTED)
        files = list(self.files)
        out_dir = self.output_dir
        self._worker = threading.Thread(target=self._work, args=(files, out_dir),
                                        daemon=True)
        self._worker.start()

    def _work(self, files: List[str], out_dir: Optional[str]) -> None:
        total = len(files)
        ok = 0
        dependency_errors = set()
        for i, path in enumerate(files, start=1):
            self._queue.put(("status", f"Cleaning {i} of {total} files\u2026", None))
            result = clean_document(path, output_dir=out_dir)
            if result.success:
                ok += 1
                self._queue.put(("log", f"\u2713 {os.path.basename(path)}  \u2192  {os.path.basename(result.output_path)}", "ok"))
            else:
                self._queue.put(("log", f"\u2717 {os.path.basename(path)}  \u2014  {result.error}", "err"))
                if result.error and "require" in result.error and "package" in result.error:
                    dependency_errors.add(result.error)
            self._queue.put(("progress", (i / total) * 100, None))
        for err in dependency_errors:
            self._queue.put(("dependency_error", err, None))
        self._queue.put(("done", (ok, total), None))

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, payload, tag = self._queue.get_nowait()
                if kind == "log":
                    self._log(payload, tag or "info")
                elif kind == "status":
                    self._set_status(payload, Color.TEXT)
                elif kind == "progress":
                    self.progress.set_value(payload)
                elif kind == "dependency_error":
                    messagebox.showerror("Missing Dependency", payload)
                elif kind == "done":
                    ok, total = payload
                    self._busy = False
                    self.run_btn.set_text("Remove metadata")
                    self.run_btn.set_enabled(len(self.files) > 0)
                    failed = total - ok
                    if ok and not failed:
                        self._draw_status_icon("success")
                        word = "file" if ok == 1 else "files"
                        self._set_status(f"{ok} {word} cleaned successfully",
                                         Color.SUCCESS)
                    elif ok:
                        self._draw_status_icon("error")
                        self._set_status(
                            f"{ok} of {total} cleaned \u00b7 {failed} failed",
                            Color.WARNING)
                    else:
                        self._draw_status_icon("error")
                        self._set_status("No files could be cleaned",
                                         Color.ERROR)
                    self._log(f"Finished: {ok} of {total} file(s) cleaned.", "info")
                    if ok:
                        messagebox.showinfo("Complete", f"Cleaned {ok} of {total} file(s).")
                    elif failed:
                        self._open_details()
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    # -- helpers ------------------------------------------------------------
    def _open_details(self) -> None:
        if not self._details_open:
            self.toggle_details()

    def _log(self, msg: str, tag: str = "info") -> None:
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

    def _set_status(self, text: str, color: str = Color.TEXT_MUTED) -> None:
        self.status_label.config(text=text, fg=color)


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
