"""Reusable, custom-drawn Tkinter widgets for a premium dark UI.

These widgets are pure presentation — they hold no application logic.  They are
built on ``tk.Canvas`` so we can render rounded corners, soft borders, hover /
pressed / focus states, and small vector icons that the stock ``tk``/``ttk``
controls cannot produce.

Nothing here imports :mod:`metadata_remover.core`; the GUI wires callbacks in.
"""

from __future__ import annotations

import tkinter as tk
from typing import Callable, Optional

from .theme import Color, Radius, Space, font_body, font_button, font_caption, TYPE_COLORS


# ---------------------------------------------------------------------------
# Low-level drawing helpers
# ---------------------------------------------------------------------------
def round_rect_points(x1, y1, x2, y2, r):
    """Return a point list approximating a rounded rectangle for a polygon."""
    r = max(0, min(r, (x2 - x1) / 2, (y2 - y1) / 2))
    return [
        x1 + r, y1,
        x2 - r, y1,
        x2, y1,
        x2, y1 + r,
        x2, y2 - r,
        x2, y2,
        x2 - r, y2,
        x1 + r, y2,
        x1, y2,
        x1, y2 - r,
        x1, y1 + r,
        x1, y1,
    ]


def draw_round_rect(canvas: tk.Canvas, x1, y1, x2, y2, r, fill="", outline="",
                    width=1, dash=None, tags=""):
    """Draw a smooth rounded rectangle on ``canvas`` and return the item id."""
    pts = round_rect_points(x1, y1, x2, y2, r)
    kwargs = dict(smooth=True, splinesteps=36, fill=fill, outline=outline,
                  width=width, tags=tags)
    if dash:
        try:
            return canvas.create_polygon(pts, dash=dash, **kwargs)
        except tk.TclError:
            pass  # some Tk builds reject dashed polygons — fall back to solid
    return canvas.create_polygon(pts, **kwargs)


# ---------------------------------------------------------------------------
# Vector icon drawing (no emoji, no image files)
# ---------------------------------------------------------------------------
def draw_shield(canvas, cx, cy, size, fill, check_color=None, outline=""):
    """Draw a shield centred at (cx, cy). Optionally a check mark inside."""
    hw = size * 0.42
    top = cy - size * 0.5
    shoulder = cy - size * 0.18
    bottom = cy + size * 0.52
    pts = [
        cx, top,
        cx + hw, shoulder,
        cx + hw, cy + size * 0.12,
        cx, bottom,
        cx - hw, cy + size * 0.12,
        cx - hw, shoulder,
    ]
    canvas.create_polygon(pts, smooth=True, splinesteps=24, fill=fill,
                          outline=outline or fill)
    if check_color:
        s = size * 0.22
        canvas.create_line(cx - s, cy + size * 0.02,
                           cx - s * 0.2, cy + s * 0.9,
                           cx + s * 1.1, cy - s * 0.8,
                           fill=check_color, width=max(2, int(size * 0.09)),
                           capstyle="round", joinstyle="round")


def draw_lock(canvas, cx, cy, size, fill, accent=None):
    """Draw a padlock centred at (cx, cy)."""
    bw = size * 0.62
    bh = size * 0.52
    bx1, by1 = cx - bw / 2, cy - bh * 0.15
    bx2, by2 = cx + bw / 2, cy - bh * 0.15 + bh
    # Shackle.
    r = size * 0.26
    canvas.create_arc(cx - r, by1 - r * 1.5, cx + r, by1 + r * 0.5,
                      start=0, extent=180, style="arc",
                      outline=accent or fill, width=max(2, int(size * 0.1)))
    # Body.
    draw_round_rect(canvas, bx1, by1, bx2, by2, size * 0.14, fill=fill)
    # Keyhole.
    kh = accent or Color.BG
    canvas.create_oval(cx - size * 0.07, cy + size * 0.02,
                       cx + size * 0.07, cy + size * 0.16, fill=kh, outline="")


def draw_document(canvas, cx, cy, size, fill, line_color):
    """Draw a document/page glyph centred at (cx, cy)."""
    w = size * 0.62
    h = size * 0.8
    x1, y1 = cx - w / 2, cy - h / 2
    x2, y2 = cx + w / 2, cy + h / 2
    fold = size * 0.22
    pts = [x1, y1, x2 - fold, y1, x2, y1 + fold, x2, y2, x1, y2]
    canvas.create_polygon(pts, fill=fill, outline=line_color, width=1,
                          joinstyle="round")
    # Folded corner.
    canvas.create_line(x2 - fold, y1, x2 - fold, y1 + fold, x2, y1 + fold,
                       fill=line_color, width=1)
    # Text lines.
    for i in range(3):
        ly = y1 + h * (0.42 + i * 0.18)
        canvas.create_line(x1 + w * 0.2, ly, x2 - w * 0.2, ly,
                           fill=line_color, width=max(1, int(size * 0.045)))


def draw_folder(canvas, cx, cy, size, fill):
    w = size * 0.82
    h = size * 0.6
    x1, y1 = cx - w / 2, cy - h / 2
    x2, y2 = cx + w / 2, cy + h / 2
    # Tab.
    draw_round_rect(canvas, x1, y1 - h * 0.18, x1 + w * 0.45, y1 + h * 0.2,
                    size * 0.08, fill=fill)
    draw_round_rect(canvas, x1, y1, x2, y2, size * 0.1, fill=fill)


def draw_check(canvas, cx, cy, size, color, width=None):
    s = size * 0.5
    canvas.create_line(cx - s, cy,
                       cx - s * 0.15, cy + s * 0.75,
                       cx + s, cy - s * 0.7,
                       fill=color, width=width or max(2, int(size * 0.16)),
                       capstyle="round", joinstyle="round")


def draw_x(canvas, cx, cy, size, color, width=None):
    s = size * 0.5
    w = width or max(2, int(size * 0.16))
    canvas.create_line(cx - s, cy - s, cx + s, cy + s, fill=color, width=w,
                       capstyle="round")
    canvas.create_line(cx - s, cy + s, cx + s, cy - s, fill=color, width=w,
                       capstyle="round")


# ---------------------------------------------------------------------------
# Tooltip
# ---------------------------------------------------------------------------
class Tooltip:
    """A lightweight hover tooltip for compact / icon-only controls."""

    def __init__(self, widget, text: str, delay: int = 450):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._after = None
        self._tip = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _schedule(self, _=None):
        self._cancel()
        self._after = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._after:
            self.widget.after_cancel(self._after)
            self._after = None

    def _show(self):
        if self._tip or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        except tk.TclError:
            return
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        self._tip.configure(bg=Color.BORDER_HI)
        lbl = tk.Label(self._tip, text=self.text, bg=Color.SURFACE_HI,
                       fg=Color.TEXT, font=font_caption(), padx=10, pady=5,
                       justify="left")
        lbl.pack(padx=1, pady=1)
        self._tip.wm_geometry(f"+{x}+{y}")

    def _hide(self, _=None):
        self._cancel()
        if self._tip:
            self._tip.destroy()
            self._tip = None


# ---------------------------------------------------------------------------
# RoundedCard — a surface with rounded corners and a subtle border
# ---------------------------------------------------------------------------
class RoundedCard(tk.Frame):
    """A card with rounded corners drawn on a background canvas.

    Content goes into ``card.body`` (a normal Frame with the surface colour).
    ``expand=True`` makes the card fill its parent (height driven by layout);
    ``expand=False`` makes the card's height follow its content.
    """

    def __init__(self, master, *, radius=Radius.LG, surface=Color.SURFACE,
                 border=Color.BORDER, page_bg=Color.BG, pad=Space.LG,
                 expand=False, **kw):
        super().__init__(master, bg=page_bg, **kw)
        self.radius = radius
        self.surface = surface
        self.border = border
        self.page_bg = page_bg
        self.pad = pad
        self._expand = expand
        self._cw = 1
        self._ch = 1

        self.canvas = tk.Canvas(self, bg=page_bg, highlightthickness=0, bd=0)
        self.canvas.pack(fill="both", expand=True)
        self.body = tk.Frame(self.canvas, bg=surface)
        self._win = self.canvas.create_window(pad, pad, anchor="nw",
                                               window=self.body)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        if not expand:
            self.body.bind("<Configure>", self._on_body_configure)

    def set_border(self, color):
        if color != self.border:
            self.border = color
            self._redraw()

    def set_surface(self, color):
        if color != self.surface:
            self.surface = color
            self.body.configure(bg=color)
            self._redraw()

    def _on_canvas_configure(self, event):
        self._cw = event.width
        self._ch = event.height
        self.canvas.itemconfigure(self._win, width=event.width - 2 * self.pad)
        if self._expand:
            self.canvas.itemconfigure(self._win, height=event.height - 2 * self.pad)
        self._redraw()

    def _on_body_configure(self, event):
        need = event.height + 2 * self.pad
        if abs(need - self.canvas.winfo_reqheight()) > 1:
            self.canvas.configure(height=need)
        self._redraw()

    def _redraw(self):
        c = self.canvas
        c.delete("card_bg")
        w = self._cw
        h = self.canvas.winfo_height() if self._expand else (
            self.body.winfo_reqheight() + 2 * self.pad)
        draw_round_rect(c, 1, 1, w - 1, h - 1, self.radius,
                        fill=self.surface, outline=self.border, width=1,
                        tags="card_bg")
        c.tag_lower("card_bg")


# ---------------------------------------------------------------------------
# CanvasButton — a fully custom button with hover/pressed/focus/disabled states
# ---------------------------------------------------------------------------
class CanvasButton(tk.Canvas):
    """A rounded button drawn on a canvas.

    ``kind`` selects the palette: ``"primary"``, ``"ghost"`` or ``"danger"``.
    ``icon`` is an optional callable ``(canvas, cx, cy, size, color)``.
    Fully keyboard accessible (Tab focus + Enter/Space) with a visible focus ring.
    """

    def __init__(self, master, text="", command: Optional[Callable] = None, *,
                 kind="primary", icon: Optional[Callable] = None,
                 width=None, height=40, radius=Radius.MD, page_bg=Color.BG,
                 tooltip: Optional[str] = None, font=None, **kw):
        self._pad_x = 20
        self.text = text
        self.command = command
        self.kind = kind
        self.icon = icon
        self.radius = radius
        self.page_bg = page_bg
        self._font = font or font_button()
        self._state = "normal"   # normal | hover | pressed
        self._enabled = True
        self._focused = False

        self._palette(kind)
        w = width or self._measure_width()
        super().__init__(master, width=w, height=height, bg=page_bg,
                         highlightthickness=0, bd=0, takefocus=1, **kw)
        self.bind("<Configure>", lambda e: self._draw())
        self.bind("<Enter>", self._on_enter)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<FocusIn>", self._on_focus_in)
        self.bind("<FocusOut>", self._on_focus_out)
        self.bind("<Return>", self._on_key_activate)
        self.bind("<space>", self._on_key_activate)
        if tooltip:
            Tooltip(self, tooltip)
        self._draw()

    # -- palette ---------------------------------------------------------
    def _palette(self, kind):
        if kind == "primary":
            self.c_fill = Color.ACCENT
            self.c_hover = Color.ACCENT_HOVER
            self.c_press = Color.ACCENT_PRESSED
            self.c_text = Color.TEXT_ON_ACCENT
            self.c_border = ""
            self.c_disabled = Color.ACCENT_DISABLED
        elif kind == "danger":
            self.c_fill = Color.SURFACE_ALT
            self.c_hover = Color.ERROR_SOFT
            self.c_press = Color.ERROR_SOFT
            self.c_text = Color.ERROR
            self.c_border = Color.BORDER
            self.c_disabled = Color.SURFACE_ALT
        else:  # ghost
            self.c_fill = Color.SURFACE_ALT
            self.c_hover = Color.SURFACE_HI
            self.c_press = Color.SURFACE
            self.c_text = Color.TEXT_MUTED
            self.c_border = Color.BORDER
            self.c_disabled = Color.SURFACE

    def _measure_width(self):
        import tkinter.font as tkfont
        try:
            f = tkfont.Font(font=self._font)
            tw = f.measure(self.text)
        except Exception:
            tw = len(self.text) * 8
        icon_space = 26 if self.icon else 0
        return int(tw + icon_space + self._pad_x * 2)

    # -- state -----------------------------------------------------------
    def set_enabled(self, enabled: bool):
        self._enabled = enabled
        self.configure(takefocus=1 if enabled else 0)
        self._state = "normal"
        self._draw()

    def set_text(self, text: str):
        self.text = text
        self._draw()

    def _on_enter(self, _):
        if self._enabled:
            self._state = "hover"
            self.configure(cursor="hand2")
            self._draw()

    def _on_leave(self, _):
        self._state = "normal"
        self.configure(cursor="")
        self._draw()

    def _on_press(self, _):
        if self._enabled:
            self._state = "pressed"
            self._draw()

    def _on_release(self, event):
        if not self._enabled:
            return
        was_inside = 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height()
        self._state = "hover" if was_inside else "normal"
        self._draw()
        if was_inside and self.command:
            self.command()

    def _on_focus_in(self, _):
        self._focused = True
        self._draw()

    def _on_focus_out(self, _):
        self._focused = False
        self._draw()

    def _on_key_activate(self, _):
        if self._enabled and self.command:
            self.command()
        return "break"

    # -- drawing ---------------------------------------------------------
    def _draw(self):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= 1:
            w = int(self["width"])
        if h <= 1:
            h = int(self["height"])

        if not self._enabled:
            fill = self.c_disabled
            text_color = Color.TEXT_DISABLED
            border = self.c_border
        else:
            fill = {"normal": self.c_fill, "hover": self.c_hover,
                    "pressed": self.c_press}[self._state]
            text_color = self.c_text
            border = self.c_border

        # Focus ring (drawn slightly outside the fill).
        if self._focused and self._enabled:
            draw_round_rect(self, 1, 1, w - 1, h - 1, self.radius + 1,
                            fill="", outline=Color.ACCENT_GLOW, width=2)
            inset = 3
        else:
            inset = 1
        draw_round_rect(self, inset, inset, w - inset, h - inset, self.radius,
                        fill=fill, outline=border, width=1 if border else 0)

        cy = h / 2
        if self.icon:
            icon_size = min(18, h * 0.5)
            # Centre icon+text as a group.
            import tkinter.font as tkfont
            try:
                tw = tkfont.Font(font=self._font).measure(self.text)
            except Exception:
                tw = len(self.text) * 8
            gap = 8
            group_w = icon_size + gap + tw
            start = (w - group_w) / 2
            self.icon(self, start + icon_size / 2, cy, icon_size, text_color)
            self.create_text(start + icon_size + gap, cy, text=self.text,
                             fill=text_color, font=self._font, anchor="w")
        else:
            self.create_text(w / 2, cy, text=self.text, fill=text_color,
                             font=self._font, anchor="center")


# ---------------------------------------------------------------------------
# Chip / badge — a small rounded pill
# ---------------------------------------------------------------------------
class Chip(tk.Canvas):
    def __init__(self, master, text="", *, fg=Color.TEXT, bg=Color.SURFACE_HI,
                 page_bg=Color.SURFACE, icon: Optional[Callable] = None,
                 icon_color=None, font=None, height=24, **kw):
        self.text = text
        self.fg = fg
        self.bg = bg
        self.page_bg = page_bg
        self.icon = icon
        self.icon_color = icon_color or fg
        self._font = font or font_caption()
        super().__init__(master, height=height, bg=page_bg,
                         highlightthickness=0, bd=0, **kw)
        self.configure(width=self._measure_width(height))
        self.bind("<Configure>", lambda e: self._draw())
        self._draw()

    def _measure_width(self, height):
        import tkinter.font as tkfont
        try:
            tw = tkfont.Font(font=self._font).measure(self.text)
        except Exception:
            tw = len(self.text) * 7
        icon_space = (height * 0.6 + 6) if self.icon else 0
        return int(tw + icon_space + 22)

    def set_text(self, text):
        self.text = text
        self.configure(width=self._measure_width(int(self["height"])))
        self._draw()

    def set_colors(self, fg=None, bg=None, icon_color=None):
        if fg:
            self.fg = fg
        if bg:
            self.bg = bg
        if icon_color:
            self.icon_color = icon_color
        self._draw()

    def _draw(self):
        self.delete("all")
        w = self.winfo_width() or int(self["width"])
        h = self.winfo_height() or int(self["height"])
        draw_round_rect(self, 1, 1, w - 1, h - 1, h / 2, fill=self.bg, outline="")
        cy = h / 2
        x = 11
        if self.icon:
            isize = h * 0.5
            self.icon(self, x + isize / 2, cy, isize, self.icon_color)
            x += isize + 6
        self.create_text(x, cy, text=self.text, fill=self.fg, font=self._font,
                         anchor="w")


# ---------------------------------------------------------------------------
# Rounded progress bar with a subtle accent glow
# ---------------------------------------------------------------------------
class ProgressBar(tk.Canvas):
    def __init__(self, master, *, height=10, page_bg=Color.SURFACE,
                 track=Color.SURFACE_ALT, fill=Color.ACCENT, **kw):
        self._value = 0.0
        self.track = track
        self.fill = fill
        self.page_bg = page_bg
        super().__init__(master, height=height, bg=page_bg,
                         highlightthickness=0, bd=0, **kw)
        self.bind("<Configure>", lambda e: self._draw())
        self._draw()

    def set_value(self, value: float):
        self._value = max(0.0, min(100.0, value))
        self._draw()

    def _draw(self):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= 1:
            return
        r = h / 2
        draw_round_rect(self, 1, 1, w - 1, h - 1, r, fill=self.track, outline="")
        if self._value > 0:
            fw = max(h, (w - 2) * (self._value / 100.0))
            # Soft glow underlay.
            draw_round_rect(self, 1, 0, fw + 1, h, r, fill=Color.ACCENT_SOFT,
                            outline="")
            draw_round_rect(self, 1, 1, fw + 1, h - 1, r, fill=self.fill,
                            outline="")


# ---------------------------------------------------------------------------
# File-type chip icon (used inside file rows)
# ---------------------------------------------------------------------------
def draw_type_chip(canvas, cx, cy, size, label, color_key):
    """Draw a small rounded chip with a 2–3 letter file-type label."""
    color = TYPE_COLORS.get(color_key, TYPE_COLORS["other"])
    w = size
    h = size * 0.86
    draw_round_rect(canvas, cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2,
                    size * 0.22, fill=color, outline="")
    from .theme import font
    canvas.create_text(cx, cy, text=label, fill="#ffffff",
                       font=font(int(size * 0.32), "bold"))
