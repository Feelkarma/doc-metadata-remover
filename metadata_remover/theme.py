"""Design tokens for the doc-metadata-remover desktop UI.

A single source of truth for colours, typography, spacing, and corner radii so
the interface has one consistent "privacy vault" look.  Everything here is pure
data / helpers — no widgets and no business logic — so it can be imported freely
without side effects.

The palette is a modern dark theme: deep navy/charcoal background, layered
blue-black surfaces, a restrained electric-blue accent, and near-white text with
muted blue-grey secondary text.
"""

from __future__ import annotations

import tkinter.font as tkfont
from typing import Optional


# ---------------------------------------------------------------------------
# Colour tokens
# ---------------------------------------------------------------------------
class Color:
    # Backgrounds (deep charcoal / navy) and layered surfaces.
    BG = "#0b0f1a"            # app background (near-black navy)
    SURFACE = "#121826"       # primary card surface
    SURFACE_ALT = "#161d2e"   # slightly raised surface (rows, insets)
    SURFACE_HI = "#1c2437"    # hover / highlighted surface
    ELEVATED = "#0f1524"      # console / recessed panel

    # Hairline borders — very subtle, low-contrast blue-grey.
    BORDER = "#232c42"
    BORDER_HI = "#33405f"     # focus / drag-over border

    # Accent (professional electric blue) + subtle secondary highlights.
    ACCENT = "#3d7bff"
    ACCENT_HOVER = "#5590ff"
    ACCENT_PRESSED = "#2f63d6"
    ACCENT_SOFT = "#1a2murder"  # placeholder, overwritten below
    ACCENT_GLOW = "#6aa1ff"
    CYAN = "#39d0d8"
    VIOLET = "#8b7dff"

    # Semantic states.
    SUCCESS = "#31c48d"
    SUCCESS_SOFT = "#123028"
    WARNING = "#f5b451"
    ERROR = "#f76d6d"
    ERROR_SOFT = "#2c1620"

    # Text.
    TEXT = "#eef2fb"          # near-white primary
    TEXT_MUTED = "#8b98b8"    # muted blue-grey secondary
    TEXT_FAINT = "#5c6885"    # captions / disabled
    TEXT_ON_ACCENT = "#ffffff"

    # Disabled accents.
    ACCENT_DISABLED = "#2a3550"
    TEXT_DISABLED = "#5a6party"  # placeholder, overwritten below


# Fix up a couple of tokens that must be valid hex (kept readable above).
Color.ACCENT_SOFT = "#16233f"       # translucent-looking blue wash on dark
Color.TEXT_DISABLED = "#566179"


# File-type accent colours (used for the little file-type chips).
TYPE_COLORS = {
    "word": "#2b7cd3",
    "powerpoint": "#d24726",
    "excel": "#1d8a4c",
    "pdf": "#d0342c",
    "other": "#5b6b8c",
}

TYPE_LABELS = {
    ".docx": ("DOC", "word"),
    ".doc": ("DOC", "word"),
    ".pptx": ("PPT", "powerpoint"),
    ".ppt": ("PPT", "powerpoint"),
    ".xlsx": ("XLS", "excel"),
    ".xls": ("XLS", "excel"),
    ".pdf": ("PDF", "pdf"),
}


def type_for_extension(ext: str):
    """Return ``(label, color_key)`` for a file extension (lower-case, dotted)."""
    return TYPE_LABELS.get(ext.lower(), ("FILE", "other"))


# ---------------------------------------------------------------------------
# Spacing (deliberate 8px system) and radii
# ---------------------------------------------------------------------------
class Space:
    XS = 4
    SM = 8
    MD = 12
    LG = 16
    XL = 24
    XXL = 32


class Radius:
    SM = 8
    MD = 12
    LG = 16
    PILL = 999


# ---------------------------------------------------------------------------
# Typography
# ---------------------------------------------------------------------------
# Preferred families per platform; we pick the first one actually installed.
_FONT_STACK = [
    "Segoe UI",          # Windows
    "SF Pro Text",       # macOS (modern)
    "Helvetica Neue",    # macOS (fallback)
    "Inter",
    "Ubuntu",            # common Linux
    "Cantarell",
    "DejaVu Sans",       # very common Linux fallback
    "Arial",
]
_MONO_STACK = [
    "Cascadia Code",
    "Cascadia Mono",
    "SF Mono",
    "JetBrains Mono",
    "Consolas",
    "Menlo",
    "DejaVu Sans Mono",
    "Courier New",
]

_family_cache: Optional[str] = None
_mono_cache: Optional[str] = None


def _pick(stack, fallback):
    try:
        available = {f.lower() for f in tkfont.families()}
    except Exception:
        return fallback
    for name in stack:
        if name.lower() in available:
            return name
    return fallback


def ui_family() -> str:
    """Return the best available UI font family (cached)."""
    global _family_cache
    if _family_cache is None:
        _family_cache = _pick(_FONT_STACK, "TkDefaultFont")
    return _family_cache


def mono_family() -> str:
    """Return the best available monospace font family (cached)."""
    global _mono_cache
    if _mono_cache is None:
        _mono_cache = _pick(_MONO_STACK, "TkFixedFont")
    return _mono_cache


def font(size: int = 10, weight: str = "normal", slant: str = "roman"):
    """Build a UI font tuple, e.g. ``font(11, "bold")``."""
    return (ui_family(), size, weight) if slant == "roman" else (ui_family(), size, weight, slant)


def mono(size: int = 9, weight: str = "normal"):
    """Build a monospace font tuple."""
    return (mono_family(), size, weight)


# Named text roles (call these to get a font tuple).
def font_title():
    return font(19, "bold")


def font_subtitle():
    return font(10)


def font_heading():
    return font(12, "bold")


def font_body():
    return font(10)


def font_body_strong():
    return font(10, "bold")


def font_button():
    return font(11, "bold")


def font_caption():
    return font(9)


def font_chip():
    return font(9, "bold")
