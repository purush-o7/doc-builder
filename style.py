"""
Style defaults and merge logic.

Defaults cascade: global defaults < operation defaults < element style < run attrs.
All values are plain Python types; alignment is a string resolved to WD_* on use.
"""

from docx.enum.text import WD_ALIGN_PARAGRAPH

GLOBAL_DEFAULTS = {
    # text
    "font_name": None,          # None = inherit from document default
    "font_size": None,          # None = inherit; number = points
    "body_align": "justify",    # default body paragraph alignment
    # table
    "table_style":        "TableNormal",  # Word table style name
    "table_font_size":    8,
    "table_header_bold":  True,
    "table_borders":      True,
    "table_border_sz":    "4",            # border line width (eighths of a point)
    "table_border_color": "000000",       # border colour (hex, no #)
    "table_cell_align":   "center",
    # figure / caption
    "figure_width":   3.4,               # inches, used when no explicit width set
    "figure_align":   "center",
    "caption_align":  "center",
}

_ALIGN = {
    "center":  WD_ALIGN_PARAGRAPH.CENTER,
    "left":    WD_ALIGN_PARAGRAPH.LEFT,
    "right":   WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "both":    WD_ALIGN_PARAGRAPH.JUSTIFY,   # OOXML uses "both" for justify
    None:      None,
}


def merge(base: dict, *overrides) -> dict:
    """Return a new dict merging base then each override, skipping None values."""
    result = dict(base)
    for override in overrides:
        if override:
            for k, v in override.items():
                if v is not None:
                    result[k] = v
    return result


def resolve_align(key: str, style: dict):
    """Return WD_ALIGN_PARAGRAPH constant for style[key], or None."""
    return _ALIGN.get(style.get(key))
