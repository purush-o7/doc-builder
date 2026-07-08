"""
classifier.py — converts raw parsed elements into typed content elements
and groups them into named sections.

Image paragraphs (with "image" key from parser) become "figure" elements.
Alignment, spacing, section breaks and all formatting are preserved verbatim.
"""

import re
from pathlib import Path

# ── Document-type patterns (override via classify_and_group params if needed) ─

# Matches "Fig. 3. Caption..." — identifies figure captions
FIGURE_CAPTION_RE = re.compile(r"Fig\.\s*\d", re.IGNORECASE)

# Prefix that marks a table caption paragraph (case-insensitive)
TABLE_CAPTION_PREFIX = "TABLE"

# Roman-numeral section headings (I., II., III., IV., V., VI., ...)
SECTION_HEADING_RE = re.compile(
    r"^(I{1,3}V?|VI{0,3}|IX|X{1,3})\.\s+\S|"
    r"^\.\s*REFERENCES|"
    r"^REFERENCES$",
    re.IGNORECASE,
)

# Bold paragraphs with no explicit alignment that still look like headings
HEADING_DETECT_RE = re.compile(r"^(I{1,3}V?|VI{0,3}|IV)\b")


def classify_and_group(
    raw: list,
    section_map: dict = None,
    spec_dir: Path = None,
    first_table_is_authors: bool = True,
    figure_caption_re: re.Pattern = FIGURE_CAPTION_RE,
    table_caption_prefix: str = TABLE_CAPTION_PREFIX,
    section_heading_re: re.Pattern = SECTION_HEADING_RE,
    heading_detect_re: re.Pattern = HEADING_DETECT_RE,
) -> dict:
    """
    Parameters
    ----------
    raw                    : flat list from parser.parse()
    section_map            : optional {heading_text: filename_stem}
    spec_dir               : directory of the spec JSON (makes image paths relative)
    first_table_is_authors : treat the first table as an authors block (IEEE default)
    figure_caption_re      : regex that identifies a figure caption paragraph
    table_caption_prefix   : string prefix that identifies a table caption
    section_heading_re     : regex that starts a new section (triggers grouping)
    heading_detect_re      : regex for un-aligned bold paragraphs that are headings

    Returns
    -------
    dict {section_name: [elements]}
    """
    cfg = _Config(first_table_is_authors, figure_caption_re,
                  table_caption_prefix, section_heading_re, heading_detect_re)
    elements = _classify_all(raw, spec_dir=spec_dir, cfg=cfg)
    return _group_sections(elements, section_map, section_heading_re)


# ── config container ──────────────────────────────────────────────────────────

class _Config:
    __slots__ = ("first_table_is_authors", "fig_cap_re", "tbl_cap_prefix",
                 "section_re", "heading_re")
    def __init__(self, first_table_is_authors, fig_cap_re, tbl_cap_prefix,
                 section_re, heading_re):
        self.first_table_is_authors = first_table_is_authors
        self.fig_cap_re   = fig_cap_re
        self.tbl_cap_prefix = tbl_cap_prefix.upper()
        self.section_re   = section_re
        self.heading_re   = heading_re


# ── step 1: classify ──────────────────────────────────────────────────────────

def _classify_all(raw: list, spec_dir: Path = None, cfg: _Config = None) -> list:
    cfg = cfg or _Config(True, FIGURE_CAPTION_RE, TABLE_CAPTION_PREFIX,
                         SECTION_HEADING_RE, HEADING_DETECT_RE)
    out         = []
    tbl_count   = 0
    pending_cap = None

    for item in raw:
        if item["type"] not in ("p", "tbl"):
            continue

        if item["type"] == "tbl":
            if tbl_count == 0 and cfg.first_table_is_authors:
                out.append(_authors_table(item))
            else:
                out.append(_table(item, caption=pending_cap or ""))
                pending_cap = None
            tbl_count += 1
            continue

        el = _classify_para(item, spec_dir=spec_dir, cfg=cfg)

        # carry section break
        if item.get("section_break"):
            el["section_break"] = item["section_break"]

        # hold table captions for the next table
        cap_text = el.get("text", _first_run_text(el)).strip().upper()
        if (el.get("style") in ("heading", "para_rich")
                and cap_text.startswith(cfg.tbl_cap_prefix)):
            pending_cap = el
            continue

        out.append(el)

    return out


def _first_run_text(el: dict) -> str:
    runs = el.get("runs", [])
    return runs[0]["text"] if runs else ""


def _classify_para(p: dict, spec_dir: Path = None, cfg: _Config = None) -> dict:
    runs  = p["runs"]
    full  = p["full"].strip()
    fmt   = p["fmt"]
    align = fmt.get("align")

    # ── image paragraph ───────────────────────────────────────────────────────
    if p.get("image"):
        img = p["image"]
        # make path relative to spec_dir so JSON is portable
        path = img["path"]
        if spec_dir:
            try:
                path = str(Path(img["path"]).relative_to(spec_dir)).replace("\\", "/")
            except ValueError:
                pass   # keep absolute if can't relativise
        el = {"type": "figure", "path": path,
              "width": img["width_in"], "align": align}
        return _with_fmt(el, fmt)

    # ── empty paragraph → spacer (preserve all formatting) ───────────────────
    cfg = cfg or _Config(True, FIGURE_CAPTION_RE, TABLE_CAPTION_PREFIX,
                         SECTION_HEADING_RE, HEADING_DETECT_RE)

    if not full:
        el = {"type": "spacer", "align": align}
        pf = _spacing_fmt(fmt)
        if pf:
            el["fmt"] = pf
        return el

    # ── single-run paragraphs ─────────────────────────────────────────────────
    if len(runs) == 1:
        r  = runs[0]
        b  = r.get("bold",   False)
        i  = r.get("italic", False)
        fs = r.get("font_size")

        if b and (align == "center" or
                  (align is None and cfg.heading_re.match(full.strip()))):
            el = {"type": "content", "style": "heading",
                  "text": full, "align": align}
            if i:  el["bold_italic"] = True
            if fs: el["font_size"]   = fs
            return _with_fmt(el, fmt)

        if i and not b:
            el = {"type": "content", "style": "subsection",
                  "text": full, "align": align}
            if fs: el["font_size"] = fs
            return _with_fmt(el, fmt)

        if not b and not i:
            el = {"type": "content", "style": "para",
                  "text": full, "align": align}
            if fs: el["font_size"] = fs
            return _with_fmt(el, fmt)

        return _with_fmt({"type": "content", "style": "para_rich",
                          "runs": [_clean_run(r)], "align": align}, fmt)

    # ── multiple runs ─────────────────────────────────────────────────────────
    first_txt = next((r["text"] for r in runs if r["text"].strip()), "")

    if cfg.fig_cap_re.match(first_txt.strip()):
        return _with_fmt({"type": "figure_caption",
                          "runs": [_clean_run(r) for r in runs],
                          "align": align}, fmt)

    if align == "center":
        all_bold = all(r.get("bold") for r in runs)
        if all_bold and not any(r.get("italic") for r in runs):
            return _with_fmt({"type": "content", "style": "heading",
                              "text": full, "align": align}, fmt)
        return _with_fmt({"type": "content", "style": "para_rich",
                          "runs": [_clean_run(r) for r in runs],
                          "align": align}, fmt)

    return _with_fmt({"type": "content", "style": "para_rich",
                      "runs": [_clean_run(r) for r in runs],
                      "align": align}, fmt)


def _clean_run(r: dict) -> dict:
    out = {"text": r["text"]}
    for k in ("bold", "italic", "underline", "font_size", "font_name", "color"):
        if r.get(k):
            out[k] = r[k]
    return out


def _spacing_fmt(fmt: dict) -> dict:
    return {k: v for k, v in fmt.items()
            if k not in ("align",) and v is not None}


def _with_fmt(el: dict, fmt: dict) -> dict:
    pf = _spacing_fmt(fmt)
    if pf:
        el["fmt"] = pf
    return el


def _authors_table(item: dict) -> dict:
    authors = []
    for row in item["rows"]:
        for cell_lines in row:
            if cell_lines:
                authors.append({
                    "lines": [l["text"] for l in cell_lines],
                    "runs":  [l["runs"] for l in cell_lines],
                })
    return {"type": "authors_table", "authors": authors}


def _table(item: dict, caption) -> dict:
    rows      = item["rows"]
    headers   = [" ".join(l["text"] for l in c) for c in rows[0]]
    data_rows = [[" ".join(l["text"] for l in c) for c in row]
                 for row in rows[1:]]
    if isinstance(caption, dict):
        cap_text   = caption.get("text", "")
        cap_runs   = caption.get("runs")
        cap_fmt    = caption.get("fmt", {})
        cap_italic = bool(caption.get("bold_italic"))
    else:
        cap_text   = caption or ""
        cap_runs   = None
        cap_fmt    = {}
        cap_italic = False
    el = {"type": "content", "style": "table",
          "caption": cap_text, "headers": headers, "rows": data_rows}
    if cap_runs:   el["caption_runs"]   = cap_runs
    if cap_fmt:    el["caption_fmt"]    = cap_fmt
    if cap_italic: el["caption_italic"] = True
    return el


# ── step 2: group into sections ───────────────────────────────────────────────

def _group_sections(elements: list, section_map: dict,
                    section_re: re.Pattern = None) -> dict:
    section_re   = section_re or SECTION_HEADING_RE
    sections     = {}
    current_name = "header"
    current_buf  = []

    for el in elements:
        if el.get("style") == "heading":
            text = el.get("text", "")
            if section_re.match(text):
                if current_buf:
                    sections[current_name] = list(current_buf)
                current_buf  = [el]
                current_name = _section_name(text, section_map)
                continue
        current_buf.append(el)

    if current_buf:
        sections[current_name] = list(current_buf)

    return sections


def _section_name(text: str, section_map: dict) -> str:
    if section_map and text in section_map:
        return section_map[text]
    slug = re.sub(r"^[IVXivx]+\.\s*", "", text.strip())
    slug = re.sub(r"[^a-z0-9]+", "_", slug.lower()).strip("_")
    return slug or "section"
