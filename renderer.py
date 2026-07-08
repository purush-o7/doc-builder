"""
Renders content element dicts into python-docx XML elements.

Element types
-------------
content          Unified type with "style" sub-field:
                   heading      bold, centered
                   subsection   italic, no forced align
                   para         plain body paragraph
                   para_rich    mixed-format runs  (needs "runs" list)
                   table        bordered table     (needs caption, headers, rows)
figure_caption   Centered caption paragraph (runs list)
figure           Centered image + caption  (path, width, caption runs)
authors_table    2-column author block (authors list)
placeholder      Red bold text
spacer           Empty paragraph

Formatting
----------
Every element may carry a "fmt" dict with explicitly-set paragraph formatting:
  space_before  float (points)
  space_after   float (points)
  line_spacing  int   (240 = single, 360 = 1.5×, 480 = double)
  line_rule     "auto" | "exact" | "atLeast"
  indent_left   float (points)
  indent_right  float (points)
  indent_first  float (points)

Every run may carry explicit overrides:
  font_size     float (points)
  font_name     str
  color         str   (hex, e.g. "FF0000")
"""

from pathlib import Path
from docx.shared import Pt, Inches, RGBColor, Emu
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from lxml import etree

from .style import merge, resolve_align, GLOBAL_DEFAULTS

_UNSET = object()   # sentinel: "align key absent" vs "align key present but null"


def _get_align(item: dict, default: str = None):
    """
    Return the alignment to apply, respecting three states:
      key missing           → use *default* (backward compat with hand-written JSON)
      key present, null     → None  (don't set — let style inherit)
      key present, "both"   → "both" (set explicitly)
    """
    val = item.get("align", _UNSET)
    if val is _UNSET:
        return default
    return val   # could be None or a string


class Renderer:
    def __init__(self, doc, defaults: dict, base_dir=None):
        self.doc      = doc
        self.body     = doc.element.body
        self.defaults = merge(GLOBAL_DEFAULTS, defaults)
        self.base_dir = Path(base_dir) if base_dir else Path(".")

    # ── public ───────────────────────────────────────────────────────────────

    def render_content(self, content: list) -> list:
        out = []
        for item in content:
            out.extend(self._render(item))
        return out

    # ── dispatcher ───────────────────────────────────────────────────────────

    def _render(self, item: dict) -> list:
        t = item.get("type")
        raw_style = item.get("style")
        style_overrides = raw_style if isinstance(raw_style, dict) else {}
        style = merge(self.defaults, style_overrides)

        if t == "content":
            return self._render_content(item, style)
        if t == "figure_caption":
            return [self._fig_caption(item["runs"], item.get("fmt", {}), style,
                                      align=item.get("align", _UNSET))]
        if t == "figure":
            return self._figure(item, style)
        if t == "authors_table":
            return self._authors_table(item, style)
        if t == "spacer":
            return [self._spacer(item)]
        if t == "placeholder":
            return [self._placeholder(item, style)]
        # backward compat
        if t in ("heading", "subsection", "para", "para_rich", "table"):
            return self._render_content({**item, "type": "content",
                                         "style": t}, style)
        raise ValueError(f"Unknown element type: {t!r}")

    def _render_content(self, item: dict, style: dict) -> list:
        s = item.get("style")
        if isinstance(s, str):
            pass
        else:
            s = "para"
        if s == "heading":    return [self._heading(item, style)]
        if s == "subsection": return [self._subsection(item, style)]
        if s == "para":       return [self._para(item, style)]
        if s == "para_rich":  return [self._para_rich(item, style)]
        if s == "table":      return self._table(item, style)
        raise ValueError(f"Unknown content style: {s!r}")

    # ── element builders ─────────────────────────────────────────────────────

    def _heading(self, item: dict, style: dict):
        p = self.doc.add_paragraph()
        r = p.add_run(item["text"])
        r.bold   = True
        r.italic = item.get("bold_italic") or None
        align = _get_align(item, default="center")
        if align is not None:
            p.alignment = resolve_align("align", {"align": align})
        run_fmt = {"font_size": item.get("font_size")} if item.get("font_size") else {}
        self._apply_run_fmt(r, run_fmt, style)
        self._apply_para_fmt(p, item.get("fmt", {}))
        sb = item.get("section_break") if hasattr(item, "get") and isinstance(item, dict) else None
        return self._detach(p, sb)

    def _subsection(self, item: dict, style: dict):
        p = self.doc.add_paragraph()
        r = p.add_run(item["text"])
        r.bold   = item.get("bold") or None
        r.italic = item.get("italic", True)
        align = _get_align(item, default=None)
        if align is not None:
            p.alignment = resolve_align("align", {"align": align})
        run_fmt = {"font_size": item.get("font_size")} if item.get("font_size") else {}
        self._apply_run_fmt(r, run_fmt, style)
        self._apply_para_fmt(p, item.get("fmt", {}))
        sb = item.get("section_break") if hasattr(item, "get") and isinstance(item, dict) else None
        return self._detach(p, sb)

    def _para(self, item: dict, style: dict):
        p = self.doc.add_paragraph()
        r = p.add_run(item["text"])
        # use original alignment; fall back to body_align default only if key absent
        align = _get_align(item, default=style.get("body_align", "justify"))
        if align is not None:
            p.alignment = resolve_align("align", {"align": align})
        run_fmt = {"font_size": item.get("font_size")} if item.get("font_size") else {}
        self._apply_run_fmt(r, run_fmt, style)
        self._apply_para_fmt(p, item.get("fmt", {}))
        sb = item.get("section_break") if hasattr(item, "get") and isinstance(item, dict) else None
        return self._detach(p, sb)

    def _para_rich(self, item: dict, style: dict):
        p = self.doc.add_paragraph()
        for rs in item["runs"]:
            r = p.add_run(rs["text"])
            r.bold      = rs.get("bold")      or None
            r.italic    = rs.get("italic")    or None
            r.underline = rs.get("underline") or None
            self._apply_run_fmt(r, rs, style)
        align = _get_align(item, default=style.get("body_align", "justify"))
        if align is not None:
            p.alignment = resolve_align("align", {"align": align})
        self._apply_para_fmt(p, item.get("fmt", {}))
        sb = item.get("section_break") if hasattr(item, "get") and isinstance(item, dict) else None
        return self._detach(p, sb)

    def _table(self, item: dict, style: dict) -> list:
        cap_runs = item.get("caption_runs")
        cap_fmt  = item.get("caption_fmt", {})
        if cap_runs:
            caption_el = self._fig_caption(cap_runs, cap_fmt, style,
                                           align=item.get("caption_align", "center"))
        else:
            cap_italic = item.get("caption_italic", False)
            caption_el = self._caption_para(item.get("caption", ""), style,
                                            cap_fmt, italic=cap_italic)

        # Verbatim path: re-inject the original table XML exactly as captured
        # (borders, fixed width, column widths, cell fonts, tblStyle, header-repeat).
        raw_xml = item.get("raw_xml")
        if raw_xml:
            return [caption_el, parse_xml(raw_xml)]

        headers    = item["headers"]
        rows       = item["rows"]
        col_widths = item.get("col_widths") or style.get("col_widths")
        font_size  = item.get("font_size")  or style.get("table_font_size", 8)
        header_bold = item.get("header_bold")
        if header_bold is None:
            header_bold = style.get("table_header_bold", True)
        cell_align = item.get("cell_align") or style.get("table_cell_align", "center")

        table = self.doc.add_table(rows=1 + len(rows), cols=len(headers))
        table.style = style.get("table_style", "TableNormal")
        if style.get("table_borders", True):
            self._add_borders(table._tbl, style)

        for i, h in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = ""
            run = cell.paragraphs[0].add_run(str(h))
            run.bold = header_bold
            run.font.size = Pt(font_size)
            cell.paragraphs[0].alignment = resolve_align("align", {"align": cell_align})
            if col_widths:
                cell.width = Inches(col_widths[i])

        for ri, row_data in enumerate(rows):
            for ci, val in enumerate(row_data):
                cell = table.rows[ri + 1].cells[ci]
                cell.text = ""
                run = cell.paragraphs[0].add_run(str(val))
                run.font.size = Pt(font_size)
                cell.paragraphs[0].alignment = resolve_align("align", {"align": cell_align})

        tbl_el = table._tbl
        self.body.remove(tbl_el)
        return [caption_el, tbl_el]

    def _fig_caption(self, runs: list, fmt: dict, style: dict, align=_UNSET):
        p = self.doc.add_paragraph()
        # use element's original align; fall back to caption_align default
        a = align if align is not _UNSET else style.get("caption_align", "center")
        if a is not None:
            p.alignment = resolve_align("align", {"align": a})
        for rs in runs:
            r = p.add_run(rs["text"])
            r.bold   = rs.get("bold")   or None
            r.italic = rs.get("italic") or None
            self._apply_run_fmt(r, rs, style)
        self._apply_para_fmt(p, fmt)
        return self._detach(p)

    def _figure(self, item: dict, style: dict) -> list:
        """Render the image paragraph only — caption is a separate figure_caption element."""
        width     = item.get("width") or style.get("figure_width", 3.4)
        fmt       = item.get("fmt", {})

        p = self.doc.add_paragraph()
        align = _get_align(item, default=style.get("figure_align", "center"))
        if align is not None:
            p.alignment = resolve_align("align", {"align": align})

        img_path = (self.base_dir / item["path"]).resolve()
        w_emu = item.get("width_emu")
        h_emu = item.get("height_emu")
        if w_emu and h_emu:
            # exact original display box (both dimensions) — preserves any
            # manual reshaping instead of recomputing height from native AR
            p.add_run().add_picture(str(img_path), width=Emu(w_emu), height=Emu(h_emu))
        else:
            p.add_run().add_picture(str(img_path), width=Inches(width))
        self._apply_para_fmt(p, fmt)

        sb = item.get("section_break") if hasattr(item, "get") and isinstance(item, dict) else None
        return [self._detach(p, sb)]

    def _authors_table(self, item: dict, style: dict) -> list:
        # Verbatim path: re-inject the original author-block table XML exactly.
        raw_xml = item.get("raw_xml")
        if raw_xml:
            return [parse_xml(raw_xml)]

        authors = item["authors"]
        ncols   = 2
        nrows   = max(1, (len(authors) + 1) // 2)

        table = self.doc.add_table(rows=nrows, cols=ncols)
        table.style = style.get("table_style", "TableNormal")

        for ai, author in enumerate(authors):
            ri = ai // ncols
            ci = ai % ncols
            cell = table.cell(ri, ci)
            cell.text = ""
            for li, line in enumerate(author.get("lines", [])):
                p = cell.paragraphs[-1] if (li == 0 and not cell.paragraphs[0].text) \
                    else cell.add_paragraph()
                p.alignment = resolve_align("align", {"align": "center"})
                # apply per-line run formatting if available
                line_runs = author.get("runs", [None] * len(author.get("lines", [])))[li]
                if line_runs:
                    for rs in line_runs:
                        r = p.add_run(rs["text"])
                        r.bold   = rs.get("bold")   or None
                        r.italic = rs.get("italic") or None
                        self._apply_run_fmt(r, rs, style)
                else:
                    p.add_run(line)

        tbl_el = table._tbl
        self.body.remove(tbl_el)
        return [tbl_el]

    def _spacer(self, item: dict = None):
        p = self.doc.add_paragraph()
        item = item or {}
        align = _get_align(item, default=None)
        if align is not None:
            p.alignment = resolve_align("align", {"align": align})
        self._apply_para_fmt(p, item.get("fmt", {}))
        sb = item.get("section_break") if hasattr(item, "get") and isinstance(item, dict) else None
        return self._detach(p, sb)

    def _placeholder(self, item: dict, style: dict):
        p = self.doc.add_paragraph()
        r = p.add_run(item.get("text", "[Figure placeholder]"))
        r.bold = True
        r.font.color.rgb = RGBColor(255, 0, 0)
        r.font.size = Pt(item.get("font_size") or 11)
        p.alignment = resolve_align("align", {"align": item.get("align", "center")})
        self._apply_para_fmt(p, item.get("fmt", {}))
        sb = item.get("section_break") if hasattr(item, "get") and isinstance(item, dict) else None
        return self._detach(p, sb)

    # ── formatting helpers ───────────────────────────────────────────────────

    def _apply_para_fmt(self, p, fmt: dict):
        """Apply explicit paragraph spacing / indent from fmt dict."""
        if not fmt:
            return
        pf = p.paragraph_format
        if fmt.get("space_before") is not None:
            pf.space_before = Pt(fmt["space_before"])
        if fmt.get("space_after") is not None:
            pf.space_after  = Pt(fmt["space_after"])
        if fmt.get("indent_left") is not None:
            pf.left_indent  = Pt(fmt["indent_left"])
        if fmt.get("indent_right") is not None:
            pf.right_indent = Pt(fmt["indent_right"])
        if fmt.get("indent_first") is not None:
            pf.first_line_indent = Pt(fmt["indent_first"])

    @staticmethod
    def _apply_run_fmt(run, rs: dict, style: dict):
        """Apply explicit run-level font overrides."""
        fs = rs.get("font_size") or style.get("font_size")
        if fs:
            run.font.size = Pt(fs)
        fn = rs.get("font_name") or style.get("font_name")
        if fn:
            run.font.name = fn
        color = rs.get("color")
        if color:
            run.font.color.rgb = RGBColor(
                int(color[0:2], 16),
                int(color[2:4], 16),
                int(color[4:6], 16),
            )

    # ── shared helpers ───────────────────────────────────────────────────────

    def _detach(self, p, section_break: dict = None):
        p.style = self.doc.styles["normal"]
        if section_break:
            self._inject_sect_pr(p._element, section_break)
        el = p._element
        self.body.remove(el)
        return el

    @staticmethod
    def _inject_sect_pr(para_el, sb: dict):
        """Embed a w:sectPr into a paragraph's pPr (creates pPr if needed)."""
        pPr = para_el.find(qn("w:pPr"))
        if pPr is None:
            pPr = OxmlElement("w:pPr")
            para_el.insert(0, pPr)

        sectPr = OxmlElement("w:sectPr")

        if sb.get("type"):
            t = OxmlElement("w:type")
            t.set(qn("w:val"), sb["type"])
            sectPr.append(t)

        if sb.get("cols"):
            cols_el = OxmlElement("w:cols")
            c = sb["cols"]
            if "num"        in c: cols_el.set(qn("w:num"),        str(c["num"]))
            if "space"      in c: cols_el.set(qn("w:space"),      str(c["space"]))
            if "equalWidth" in c: cols_el.set(qn("w:equalWidth"), str(c["equalWidth"]))
            sectPr.append(cols_el)

        if sb.get("pg_sz"):
            pgSz = OxmlElement("w:pgSz")
            for k, v in sb["pg_sz"].items():
                pgSz.set(qn(f"w:{k}"), str(v))
            sectPr.append(pgSz)

        if sb.get("pg_mar"):
            pgMar = OxmlElement("w:pgMar")
            for k, v in sb["pg_mar"].items():
                pgMar.set(qn(f"w:{k}"), str(v))
            sectPr.append(pgMar)

        if sb.get("pg_num_start") is not None:
            pgNum = OxmlElement("w:pgNumType")
            pgNum.set(qn("w:start"), str(sb["pg_num_start"]))
            sectPr.append(pgNum)

        # sectPr must be the last child of pPr
        pPr.append(sectPr)

    def _caption_para(self, text: str, style: dict,
                      fmt: dict = None, italic: bool = False):
        p = self.doc.add_paragraph()
        r = p.add_run(text)
        r.bold   = True
        r.italic = italic or None
        p.alignment = resolve_align("align", {"align": "center"})
        if fmt:
            self._apply_para_fmt(p, fmt)
        return self._detach(p)

    @staticmethod
    def _add_borders(tbl, style: dict = None):
        sz    = (style or {}).get("table_border_sz",    "4")
        color = (style or {}).get("table_border_color", "000000")
        tblPr = tbl.tblPr if tbl.tblPr is not None else tbl._add_tblPr()
        tblBorders = etree.SubElement(tblPr, qn("w:tblBorders"))
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = etree.SubElement(tblBorders, qn(f"w:{edge}"))
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"),  str(sz))
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), str(color))
