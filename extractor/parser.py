"""
parser.py — reads a .docx and produces a flat list of raw elements,
capturing ALL explicitly-set formatting plus embedded images.

Image paragraphs are detected via w:drawing elements. Their images are
extracted to media_dir and referenced by relative path.
"""

from pathlib import Path
from docx import Document
from docx.oxml.ns import qn

# Namespace URIs for drawing / image elements
_NS_WP  = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
_NS_A   = "http://schemas.openxmlformats.org/drawingml/2006/main"
_NS_R   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

# w:drawing lives in the main wordprocessingml namespace
_WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Fallback image width when the document provides no extent value (inches)
_DEFAULT_IMAGE_WIDTH_IN = 3.4


def parse(docx_path: str, media_dir: str = None) -> tuple:
    """
    Parse a .docx file into a flat list of raw elements.

    Parameters
    ----------
    docx_path : path to the source .docx
    media_dir : if given, extract embedded images into this directory

    Returns
    -------
    (raw_elements, image_map)
    raw_elements : list of dicts  (type "p", "tbl", or other)
    image_map    : {rId: absolute_path_str}  (empty if media_dir is None)
    """
    doc  = Document(docx_path)
    body = doc.element.body

    image_map: dict[str, str] = {}
    if media_dir:
        image_map = _extract_images(doc, Path(media_dir))

    raw = []
    for idx, child in enumerate(body):
        tag = child.tag.split("}")[1] if "}" in child.tag else child.tag
        if tag == "p":
            raw.append(_parse_para(idx, child, doc, image_map))
        elif tag == "tbl":
            raw.append(_parse_table(idx, child))
        else:
            raw.append({"idx": idx, "type": tag})

    return raw, image_map


# ── image extraction ──────────────────────────────────────────────────────────

def _extract_images(doc, media_dir: Path) -> dict:
    """Save every image relationship to media_dir; return {rId: abs_path}."""
    media_dir.mkdir(parents=True, exist_ok=True)
    image_map = {}
    IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

    for rId, rel in doc.part.rels.items():
        if rel.reltype != IMAGE_REL:
            continue
        part = rel.target_part
        ext  = Path(part.partname).suffix.lower()   # e.g. ".png"
        dest = media_dir / f"{rId}{ext}"
        with open(dest, "wb") as f:
            f.write(part.blob)
        image_map[rId] = str(dest)

    return image_map


def _find_image(para_el, image_map: dict) -> dict | None:
    """
    Return image info dict if this paragraph embeds a drawing, else None.
    dict keys: rId, path (absolute), width_in (float, inches)
    """
    drawings = para_el.findall(f"{{{_WML}}}r/{{{_WML}}}drawing")
    if not drawings:
        # also try nested (sometimes drawing is deeper)
        drawings = para_el.findall(f".//{{{_WML}}}drawing")
    if not drawings:
        return None

    drawing = drawings[0]

    # width from wp:extent/@cx (EMU → inches)
    extent = drawing.find(f".//{{{_NS_WP}}}extent")
    width_in = _DEFAULT_IMAGE_WIDTH_IN
    if extent is not None:
        cx = extent.get("cx")
        if cx:
            width_in = round(int(cx) / 914400, 3)

    # rId from a:blip/@r:embed
    blip = drawing.find(f".//{{{_NS_A}}}blip")
    if blip is None:
        return None
    rId = blip.get(f"{{{_NS_R}}}embed")
    if not rId or rId not in image_map:
        return None

    return {"rId": rId, "path": image_map[rId], "width_in": width_in}


# ── paragraph parsing ─────────────────────────────────────────────────────────

def _parse_para(idx: int, node, doc, image_map: dict) -> dict:
    runs = []
    for r in node.iter(qn("w:r")):
        rPr    = r.find(qn("w:rPr"))
        txt    = "".join(t.text or "" for t in r.iter(qn("w:t")))
        if not txt:
            continue
        run = {"text": txt}
        if rPr is not None:
            if rPr.find(qn("w:b"))  is not None: run["bold"]   = True
            if rPr.find(qn("w:i"))  is not None: run["italic"] = True
            if rPr.find(qn("w:u"))  is not None: run["underline"] = True
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                run["font_size"] = int(sz.get(qn("w:val"))) / 2
            fn = rPr.find(qn("w:rFonts"))
            if fn is not None:
                name = fn.get(qn("w:ascii")) or fn.get(qn("w:hAnsi"))
                if name:
                    run["font_name"] = name
            color = rPr.find(qn("w:color"))
            if color is not None:
                val = color.get(qn("w:val"))
                if val and val.lower() not in ("auto", "000000"):
                    run["color"] = val
        runs.append(run)

    fmt  = _para_fmt(node)
    full = "".join(r["text"] for r in runs)
    el   = {"idx": idx, "type": "p", "full": full, "runs": runs, "fmt": fmt}

    # Detect embedded image
    img = _find_image(node, image_map)
    if img:
        el["image"] = img

    # Detect inline section break
    _pPr = node.find(qn("w:pPr"))
    if _pPr is not None:
        sp = _pPr.find(qn("w:sectPr"))
        if sp is not None:
            el["section_break"] = _parse_sect_pr(sp)

    return el


def _para_fmt(node) -> dict:
    fmt = {}
    pPr = node.find(qn("w:pPr"))
    if pPr is None:
        return fmt
    jc = pPr.find(qn("w:jc"))
    if jc is not None:
        fmt["align"] = jc.get(qn("w:val"))
    sp = pPr.find(qn("w:spacing"))
    if sp is not None:
        before = sp.get(qn("w:before"))
        after  = sp.get(qn("w:after"))
        line   = sp.get(qn("w:line"))
        rule   = sp.get(qn("w:lineRule"))
        if before is not None: fmt["space_before"] = _twips(before)
        if after  is not None: fmt["space_after"]  = _twips(after)
        if line   is not None: fmt["line_spacing"]  = int(line)
        if rule   is not None: fmt["line_rule"]     = rule
    ind = pPr.find(qn("w:ind"))
    if ind is not None:
        left  = ind.get(qn("w:left"))
        right = ind.get(qn("w:right"))
        first = ind.get(qn("w:firstLine"))
        if left  is not None and int(left)  != 0: fmt["indent_left"]  = _twips(left)
        if right is not None and int(right) != 0: fmt["indent_right"] = _twips(right)
        if first is not None and int(first) != 0: fmt["indent_first"] = _twips(first)
    return fmt


# ── table parsing ─────────────────────────────────────────────────────────────

def _parse_table(idx: int, node) -> dict:
    rows_data = []
    for row in node.findall(".//" + qn("w:tr")):
        cells = []
        for cell in row.findall(".//" + qn("w:tc")):
            lines = []
            for cp in cell.findall(".//" + qn("w:p")):
                txt = "".join(t.text or "" for t in cp.iter(qn("w:t")))
                runs = []
                for r in cp.iter(qn("w:r")):
                    rPr = r.find(qn("w:rPr"))
                    t2  = "".join(t.text or "" for t in r.iter(qn("w:t")))
                    if not t2:
                        continue
                    run = {"text": t2}
                    if rPr is not None:
                        if rPr.find(qn("w:b")) is not None: run["bold"]   = True
                        if rPr.find(qn("w:i")) is not None: run["italic"] = True
                        sz = rPr.find(qn("w:sz"))
                        if sz is not None:
                            run["font_size"] = int(sz.get(qn("w:val"))) / 2
                    runs.append(run)
                if txt.strip():
                    lines.append({"text": txt, "runs": runs, "fmt": _para_fmt(cp)})
            cells.append(lines)
        rows_data.append(cells)
    return {"idx": idx, "type": "tbl", "rows": rows_data}


# ── helpers ───────────────────────────────────────────────────────────────────

def _twips(val) -> float:
    return round(int(val) / 20, 1)


def _parse_sect_pr(sp) -> dict:
    sb = {}
    typ = sp.find(qn("w:type"))
    if typ is not None:
        sb["type"] = typ.get(qn("w:val"))
    cols = sp.find(qn("w:cols"))
    if cols is not None:
        c = {}
        num   = cols.get(qn("w:num"))
        space = cols.get(qn("w:space"))
        eq    = cols.get(qn("w:equalWidth"))
        if num:   c["num"]        = int(num)
        if space: c["space"]      = int(space)
        if eq:    c["equalWidth"] = eq
        sb["cols"] = c
    pgSz = sp.find(qn("w:pgSz"))
    if pgSz is not None:
        sb["pg_sz"] = {k: pgSz.get(qn(f"w:{k}"))
                       for k in ("w", "h", "orient") if pgSz.get(qn(f"w:{k}"))}
    pgMar = sp.find(qn("w:pgMar"))
    if pgMar is not None:
        sb["pg_mar"] = {k: int(pgMar.get(qn(f"w:{k}")))
                        for k in ("top", "bottom", "left", "right", "header", "footer")
                        if pgMar.get(qn(f"w:{k}"))}
    pgNum = sp.find(qn("w:pgNumType"))
    if pgNum is not None:
        start = pgNum.get(qn("w:start"))
        if start is not None:
            sb["pg_num_start"] = int(start)
    return sb
