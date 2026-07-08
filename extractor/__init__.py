"""
doc_builder.extractor — extract content from a .docx into section JSON files.

Usage (Python)
--------------
from doc_builder.extractor import extract

extract(
    docx_path  = "KovilAR_IEEE_Paper_v9.docx",
    output_dir = "doc_builder/content/sections",
    spec_out   = "doc_builder/content/paper_v10.json",   # optional
    input_ref  = "../../KovilAR_IEEE_Paper_v9.docx",     # optional, for spec
    output_ref = "output/KovilAR_IEEE_Paper_v10.docx",   # optional, for spec
)

Usage (CLI)
-----------
python -m doc_builder.extractor KovilAR_IEEE_Paper_v9.docx \\
    --out     doc_builder/content/sections  \\
    --spec    doc_builder/content/paper_v10.json \\
    --version v10
"""

import json
import sys
from pathlib import Path

from .parser     import parse
from .classifier import classify_and_group


def extract(
    docx_path:  str,
    output_dir: str,
    spec_out:   str  = None,
    input_ref:  str  = None,
    output_ref: str  = None,
    version:    str  = None,
    section_map: dict = None,
    overwrite:  bool  = True,
):
    """
    Parse *docx_path*, classify content, write one JSON per section to
    *output_dir*, and optionally write a rebuild spec to *spec_out*.

    Parameters
    ----------
    docx_path   : path to the source .docx
    output_dir  : directory to write section JSON files into
    spec_out    : if given, write a paper_vX.json spec referencing the sections
    input_ref   : path to put in spec meta.input  (relative to spec_out dir)
    output_ref  : path to put in spec meta.output (relative to spec_out dir)
    version     : version tag for the output filename in the spec (e.g. "v10")
    section_map : optional {heading_text: filename_stem} overrides
    overwrite   : overwrite existing section files (default True)
    """
    docx_path  = Path(docx_path).resolve()
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    # media dir sits next to sections dir  (content/media/)
    media_dir = output_dir.parent / "media"
    media_dir.mkdir(parents=True, exist_ok=True)

    print(f"Parsing: {docx_path.name}")
    raw, image_map = parse(str(docx_path), str(media_dir))
    print(f"  {len(raw)} raw body elements  |  {len(image_map)} images extracted to {media_dir.name}/")

    # spec_dir is output_dir's parent (content/)  so image paths are relative to it
    spec_dir = output_dir.parent
    sections = classify_and_group(raw, section_map, spec_dir=spec_dir,
                                  first_table_is_authors=True)
    print(f"  {len(sections)} sections detected: {list(sections.keys())}")

    written = []
    for name, elements in sections.items():
        filename = f"{name}.json"
        dest = output_dir / filename
        if dest.exists() and not overwrite:
            print(f"  skip (exists): {filename}")
            written.append(filename)
            continue
        with open(dest, "w", encoding="utf-8") as f:
            json.dump(elements, f, ensure_ascii=False, indent=2)
        print(f"  wrote: {filename}  ({len(elements)} elements)")
        written.append(filename)

    if spec_out:
        _write_spec(
            spec_path  = Path(spec_out).resolve(),
            sections   = written,
            sections_rel_dir = output_dir,
            input_ref  = input_ref  or f"../../{docx_path.name}",
            output_ref = output_ref or _default_output_ref(docx_path, version),
        )

    print("Done.")
    return sections


def _write_spec(spec_path, sections, sections_rel_dir, input_ref, output_ref):
    spec_dir = spec_path.parent
    sec_rel  = [
        str(Path("sections") / s).replace("\\", "/")
        for s in sections
    ]
    spec = {
        "meta": {
            "input":  input_ref,
            "output": output_ref,
        },
        "defaults": {
            "body_align":        "justify",
            "table_font_size":   8,
            "table_header_bold": True,
            "table_borders":     True,
            "figure_width":      3.4,
            "figure_align":      "center",
            "caption_align":     "center",
        },
        "operations": [
            {
                "op":      "rebuild_doc",
                "sections": sec_rel,
            }
        ],
    }
    spec_path.parent.mkdir(parents=True, exist_ok=True)
    with open(spec_path, "w", encoding="utf-8") as f:
        json.dump(spec, f, ensure_ascii=False, indent=2)
    print(f"  spec:  {spec_path}")


def _default_output_ref(docx_path: Path, version: str) -> str:
    stem = docx_path.stem
    # strip trailing version like _v9, _v8
    stem = stem.rstrip("0123456789").rstrip("v").rstrip("_")
    tag  = version or "v_next"
    return f"../output/{stem}_{tag}.docx"


# ── CLI ───────────────────────────────────────────────────────────────────────

def _cli():
    import argparse
    p = argparse.ArgumentParser(
        prog="python -m doc_builder.extractor",
        description="Extract a .docx into section JSON files for doc_builder.",
    )
    p.add_argument("docx",              help="Source .docx file")
    p.add_argument("--out",  "-o",      default="doc_builder/content/sections",
                   help="Output directory for section JSON files")
    p.add_argument("--spec", "-s",      default=None,
                   help="Path to write the rebuild spec JSON (e.g. doc_builder/content/paper_v10.json)")
    p.add_argument("--version",         default=None,
                   help="Version tag for output filename in spec (e.g. v10)")
    p.add_argument("--no-overwrite",    action="store_true",
                   help="Skip section files that already exist")
    args = p.parse_args()

    extract(
        docx_path  = args.docx,
        output_dir = args.out,
        spec_out   = args.spec,
        version    = args.version,
        overwrite  = not args.no_overwrite,
    )


if __name__ == "__main__":
    _cli()
