"""
matcher.py — compare an original .docx with a rebuilt one and report
every difference that matters for fidelity.

Usage (Python)
--------------
from doc_builder.matcher import match
report = match("original.docx", "rebuilt.docx")
report.print_summary()

Usage (CLI)
-----------
python -m doc_builder.matcher original.docx rebuilt.docx
python -m doc_builder.matcher original.docx rebuilt.docx --verbose
"""

import sys
from dataclasses import dataclass, field
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn


# ─────────────────────────────────────────────────────────────────────────────
# Data structures
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class Diff:
    location: str          # e.g. "para[5]" or "table[2].row[1].col[0]"
    field:    str          # what differs: "text", "align", "space_after", ...
    original: object
    rebuilt:  object

    def __str__(self):
        return f"  [{self.location}] {self.field}: {self.original!r} → {self.rebuilt!r}"


@dataclass
class MatchReport:
    original_path: str
    rebuilt_path:  str
    diffs:         list = field(default_factory=list)
    para_count:    tuple = (0, 0)   # (original, rebuilt)
    table_count:   tuple = (0, 0)

    # ── categorised diff lists ────────────────────────────────────────────────
    @property
    def text_diffs(self):
        return [d for d in self.diffs if d.field == "text"]

    @property
    def fmt_diffs(self):
        return [d for d in self.diffs if d.field not in ("text", "table_rows",
                                                          "table_cols", "cell_text")]

    @property
    def table_diffs(self):
        return [d for d in self.diffs if "table" in d.location]

    @property
    def match_pct(self) -> float:
        total = max(self.para_count[0], 1) + max(self.table_count[0], 1)
        diff_items = len(set(d.location for d in self.diffs))
        return max(0.0, 100 * (1 - diff_items / total))

    def print_summary(self, verbose: bool = False):
        sep = "-" * 60
        print(sep)
        print("MATCH REPORT")
        print(f"  original : {self.original_path}")
        print(f"  rebuilt  : {self.rebuilt_path}")
        print(sep)
        cnt_diff = self.para_count[0] - self.para_count[1]
        p_ok = "OK" if cnt_diff == 0 else f"*** {abs(cnt_diff)} {'missing' if cnt_diff>0 else 'extra'} ***"
        print(f"  paragraphs  original={self.para_count[0]}  rebuilt={self.para_count[1]}  {p_ok}")
        tdiff = self.table_count[0] - self.table_count[1]
        t_ok = "OK" if tdiff == 0 else f"*** {abs(tdiff)} {'missing' if tdiff>0 else 'extra'} ***"
        print(f"  tables      original={self.table_count[0]}  rebuilt={self.table_count[1]}  {t_ok}")

        categories = [
            ("Text mismatches",      self.text_diffs),
            ("Formatting mismatches", self.fmt_diffs),
            ("Table mismatches",      self.table_diffs),
        ]
        for label, dl in categories:
            if dl:
                print(f"\n  {label} ({len(dl)}):")
                for d in dl[:20]:
                    print(str(d))
                if len(dl) > 20:
                    print(f"    ... and {len(dl)-20} more")
            elif verbose:
                print(f"\n  {label}: none")

        print(sep)
        status = "PASS" if not self.diffs and cnt_diff == 0 and tdiff == 0 else "FAIL"
        print(f"  Overall match: {self.match_pct:.1f}%   [{status}]")
        print(sep)
        print(sep)


# ─────────────────────────────────────────────────────────────────────────────
# Main comparison logic
# ─────────────────────────────────────────────────────────────────────────────

def match(original_path: str, rebuilt_path: str, verbose: bool = False) -> MatchReport:
    orig = _load(original_path)
    rblt = _load(rebuilt_path)

    report = MatchReport(
        original_path=str(original_path),
        rebuilt_path=str(rebuilt_path),
        para_count=(len(orig["paras"]), len(rblt["paras"])),
        table_count=(len(orig["tables"]), len(rblt["tables"])),
    )

    _compare_paras(orig["paras"], rblt["paras"], report)
    _compare_tables(orig["tables"], rblt["tables"], report)
    return report


# ─────────────────────────────────────────────────────────────────────────────
# Extraction helpers
# ─────────────────────────────────────────────────────────────────────────────

def _load(path: str) -> dict:
    doc  = Document(str(path))
    body = doc.element.body
    paras, tables = [], []

    for child in body:
        tag = child.tag.split("}")[1] if "}" in child.tag else child.tag
        if tag == "p":
            paras.append(_read_para(child))
        elif tag == "tbl":
            tables.append(_read_table(child))

    return {"paras": paras, "tables": tables}


def _read_para(node) -> dict:
    runs = []
    for r in node.iter(qn("w:r")):
        rPr  = r.find(qn("w:rPr"))
        txt  = "".join(t.text or "" for t in r.iter(qn("w:t")))
        if not txt:
            continue
        run = {"text": txt}
        if rPr is not None:
            if rPr.find(qn("w:b"))  is not None: run["bold"]   = True
            if rPr.find(qn("w:i"))  is not None: run["italic"] = True
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                run["font_size"] = round(int(sz.get(qn("w:val"))) / 2, 1)
        runs.append(run)

    fmt = {}
    pPr = node.find(qn("w:pPr"))
    if pPr is not None:
        jc = pPr.find(qn("w:jc"))
        if jc is not None:
            fmt["align"] = jc.get(qn("w:val"))
        sp = pPr.find(qn("w:spacing"))
        if sp is not None:
            for attr, key in [("w:before", "space_before"), ("w:after", "space_after")]:
                v = sp.get(qn(attr))
                if v is not None:
                    fmt[key] = round(int(v) / 20, 1)

    return {
        "text": "".join(r["text"] for r in runs),
        "runs": runs,
        "fmt":  fmt,
    }


def _read_table(node) -> dict:
    rows = []
    for tr in node.findall(".//" + qn("w:tr")):
        cells = []
        for tc in tr.findall(".//" + qn("w:tc")):
            cells.append("".join(t.text or "" for t in tc.iter(qn("w:t"))))
        rows.append(cells)
    return {"rows": rows}


# ─────────────────────────────────────────────────────────────────────────────
# Comparators
# ─────────────────────────────────────────────────────────────────────────────

def _compare_paras(orig: list, rblt: list, report: MatchReport):
    limit = min(len(orig), len(rblt))
    for i in range(limit):
        o, r = orig[i], rblt[i]
        loc  = f"para[{i}]"

        # text
        if _norm(o["text"]) != _norm(r["text"]):
            report.diffs.append(Diff(loc, "text", _snip(o["text"]), _snip(r["text"])))
            continue   # skip fmt check when text differs — misaligned para

        # alignment
        oa = o["fmt"].get("align")
        ra = r["fmt"].get("align")
        if oa != ra:
            report.diffs.append(Diff(loc, "align", oa, ra))

        # space_before / space_after
        for key in ("space_before", "space_after"):
            ov = o["fmt"].get(key)
            rv = r["fmt"].get(key)
            if _fmt_differs(ov, rv):
                report.diffs.append(Diff(loc, key, ov, rv))

        # run-level font sizes (where explicitly set)
        _compare_runs(o["runs"], r["runs"], loc, report)


def _compare_runs(orig_runs, rblt_runs, loc, report):
    limit = min(len(orig_runs), len(rblt_runs))
    for i in range(limit):
        o, r = orig_runs[i], rblt_runs[i]
        # bold / italic
        if o.get("bold") != r.get("bold"):
            report.diffs.append(Diff(f"{loc}.run[{i}]", "bold",
                                     o.get("bold"), r.get("bold")))
        if o.get("italic") != r.get("italic"):
            report.diffs.append(Diff(f"{loc}.run[{i}]", "italic",
                                     o.get("italic"), r.get("italic")))
        # font size (only when one or both have it)
        if o.get("font_size") or r.get("font_size"):
            if _fmt_differs(o.get("font_size"), r.get("font_size"), tol=0.5):
                report.diffs.append(Diff(f"{loc}.run[{i}]", "font_size",
                                         o.get("font_size"), r.get("font_size")))


def _compare_tables(orig: list, rblt: list, report: MatchReport):
    limit = min(len(orig), len(rblt))
    for ti in range(limit):
        o, r = orig[ti], rblt[ti]
        loc  = f"table[{ti}]"

        if len(o["rows"]) != len(r["rows"]):
            report.diffs.append(Diff(loc, "table_rows",
                                     len(o["rows"]), len(r["rows"])))
            continue

        for ri, (orow, rrow) in enumerate(zip(o["rows"], r["rows"])):
            if len(orow) != len(rrow):
                report.diffs.append(Diff(f"{loc}.row[{ri}]", "table_cols",
                                         len(orow), len(rrow)))
                continue
            for ci, (oc, rc) in enumerate(zip(orow, rrow)):
                if _norm(oc) != _norm(rc):
                    report.diffs.append(Diff(f"{loc}.row[{ri}].col[{ci}]",
                                             "cell_text", _snip(oc), _snip(rc)))


# ─────────────────────────────────────────────────────────────────────────────
# Utilities
# ─────────────────────────────────────────────────────────────────────────────

def _norm(text: str) -> str:
    return " ".join(text.split()).strip()


def _snip(text: str, n: int = 60) -> str:
    t = text.strip()
    return t[:n] + "…" if len(t) > n else t


def _fmt_differs(a, b, tol: float = 0.0) -> bool:
    """Return True if a and b differ beyond tolerance."""
    if a is None and b is None:
        return False
    if a is None or b is None:
        # treat None vs 0 as same (both mean "not set")
        return not (a in (None, 0.0) and b in (None, 0.0))
    return abs(float(a) - float(b)) > tol


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def _cli():
    import argparse, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    p = argparse.ArgumentParser(
        prog="python -m doc_builder.matcher",
        description="Compare original and rebuilt .docx files for fidelity.",
    )
    p.add_argument("original", help="Original .docx")
    p.add_argument("rebuilt",  help="Rebuilt .docx")
    p.add_argument("--verbose", "-v", action="store_true",
                   help="Show all categories even when no diffs")
    args = p.parse_args()
    report = match(args.original, args.rebuilt)
    report.print_summary(verbose=args.verbose)
    sys.exit(0 if not report.diffs else 1)


if __name__ == "__main__":
    _cli()
