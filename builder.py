"""
DocBuilder — loads a build spec JSON and produces an edited Word document.

Operations
----------
replace_section   Replace body elements between start/stop headings (inclusive/exclusive).
insert_after      Insert content immediately after a heading.
append_section    Append content to end of document.
rebuild_doc       Clear all body content, rebuild from listed section files.

Spec format
-----------
{
  "meta": {
    "input":  "../../KovilAR_IEEE_Paper_v9.docx",   // relative to this JSON
    "output": "output/KovilAR_IEEE_Paper_v10.docx"  // relative to this JSON
  },
  "defaults": { ... },          // optional — overrides GLOBAL_DEFAULTS
  "operations": [
    {
      "op":      "rebuild_doc",
      "sections": [             // paths relative to this JSON
        "sections/header.json",
        "sections/introduction.json"
      ]
    },
    {
      "op":           "replace_section",
      "start":        "V. EVALUATION",
      "stop":         "VI. CONCLUSION",
      "content_file": "sections/evaluation.json"   // OR inline "content": [...]
    }
  ]
}
"""

import json
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn

from .renderer import Renderer
from .style import GLOBAL_DEFAULTS, merge


class DocBuilder:
    def __init__(self, json_path: str):
        self.json_path = Path(json_path).resolve()
        self.json_dir  = self.json_path.parent

        with open(self.json_path, encoding="utf-8") as f:
            spec = json.load(f)

        self.meta     = spec["meta"]
        self.defaults = merge(GLOBAL_DEFAULTS, spec.get("defaults") or {})
        self.ops      = [self._resolve_op(op) for op in spec.get("operations", [])]

    # ── public ────────────────────────────────────────────────────────────────

    def build(self):
        input_path  = (self.json_dir / self.meta["input"]).resolve()
        output_path = (self.json_dir / self.meta["output"]).resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)

        doc      = Document(str(input_path))
        renderer = Renderer(doc, self.defaults, base_dir=self.json_dir)

        for op in self.ops:
            op_defaults = merge(self.defaults, op.get("defaults") or {})
            renderer.defaults = op_defaults

            name = op["op"]
            if name == "rebuild_doc":
                self._rebuild(doc, renderer, op)
            elif name == "replace_section":
                self._replace(doc, renderer, op)
            elif name == "append_section":
                self._append(doc, renderer, op)
            elif name == "insert_after":
                self._insert_after(doc, renderer, op)
            else:
                raise ValueError(f"Unknown operation: {name!r}")

        renderer.defaults = self.defaults
        doc.save(str(output_path))
        print(f"Saved: {output_path}")

    # ── operations ────────────────────────────────────────────────────────────

    def _rebuild(self, doc, renderer, op):
        body = doc.element.body
        # remove all body children EXCEPT sectPr (needed for page-width calc during render)
        sectPr = body.find(qn("w:sectPr"))
        for child in list(body):
            if child is not sectPr:
                body.remove(child)

        all_content = []
        for sec_file in op["sections"]:
            path = (self.json_dir / sec_file).resolve()
            with open(path, encoding="utf-8") as f:
                all_content.extend(json.load(f))

        new_els = renderer.render_content(all_content)
        # insert before sectPr so it stays last
        for i, el in enumerate(new_els):
            body.insert(i, el)

        print(f"  rebuild_doc: {len(op['sections'])} sections, {len(new_els)} elements")

    def _replace(self, doc, renderer, op):
        body     = doc.element.body
        children = list(body)
        start    = self._find_heading(children, op["start"])
        stop     = self._find_heading(children, op["stop"])

        if start is None:
            raise ValueError(f"Heading not found: {op['start']!r}")
        if stop is None:
            raise ValueError(f"Heading not found: {op['stop']!r}")

        print(f"  replace_section [{op['start']!r}] body[{start}:{stop}]  ({stop-start} removed)")
        for el in children[start:stop]:
            body.remove(el)

        new_els = renderer.render_content(op["content"])
        for i, el in enumerate(new_els):
            body.insert(start + i, el)
        print(f"  inserted {len(new_els)} elements at {start}")

    def _append(self, doc, renderer, op):
        new_els = renderer.render_content(op["content"])
        for el in new_els:
            doc.element.body.append(el)
        print(f"  append_section: {len(new_els)} elements")

    def _insert_after(self, doc, renderer, op):
        body     = doc.element.body
        children = list(body)
        idx      = self._find_heading(children, op["after"])
        if idx is None:
            raise ValueError(f"Heading not found: {op['after']!r}")
        new_els = renderer.render_content(op["content"])
        for i, el in enumerate(new_els):
            body.insert(idx + 1 + i, el)
        print(f"  insert_after {op['after']!r}: {len(new_els)} elements")

    # ── helpers ───────────────────────────────────────────────────────────────

    def _resolve_op(self, op: dict) -> dict:
        if "content_file" in op:
            cf = (self.json_dir / op["content_file"]).resolve()
            with open(cf, encoding="utf-8") as f:
                return {**op, "content": json.load(f)}
        if "sections" in op:
            return op  # loaded lazily in _rebuild
        return op

    @staticmethod
    def _find_heading(children, text: str):
        for idx, child in enumerate(children):
            if child.tag == qn("w:p"):
                full = "".join(n.text or "" for n in child.iter(qn("w:t")))
                if text in full:
                    return idx
        return None
