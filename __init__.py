"""
doc_builder — IEEE Word document builder.

Quick start
-----------
  from doc_builder import build
  build("doc_builder/content/paper_v10.json")

CLI
---
  python -m doc_builder doc_builder/content/paper_v10.json
"""

from .builder import DocBuilder


def build(json_path: str):
    """Load spec from json_path and produce the output Word document."""
    DocBuilder(json_path).build()
