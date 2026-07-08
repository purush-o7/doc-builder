# doc_builder — Quick Guide for Claude Sessions

Three scripts, one pipeline: **extract** a `.docx` into JSON → **edit** the JSON → **build** a new `.docx` → **verify** with the matcher.

```
KovilAR_IEEE_Paper_v9.docx
        │
        ▼  (extractor)
doc_builder/content/
  ├── sections/*.json      ← edit these
  ├── media/*.png/jpg      ← images extracted from docx
  └── paper_v10.json       ← build spec (input/output paths + defaults)
        │
        ▼  (builder)
doc_builder/output/KovilAR_IEEE_Paper_v11.docx
        │
        ▼  (matcher)  optional fidelity check
   PASS / FAIL report
```

---

## 1. The Three Scripts

### 1.1 Extractor — `doc_builder/extractor/`

**When:** You have a `.docx` and want to pull its content into editable JSON files.

```bash
# CLI
python -m doc_builder.extractor KovilAR_IEEE_Paper_v9.docx \
    --out  doc_builder/content/sections \
    --spec doc_builder/content/paper_v10.json \
    --version v10

# Python
from doc_builder.extractor import extract
extract(
    docx_path  = "KovilAR_IEEE_Paper_v9.docx",
    output_dir = "doc_builder/content/sections",
    spec_out   = "doc_builder/content/paper_v10.json",
)
```

**What it produces:**
- One `sections/<slug>.json` per detected section (split on Roman-numeral headings)
- `content/media/rIdX.ext` — every embedded image extracted from the docx
- `paper_vX.json` — a ready-to-use build spec referencing all section files

**Key options:**
| Flag | Default | Meaning |
|------|---------|---------|
| `--out` | `doc_builder/content/sections` | Where to write section JSON files |
| `--spec` | _(none)_ | Path for the build spec JSON |
| `--version` | _(none)_ | Version tag used in spec output filename |
| `--no-overwrite` | _(off)_ | Skip section files that already exist |

---

### 1.2 Builder — `doc_builder/builder.py` + `renderer.py`

**When:** You have edited the JSON sections and want to produce a new `.docx`.

```bash
# CLI
python -m doc_builder doc_builder/content/paper_v10.json

# Python
from doc_builder import build
build("doc_builder/content/paper_v10.json")
```

The spec file (`paper_v10.json`) controls everything — input doc, output path, style defaults, and which section files to combine.

---

### 1.3 Matcher — `doc_builder/matcher.py`

**When:** After a build, to verify the rebuilt doc matches the original paragraph-for-paragraph.

```bash
# CLI
python -m doc_builder.matcher KovilAR_IEEE_Paper_v9.docx \
                              doc_builder/output/KovilAR_IEEE_Paper_v11.docx --verbose

# Python
from doc_builder.matcher import match
report = match("original.docx", "rebuilt.docx")
report.print_summary(verbose=True)
```

Exit code `0` = PASS, `1` = FAIL. Checks: paragraph count, table count, text content, alignment, bold/italic/font_size per run, table cell content.

---

## 2. Build Spec — `paper_v10.json`

```json
{
  "meta": {
    "input":  "../../KovilAR_IEEE_Paper_v9.docx",
    "output": "../output/KovilAR_IEEE_Paper_v11.docx"
  },
  "defaults": {
    "body_align":        "justify",
    "table_font_size":   8,
    "table_header_bold": true,
    "table_borders":     true,
    "table_border_sz":   "4",
    "table_border_color":"000000",
    "table_cell_align":  "center",
    "table_style":       "TableNormal",
    "figure_width":      3.4,
    "figure_align":      "center",
    "caption_align":     "center",
    "font_name":         null,
    "font_size":         null
  },
  "operations": [
    {
      "op":      "rebuild_doc",
      "sections": [
        "sections/header.json",
        "sections/introduction.json",
        "sections/related_work.json",
        "sections/kovil_ar_design_and_implementation.json",
        "sections/kovilar_design.json",
        "sections/evaluation.json",
        "sections/conclusion.json",
        "sections/references.json"
      ]
    }
  ]
}
```

**Paths** in `meta` are relative to the spec file's directory (`content/`).  
**Paths** in `sections` are also relative to the spec file's directory.

### Operations

| `op` | What it does |
|------|-------------|
| `rebuild_doc` | Clears the entire document body and rebuilds from listed section files |
| `replace_section` | Replaces content between two headings; needs `start`, `stop`, `content_file` |
| `insert_after` | Inserts content after a specific heading; needs `after`, `content_file` |
| `append_section` | Appends content to the end of the document; needs `content_file` |

---

## 3. Section JSON — Element Reference

Every section file is a JSON **array** of element objects. Each element has a `"type"` field.

---

### 3.1 `content` — the universal text element

Has a `"style"` sub-field that selects the render mode.

#### style: `heading`

Bold, centered (or explicitly aligned). Used for Roman-numeral section titles.

```json
{
  "type": "content",
  "style": "heading",
  "text": "III. SYSTEM DESIGN AND MOOVAR KOVIL DIGITISATION PIPELINE",
  "align": "center",
  "bold_italic": true,
  "font_size": 10
}
```

| Key | Type | Notes |
|-----|------|-------|
| `text` | string | Full heading text |
| `align` | string\|null | `"center"`, `"left"`, `"right"`, `"both"` (=justify), `null` (inherit) |
| `bold_italic` | bool | Adds italic on top of bold |
| `font_size` | number | Points; omit to inherit |
| `fmt` | object | Paragraph spacing — see §3.7 |
| `section_break` | object | Inline sectPr — see §3.8 |

---

#### style: `subsection`

Italic (not bold) subsection label (e.g. "A. Phase 1: Reader Tab").

```json
{
  "type": "content",
  "style": "subsection",
  "text": "A. Phase 1: Reader Tab",
  "align": null
}
```

| Key | Type | Notes |
|-----|------|-------|
| `text` | string | Subsection label |
| `align` | string\|null | Usually `null` (left, inherited) |
| `italic` | bool | Default true; set false to suppress |
| `bold` | bool | Normally absent |
| `font_size` | number | Points |

---

#### style: `para`

Plain body paragraph — single uniform run.

```json
{
  "type": "content",
  "style": "para",
  "text": "The Microsoft HoloLens 2 is a standalone optical see-through...",
  "align": "both",
  "fmt": { "indent_first": 36.0 }
}
```

| Key | Type | Notes |
|-----|------|-------|
| `text` | string | Full paragraph text |
| `align` | string\|null | Usually `"both"` (justify) for body text |
| `font_size` | number | Points; omit for default |
| `fmt` | object | Spacing/indent — see §3.7 |

---

#### style: `para_rich`

Mixed-format paragraph with multiple runs (bold, italic, different fonts, etc.).

```json
{
  "type": "content",
  "style": "para_rich",
  "runs": [
    { "text": "The software stack comprises four layers. At the base, " },
    { "text": "Unity 2020.3 LTS", "bold": true },
    { "text": " provides the scene graph and rendering pipeline." }
  ],
  "align": "both"
}
```

Each **run** object:

| Key | Type | Notes |
|-----|------|-------|
| `text` | string | Run text content |
| `bold` | bool | Bold formatting |
| `italic` | bool | Italic formatting |
| `underline` | bool | Underline formatting |
| `font_size` | number | Points; overrides paragraph default |
| `font_name` | string | e.g. `"Times New Roman"`, `"Cardo"` |
| `color` | string | Hex RGB without `#`, e.g. `"FF0000"` |

---

#### style: `table`

A bordered data table with an optional caption above it.

```json
{
  "type": "content",
  "style": "table",
  "caption": "TABLE VI: User Experience Evaluation Results",
  "headers": ["Metric", "Mean", "SD", "Positive %"],
  "rows": [
    ["Interface Usability", "4.18", "0.60", "89.8%"],
    ["Revisit Intention",   "4.82", "0.39", "100%"]
  ],
  "caption_fmt": { "space_before": 6.0, "space_after": 6.0 },
  "caption_italic": false,
  "col_widths": [1.5, 0.6, 0.6, 0.8],
  "font_size": 8,
  "header_bold": true,
  "cell_align": "center"
}
```

| Key | Type | Notes |
|-----|------|-------|
| `caption` | string | Caption text; rendered bold above the table |
| `caption_runs` | array | Rich-text caption runs (overrides `caption` string) |
| `caption_fmt` | object | Paragraph fmt for the caption paragraph |
| `caption_italic` | bool | Makes caption italic |
| `headers` | array of strings | First row — rendered bold by default |
| `rows` | array of arrays | Data rows; each cell is a string |
| `col_widths` | array of numbers | Column widths in inches (optional) |
| `font_size` | number | Overrides `table_font_size` default |
| `header_bold` | bool | Overrides `table_header_bold` default |
| `cell_align` | string | Overrides `table_cell_align` default |

---

### 3.2 `figure`

An image paragraph. Caption is a **separate** `figure_caption` element that follows it.

```json
{
  "type": "figure",
  "path": "media/rId8.png",
  "width": 3.474,
  "align": "center",
  "fmt": { "space_before": 12.0, "space_after": 6.0 }
}
```

| Key | Type | Notes |
|-----|------|-------|
| `path` | string | Relative to the spec file directory (`content/`) |
| `width` | number | Width in inches; defaults to `figure_width` in defaults |
| `align` | string\|null | Usually `"center"` |
| `fmt` | object | Paragraph spacing — see §3.7 |

> Images live in `content/media/`. New images should be copied there before referencing.

---

### 3.3 `figure_caption`

A centered caption paragraph for the preceding figure. Always placed **immediately after** its `figure` element.

```json
{
  "type": "figure_caption",
  "runs": [
    { "text": "Fig. 2. ", "bold": true },
    { "text": "Bottom-tier elements of Moovar Kovil: Adhisthana and Pada." }
  ],
  "align": "center",
  "fmt": {}
}
```

| Key | Type | Notes |
|-----|------|-------|
| `runs` | array | Same run format as `para_rich` |
| `align` | string\|null | Defaults to `caption_align` (`"center"`) |
| `fmt` | object | Paragraph spacing |

> Convention: first run = `"Fig. N. "` bold; second run = description, no bold.

---

### 3.4 `authors_table`

Special two-column author block for the IEEE header. Generated by the extractor from the first table in the document. **Do not hand-write** — edit via extractor or directly in `header.json`.

```json
{
  "type": "authors_table",
  "authors": [
    {
      "lines": ["Author Name", "Institution", "City, Country", "email@domain.com"],
      "runs":  [ [...runs for line 1...], [...runs for line 2...] ]
    }
  ]
}
```

---

### 3.5 `spacer`

An empty paragraph. Preserves vertical spacing exactly as in the original document.

```json
{ "type": "spacer", "align": "both" }
{ "type": "spacer", "align": null }
```

| `align` value | Effect |
|--------------|--------|
| `"both"` / `"justify"` | Justify alignment on empty para |
| `"center"` | Center-aligned empty para |
| `null` | No explicit alignment (inherit) |

> Do not collapse or remove spacers — they control exact vertical spacing.

---

### 3.6 `placeholder`

Red bold text used to mark sections that need content. Visible in the output document as a red paragraph.

```json
{
  "type": "placeholder",
  "text": "B.  …………?.........",
  "align": "both"
}
```

---

### 3.7 `fmt` — Paragraph Formatting

Any element that accepts a `"fmt"` key uses this structure. All values are in **points** unless noted.

```json
"fmt": {
  "space_before":  6.0,
  "space_after":   6.0,
  "line_spacing":  240,
  "line_rule":     "auto",
  "indent_left":   36.0,
  "indent_right":  0.0,
  "indent_first":  36.0
}
```

| Key | Unit | Notes |
|-----|------|-------|
| `space_before` | points | Space above paragraph |
| `space_after` | points | Space below paragraph |
| `line_spacing` | twips | 240 = single, 360 = 1.5×, 480 = double |
| `line_rule` | string | `"auto"`, `"exact"`, `"atLeast"` |
| `indent_left` | points | Left indent of entire paragraph |
| `indent_right` | points | Right indent |
| `indent_first` | points | First-line indent (body text typically 36pt) |

---

### 3.8 `section_break` — Page Layout Zones

Embedded inside `heading` or `spacer` elements. When present, injects an inline `<w:sectPr>` that ends a layout section (e.g. switches from 1-column to 2-column layout).

```json
"section_break": {
  "type": "continuous",
  "cols": { "num": 1, "space": 708, "equalWidth": "1" },
  "pg_sz": { "w": "12240", "h": "15840" },
  "pg_mar": { "top": 1080, "bottom": 1080, "left": 1080, "right": 1080, "header": 720, "footer": 720 },
  "pg_num_start": 1
}
```

> In this paper, `header.json` carries the `section_break` that separates the full-width title/author block (section 1) from the 2-column body (section 2). **Do not remove it.**

---

## 4. Align Values Reference

| JSON value | Word alignment |
|-----------|----------------|
| `"left"` | Left |
| `"center"` | Center |
| `"right"` | Right |
| `"both"` | Justify (OOXML standard value) |
| `"justify"` | Justify (alias, also accepted) |
| `null` | No explicit alignment — inherits from style |

> OOXML uses `"both"` for justify. The renderer maps both `"both"` and `"justify"` to `WD_ALIGN_PARAGRAPH.JUSTIFY`.

---

## 5. Current Paper Section Files

| File | Section heading | Contains |
|------|----------------|---------|
| `header.json` | _(no heading)_ | Title, abstract, index terms, authors table, section break |
| `introduction.json` | `I. INTRODUCTION` | Intro paragraphs, references |
| `related_work.json` | `II. RELATED WORK` | Related work paragraphs, Table I, Table II |
| `kovil_ar_design_and_implementation.json` | `III. SYSTEM DESIGN AND MOOVAR KOVIL DIGITISATION PIPELINE` | Hardware, software stack, architectural elements (Fig. 2, Fig. 3), pipeline (Fig. 1) |
| `kovilar_design.json` | `IV. KovilAR Design` | Four phases (Fig. 3–7), system architecture (Fig. 8), reset mechanism |
| `evaluation.json` | `VI. EVALUATION` | User study, results (Fig. 9, Fig. 10, Fig. 12), statistical analysis |
| `conclusion.json` | _(conclusion heading)_ | Conclusion paragraphs |
| `references.json` | `REFERENCES` | Reference list |

---

## 6. Common Workflows for Claude

### Add a new figure
1. Copy image to `doc_builder/content/media/<name>.png`
2. In the target section JSON, add:
   ```json
   { "type": "figure", "path": "media/<name>.png", "width": 3.4, "align": "center" },
   { "type": "figure_caption", "runs": [{"text": "Fig. N. ", "bold": true}, {"text": "Short caption."}], "align": "center" }
   ```
3. Run builder: `python -m doc_builder doc_builder/content/paper_v10.json`

### Replace a table with figures
Remove the `{"type": "content", "style": "table", ...}` element and insert `figure` + `figure_caption` pairs in its place. Update any paragraph that references the table to reference the figure numbers instead.

### Edit a section heading
Find the element with `"style": "heading"` in the relevant section JSON and update `"text"`. The filename slug does not need to change.

### Change caption text
Find the `figure_caption` element. Edit the second run's `"text"` (the first run is always the bold `"Fig. N. "` label).

### Add a subsection
Insert a `{"type": "content", "style": "subsection", "text": "X. Subsection Title", "align": null}` element at the right position in the section JSON, followed by a `spacer`.

### Edit table data
Find the `{"type": "content", "style": "table"}` element. Edit `"headers"` (array of strings) and `"rows"` (array of arrays of strings). Caption is the `"caption"` key.

### Rebuild after edits
```bash
python -m doc_builder doc_builder/content/paper_v10.json
# Output: doc_builder/output/KovilAR_IEEE_Paper_v11.docx
```
Bump the output version in `paper_v10.json` (`"output": "../output/KovilAR_IEEE_Paper_vNN.docx"`) when producing a new release version.

### Verify fidelity against original
```bash
python -m doc_builder.matcher KovilAR_IEEE_Paper_v9.docx \
    doc_builder/output/KovilAR_IEEE_Paper_v11.docx --verbose
```
