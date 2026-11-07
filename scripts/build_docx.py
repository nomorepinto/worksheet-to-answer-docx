#!/usr/bin/env python3
"""Build a plain Word answer sheet (.docx) from a JSON spec.

Enforces Arial 12 everywhere (Normal style + every run + every table cell),
uses only the bordered "Table Grid" style, and bolds table header rows.
No themes, shading, or colors.

Usage:
    python build_docx.py spec.json [output.docx]
    python build_docx.py --help

If output.docx is omitted, the spec's "output" key is used (default Answers.docx).

Spec schema (top-level object):
    {
      "output":   "Answers.docx",          # optional filename
      "title":    "Course - Assignment",   # optional, centered bold
      "subtitle": "Answer Sheet",          # optional, centered
      "blocks":   [ ... ]                   # required list, see block types
    }

Block types:
    {"type": "heading", "text": "..."}                 bold section header
    {"type": "para",    "text": "...", "bold": false,  # normal paragraph (default)
                         "align": "left|center|right", "size": 12}
    {"type": "table",   "rows": [["h1","h2"], ["a","b"]],
                        "header": true}                 first row bold if header
    {"type": "spacer"}                                  blank line

Dependency: python-docx  (pip install python-docx)
"""
import sys
import json

try:
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
except ImportError:
    sys.exit("python-docx is required: python -m pip install python-docx")

FONT = "Arial"
SIZE = Pt(12)

ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
}


def set_normal_font(doc):
    normal = doc.styles["Normal"]
    normal.font.name = FONT
    normal.font.size = SIZE
    rpr = normal.element.get_or_add_rPr()
    rpr.rFonts.set(qn("w:eastAsia"), FONT)


def add_para(doc, text, bold=False, align=None, size=12):
    p = doc.add_paragraph()
    if align in ALIGN:
        p.alignment = ALIGN[align]
    r = p.add_run(text)
    r.bold = bold
    r.font.name = FONT
    r.font.size = Pt(size)
    return p


def add_table(doc, rows, header=True):
    if not rows:
        return
    ncols = len(rows[0])
    t = doc.add_table(rows=0, cols=ncols)
    t.style = "Table Grid"  # borders only, no shading/color
    for i, row in enumerate(rows):
        cells = t.add_row().cells
        for c, val in zip(cells, row):
            c.text = ""
            run = c.paragraphs[0].add_run(str(val))
            run.font.name = FONT
            run.font.size = SIZE
            if header and i == 0:
                run.bold = True
    return t


def build(spec, fallback_output="Answers.docx"):
    doc = Document()
    set_normal_font(doc)

    if spec.get("title"):
        add_para(doc, spec["title"], bold=True, align="center")
    if spec.get("subtitle"):
        add_para(doc, spec["subtitle"], align="center")
    if spec.get("title") or spec.get("subtitle"):
        doc.add_paragraph()

    for b in spec.get("blocks", []):
        btype = b.get("type")
        if btype == "heading":
            add_para(doc, b.get("text", ""), bold=True)
        elif btype == "para":
            add_para(doc, b.get("text", ""), bold=b.get("bold", False),
                     align=b.get("align"), size=b.get("size", 12))
        elif btype == "table":
            add_table(doc, b.get("rows", []), header=b.get("header", True))
        elif btype == "spacer":
            doc.add_paragraph()
        else:
            raise ValueError(f"Unknown block type: {btype!r}")

    return doc, spec.get("output", fallback_output)


def main(argv):
    if len(argv) < 2 or argv[1] in ("-h", "--help"):
        print(__doc__)
        return 0
    with open(argv[1], encoding="utf-8-sig") as f:
        spec = json.load(f)
    doc, out = build(spec)
    if len(argv) > 2:
        out = argv[2]
    try:
        doc.save(out)
    except PermissionError:
        out_tmp = out.replace(".docx", "_tmp.docx")
        doc.save(out_tmp)
        sys.exit(f"{out} is locked (open in Word?). Wrote {out_tmp}; "
                 f"close the file and re-run.")
    print(f"saved {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
