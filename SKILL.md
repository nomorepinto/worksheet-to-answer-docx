---
name: worksheet-to-answer-docx
description: Turn a questions/worksheet file (homework, lab sheets, exams) into a plain Microsoft Word answer file (.docx, Arial 12, no designs/colors). Classifies each item as objective vs. personal, asks the user clarifying questions for anything requiring personal input (measurements, opinions, choices), verifies every computation before writing, then generates the document. Use when the user supplies a questions file and wants answered in a Word document, or mentions "answer the worksheet/questions", "generate answers", "lab sheet", or an .txt/.docx/.pdf of questions to be filled in.
---

# Worksheet to Answer DOCX

Given a file of questions, produce a complete answer sheet as a plain Word document: **Arial 12, no themes or colors, simple bordered tables only**. Never invent personal answers — ask the user.

## Workflow

Copy this checklist and track progress:

```
- [ ] 1. Read and parse the questions file
- [ ] 2. Classify each item: objective (computable/known) vs. personal (needs user input)
- [ ] 3. Ask the user for all personal inputs in one batch (AskUserQuestion)
- [ ] 4. Verify every numeric answer with a quick script before writing
- [ ] 5. Build a JSON spec describing the answer document
- [ ] 6. Generate the .docx with scripts/build_docx.py
- [ ] 7. Verify the output (font, table count, sections) and clean up temp files
```

### 1. Read and parse

Read the questions file. Source text from PDFs is often mangled (missing operators, duplicated values, stray labels like `17\frac{1}{7}`). Reconstruct intent from context — e.g. a "multiply by 7, error from 1" activity means the value is 1/7. **State your interpretation to the user and flag ambiguity** rather than silently guessing.

### 2. Classify items

- **Objective** — facts, definitions, and computations with a single correct result. Answer directly.
- **Personal** — measurements the user must take, first-person reflection/conclusion, or "pick a scenario/example" choices. **Do not fabricate a specific value**; ask. If the user opts out, give a clearly-labeled model answer.

### 3. Ask for personal inputs

Use **one** AskUserQuestion call bundling all personal items (up to 4 questions). Offer sensible concrete options plus a "leave blank / use formulas" escape. Typical asks: a measured value, which real-world example to use, how strongly to match a personal writing style.

### 4. Verify computations

Before writing anything, compute all table values with a short script (Python). Use exact decimal handling for truncation/rounding:

```python
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP
def trunc(x, n): return Decimal(str(x)).quantize(Decimal("1."+"0"*n), rounding=ROUND_DOWN)
def round_(x, n): return Decimal(str(x)).quantize(Decimal("1."+"0"*n), rounding=ROUND_HALF_UP)
```

Show the user the verified numbers. Accuracy first — a wrong value in an answer sheet is worse than a blank.

### 5. Build the JSON spec

Describe the document as blocks (schema in [SPEC.md](SPEC.md)):

```json
{
  "output": "Answers.docx",
  "title": "Course - Assignment Title",
  "subtitle": "Answer Sheet",
  "blocks": [
    {"type": "heading", "text": "PART A. ..."},
    {"type": "para", "text": "Measured diameter d = 2.40 cm."},
    {"type": "table", "rows": [["Item","Value"], ["3.0","7.2"]]},
    {"type": "spacer"}
  ]
}
```

### 6. Generate the document

```bash
python scripts/build_docx.py spec.json
```

The script enforces Arial 12 on the Normal style and every run/table cell, uses only the `Table Grid` border style (no shading/colors), and bolds header rows. This is the required "plain, no designs" output.

### 7. Verify and clean up

Re-open the output and confirm font, section list, and table count; print a table to sanity-check values. Delete the temp spec and any fallback files afterward.

## Handling a locked output file

If saving raises `PermissionError: [Errno 13] Permission denied`, the .docx is **open in Word**. Do not retry silently: generate to a `_tmp.docx` to prove the content is correct, ask the user to close the file, then save the real name and delete the temp.

## Dependency

`python-docx` (test with `python -c "import docx"`; install with `python -m pip install python-docx`). pandoc is an acceptable fallback but does not enforce Arial 12 without a reference doc, so prefer the script.

## Additional resources

- Block schema and options: [SPEC.md](SPEC.md)
- A worked end-to-end example: [examples.md](examples.md)
