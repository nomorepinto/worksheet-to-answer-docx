# Worked Example: Numerical Errors Lab

Input file `FA1.1_Questions.txt` (a lab sheet, partly mangled from PDF copy).

## 1-2. Parse and classify

- **Objective**: Part B table (1/7 x 7, error from 1), Part C truncation/rounding of 12.987654, Part D T/R labels, definitions in Reflection 1-2/4.
- **Personal**: coin diameter (Part A measurement), which real-world example (Reflection 5), and whether the reflection/conclusion should be filled or user-written.
- **Ambiguity flagged to user**: `17\frac{1}{7}` read as **1/7** (confirmed by the "multiply by 7, error from 1" logic); `12.98765412.987654` read as **12.987654**.

## 3. Ask (one AskUserQuestion call)

- Coin diameter? -> user chose `d = 2.40 cm`.
- Real-world example context? -> user chose `Scientific/finance`.
- Handle reflection/conclusion? -> user chose `Fill with model answers`.

## 4. Verify computations (script output)

```
B 0.14    x7 = 0.98    err=0.02
B 0.143   x7 = 1.001   err=0.001
B 0.1429  x7 = 1.0003  err=0.0003
C 1dp T=12.9   err=0.087654  R=13.0  err=0.012346
C 2dp T=12.98  err=0.007654  R=12.99 err=0.002346
C 3dp T=12.987 err=0.000654  R=12.988 err=0.000346
A d=2.40: pi3=7.200000  pi3.14=7.536000  pi3.14159=7.539816  diff=0.339816
```

## 5. Spec (excerpt `spec.json`)

```json
{
  "output": "FA1.1_Answers.docx",
  "title": "FA1.1 - Truncation and Round-Off Error Investigation",
  "subtitle": "Answer Sheet",
  "blocks": [
    {"type": "heading", "text": "PART A. TRUNCATION ERROR INVESTIGATION"},
    {"type": "para", "text": "Measured Diameter (d): 2.40 cm"},
    {"type": "table", "rows": [
      ["Approximation of pi", "Computed Circumference (cm)"],
      ["3.00000", "7.200000"],
      ["3.14000", "7.536000"],
      ["3.14159", "7.539816"]
    ]},
    {"type": "para", "text": "5. Truncation Error - pi's infinite decimal expansion is cut off partway."}
  ]
}
```

## 6. Generate

```bash
python scripts/build_docx.py spec.json
# saved FA1.1_Answers.docx
```

## 7. Verify + clean up

Open the .docx, confirm Normal font is Arial 12, table count, and section headings; then delete `spec.json`. If the save hit `PermissionError` because the file was open in Word, write `FA1.1_Answers_tmp.docx` instead, ask the user to close Word, re-save the real name, and delete the temp.
