# JSON Spec Reference

The spec is a JSON object. `blocks` is an ordered list rendered top to bottom.

## Top-level keys

| Key | Type | Required | Meaning |
|-----|------|----------|---------|
| `output` | string | no | Output filename (default `Answers.docx`; override with a 2nd CLI arg). |
| `title` | string | no | Centered, bold line at the top. |
| `subtitle` | string | no | Centered line under the title. |
| `blocks` | array | yes | The content blocks, in order. |

## Block types

| `type` | Fields | Renders |
|--------|--------|---------|
| `heading` | `text` | Bold paragraph (section header, e.g. `PART A. ...`). |
| `para` | `text`, optional `bold` (false), `align` (`left`/`center`/`right`), `size` (12) | Normal paragraph. Use for questions and answers. Indent answers with leading spaces in `text`. |
| `table` | `rows` (array of arrays), optional `header` (true) | Bordered table; first row bold when `header` is true. All cells are Arial 12. |
| `spacer` | — | One blank line. |

## Formatting guarantees

The script enforces **Arial 12** on the Normal style, every paragraph run, and every table cell. Tables use the `Table Grid` border style only — no shading, accent colors, or themes. There are no style knobs to set; that rigidity is the point (the user asked for "plain, no designs").

## Rounding vs. truncation values

Generate numbers first (see SKILL.md step 4), then paste verified results as strings into `rows`. Keep display precision consistent with the worksheet's convention.
