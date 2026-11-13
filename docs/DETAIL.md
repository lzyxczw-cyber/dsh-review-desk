![Review desk: three columns, with a table rendered in the edit view](preview.svg)

## What it solves

One manuscript has to pass source checks, logic review, wording, and consistency, and the traces usually end up in comments, chat, and a second document. Review desk keeps "read — ask — decide — apply" in one three-column workspace: the manuscript on the left, the review actions in the middle, and every result on the right.

## How it works

1. Import or paste a manuscript (`.docx` keeps headings, lists, bold, and tables; `.pdf` rebuilds paragraphs).
2. Select a passage and pick one of four analyses, or type an instruction such as "cut to 60 characters".
3. Results land on the right: the newest opens, the rest fold. Rewrites replace the passage with undo; keep asking for changes on the card itself.
4. Click a marked passage to jump to and highlight its results — every operation on that passage at once.

## The edit view

Edit renders the document as an editable canvas: tables, images, headings, and lists look final, and you click anywhere and type in place — no popup, no layout jump; it writes back on blur and table cells edit the same way. Switch to Source for a classic whole-document editor, or Preview for read-only.

## Data

The manuscript, results, column widths, and model choice stay in this browser. Export → Export review package writes one JSON file with the manuscript, every result, and the current instruction; Import → Import review package continues on another computer. The original .docx/.pdf is not included.
