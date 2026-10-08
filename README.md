# XD DOC FORGE — MCQ + Normal Document Converter

A local Streamlit application for creating polished MCQ Question/Answer Sets and for formatting ordinary documents.

## New in v2.3

### Normal Document mode
Choose **Normal Document** instead of MCQ mode when the input is an ordinary file.

Supported input:
- PDF
- Word (.docx)
- TXT

Supported output:
- PDF
- Word (.docx)

The Normal Document mode uses the same default branding as the MCQ workflow:
- First-page heading/contact block:
  - SRUTI ACADEMY
  - BARASAT ASAWANIPALLY SCHOOL ROAD
  - KOLKATA-700124
  - MOB NO.6290000542
- Watermark:
  - SRUTI ACADEMY
  - 6290000542
- Portrait is the default orientation.
- Page numbers appear at the bottom as `1/6`, `2/6`, `3/6`, etc.
- The heading is placed only at the beginning of the first page.

### PDF fidelity / equations
For a source PDF, PDF output keeps the original PDF page rendering and overlays the branding. This is the safest path for equations, mathematical symbols, diagrams and unusual fonts because the source PDF is not re-parsed.

For PDF → Word, the final branded PDF pages are placed into Word as images. This preserves visual appearance and equations exactly, but the page contents are not individually editable.

For Word → Word, the original Word document is retained as far as python-docx allows and the branding is added.

### MCQ mode
The previous MCQ functionality remains available, including table parsing, numbering controls, topic headings, watermarks, first-page heading text, PDF/Word output and answer sets.

## Run

Run `XD DOC FORGE launcher` after setup, or use the existing setup/launcher files.
