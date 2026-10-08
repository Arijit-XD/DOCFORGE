
import io
import re
from pathlib import Path

import streamlit as st
import pandas as pd
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None
from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, A3, A5, LETTER, LEGAL, landscape, portrait
from docx.shared import Inches, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)


# -----------------------------
# Parsing
# -----------------------------

OPTION_RE = re.compile(r"^\s*([A-Da-d])\s*[\)\.\:\-]\s*(.+?)\s*$")
ANSWER_RE = re.compile(
    r"^\s*(?:answer|correct\s*answer|ans)\s*[:\-]?\s*(.*)\s*$",
    re.IGNORECASE,
)
QUESTION_RE = re.compile(r"^\s*(\d+)\s*[\.\)\-:]\s*(.+?)\s*$")


def clean_text(value):
    if value is None:
        return ""
    value = str(value).replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value).strip()
    return value


def extract_answer_letter(value):
    """Return A/B/C/D from values such as C, C), C. or C) Central Asia."""
    value = clean_text(value)
    if not value:
        return None

    m = re.match(r"^\s*([A-Da-d])(?:\s*[\)\.\:\-]|\s*$)", value)
    if m:
        return m.group(1).upper()

    # Sometimes the answer cell contains the exact option text.
    return None


def option_text_by_letter(options, letter):
    for opt_letter, opt_text in options:
        if opt_letter == letter:
            return opt_text
    return None


def validate_questions(questions):
    errors = []
    seen_numbers = set()

    for q in questions:
        n = q["number"]

        if n in seen_numbers:
            errors.append(f"Question {n}: duplicate question number.")
        seen_numbers.add(n)

        if not q["question"]:
            errors.append(f"Question {n}: question text is empty.")

        letters = [x[0] for x in q["options"]]
        if letters != ["A", "B", "C", "D"]:
            errors.append(
                f"Question {n}: expected exactly A, B, C, D options; found "
                f"{', '.join(letters) if letters else 'none'}."
            )

        if any(not text for _, text in q["options"]):
            errors.append(f"Question {n}: one or more option texts are empty.")

        if not q["answer"]:
            errors.append(f"Question {n}: answer is missing or invalid.")
        elif q["answer"] not in letters:
            errors.append(
                f"Question {n}: answer '{q['answer']}' does not match A/B/C/D."
            )

        if q.get("answer_text"):
            expected = option_text_by_letter(q["options"], q["answer"])
            if expected and q["answer_text"].lower() != expected.lower():
                # Don't fail on wording differences; report as a warning.
                q.setdefault("warnings", []).append(
                    f"Answer text differs from option {q['answer']}."
                )

    nums = [q["number"] for q in questions]
    if nums and nums != sorted(nums):
        errors.append("Question numbers are not in ascending order.")

    return errors


def escape_xml(value):
    """Escape text for ReportLab Paragraph XML."""
    import html
    return html.escape(clean_text(value), quote=False)


def split_inline_question(line):
    """Parse a complete MCQ when question/options/answer are on one line."""
    line = clean_text(line)
    qmatch = QUESTION_RE.match(line)
    if not qmatch:
        return None
    number = int(qmatch.group(1))
    body = qmatch.group(2).strip()

    # Answer can appear at the end, e.g. Answer: C / Ans C / Correct Answer: B.
    answer = None
    answer_text = None
    am = re.search(r"\s+(?:answer|correct\s*answer|ans)\s*[:\-]?\s*([A-Da-d])(?:\s*\.|\s*$)", body, re.I)
    if am:
        answer = am.group(1).upper()
        body = body[:am.start()].strip()

    # Split A/B/C/D option markers even when there are no line breaks.
    matches = list(re.finditer(r"(?<!\w)([A-Da-d])\s*[\)\.\:\-]\s*", body))
    if len(matches) >= 4:
        question = clean_text(body[:matches[0].start()])
        options = []
        for i, m in enumerate(matches[:4]):
            start = m.end()
            end = matches[i + 1].start() if i + 1 < len(matches) else len(body)
            txt = clean_text(body[start:end])
            options.append((m.group(1).upper(), txt))
        return {
            "number": number, "question": question, "options": options,
            "answer": answer, "answer_text": answer_text, "warnings": []
        }
    return None


def parse_plain_text(text):
    raw_lines = [clean_text(x) for x in text.splitlines()]
    raw_lines = [x for x in raw_lines if x]

    # First support one-MCQ-per-line input.
    inline_questions = []
    inline_ok = True
    for line in raw_lines:
        parsed = split_inline_question(line)
        if parsed:
            inline_questions.append(parsed)
        else:
            inline_ok = False
            break
    if inline_questions and inline_ok:
        return inline_questions

    lines = raw_lines
    questions = []
    current = None
    pending_answer = None

    def finish():
        nonlocal current, pending_answer
        if current is not None:
            if pending_answer:
                current["answer"] = pending_answer
            questions.append(current)
        current = None
        pending_answer = None

    for line in lines:
        qmatch = QUESTION_RE.match(line)
        omatch = OPTION_RE.match(line)
        amatch = ANSWER_RE.match(line)

        if qmatch:
            finish()
            # If the numbered line itself contains A) ... B) ... C) ... D), parse it.
            inline = split_inline_question(line)
            if inline:
                current = inline
                continue
            current = {
                "number": int(qmatch.group(1)),
                "question": qmatch.group(2),
                "options": [],
                "answer": None,
                "answer_text": None,
                "warnings": [],
            }
            continue

        if current is None:
            continue

        if omatch:
            letter = omatch.group(1).upper()
            text_value = omatch.group(2).strip()
            if any(x[0] == letter for x in current["options"]):
                current["warnings"].append(f"Duplicate option {letter}.")
            else:
                current["options"].append((letter, text_value))
            continue

        if amatch:
            raw = amatch.group(1).strip()
            letter = extract_answer_letter(raw)
            if letter:
                current["answer"] = letter
                current["answer_text"] = re.sub(r"^\s*[A-Da-d]\s*[\)\.\:\-]?\s*", "", raw).strip()
            else:
                current["answer_text"] = raw
                for ltr, opt in current["options"]:
                    if opt.lower() == raw.lower():
                        current["answer"] = ltr
                        break
            continue

        if current["options"]:
            last_letter, last_text = current["options"][-1]
            current["options"][-1] = (last_letter, f"{last_text} {line}".strip())
        else:
            current["question"] = f"{current['question']} {line}".strip()

    finish()
    return questions


def parse_docx(uploaded_file):
    """Robust DOCX MCQ parser.

    Prefer the actual Word table structure when present. Real-world Word files
    often contain blank cells, wrapped/multi-paragraph cells, reordered columns,
    slightly different header names, or options written on one line. This parser
    deliberately tolerates those variations instead of requiring one exact
    template.
    """
    doc = Document(uploaded_file)
    questions = []

    def norm_header(value):
        h = clean_text(value).lower()
        h = re.sub(r"[^a-z0-9]+", " ", h).strip()
        return h

    def header_kind(h):
        h = norm_header(h)
        if not h:
            return None
        if h in {"question", "questions", "question text", "question no question", "sentence", "statement", "idiom", "idiom phrase", "phrase"} or "question" in h or h.startswith("sentence") or "statement" in h or "idiom" in h or h == "phrase":
            return "question"
        if h in {"option", "options", "choice", "choices", "meaning"} or "option" in h or "choice" in h:
            return "options"
        if h in {"answer", "ans", "correct", "correct answer", "correct ans", "key", "answer key"}:
            return "answer"
        if h in {"sl no", "sl number", "serial no", "serial number", "sr no", "sr number",
                 "q no", "q no", "question no", "question number", "number", "no"}:
            return "number"
        return None

    for table in doc.tables:
        rows = table.rows
        if not rows:
            continue

        # Find the header row within the first few rows. Some Word files have
        # a title/blank row above the actual column headings.
        header_row_idx = None
        header = None
        for ri, row in enumerate(rows[:6]):
            candidate = [norm_header(c.text) for c in row.cells]
            kinds = [header_kind(h) for h in candidate]
            if "question" in kinds and ("options" in kinds or "answer" in kinds or "number" in kinds):
                header_row_idx = ri
                header = candidate
                break

        if header_row_idx is None:
            # If there is no recognizable header, look for a table that still
            # contains numbered question rows and infer the common 4-column
            # layout from its first non-empty row.
            for ri, row in enumerate(rows[:4]):
                cells = [clean_text(c.text) for c in row.cells]
                if len(cells) >= 3 and any(re.match(r"^\s*\d+\s*$", c) for c in cells):
                    header_row_idx = ri - 1 if ri > 0 else None
                    break
            if header_row_idx is None:
                continue
            header = [norm_header(c.text) for c in rows[header_row_idx].cells] if header_row_idx is not None else []

        kinds = [header_kind(h) for h in header]
        q_col = next((i for i, k in enumerate(kinds) if k == "question"), None)
        o_col = next((i for i, k in enumerate(kinds) if k == "options"), None)
        a_col = next((i for i, k in enumerate(kinds) if k == "answer"), None)
        n_col = next((i for i, k in enumerate(kinds) if k == "number"), None)

        # Some common real-world tables use domain-specific headings such as
        # ``Sentence`` or ``Idiom/Phrase`` and ``Meaning``. If the header
        # semantics are still incomplete but the table is the conventional
        # four-column No | Question | Options | Answer shape, infer the
        # positions rather than rejecting the entire table.
        if len(header) >= 4 and q_col is None and a_col is not None:
            q_col = 1
        if len(header) >= 4 and o_col is None and q_col == 1 and a_col is not None:
            o_col = 2
        if len(header) >= 4 and n_col is None and q_col == 1 and o_col == 2 and a_col == 3:
            n_col = 0

        # A question column is essential. Answer is optional because a question
        # paper may intentionally contain no answer column.
        if q_col is None:
            continue

        table_questions = []
        next_number = 1

        for row in rows[header_row_idx + 1:]:
            cells = [clean_text(c.text) for c in row.cells]
            if not any(cells):
                continue

            # Skip repeated header rows appearing mid-table.
            if any(header_kind(c) == "question" for c in cells) and any(
                header_kind(c) in {"options", "answer"} for c in cells
            ):
                continue

            qtext = cells[q_col] if q_col < len(cells) else ""
            raw_answer = cells[a_col] if a_col is not None and a_col < len(cells) else ""

            # Some files have an accidental blank question cell. Recover the
            # question from another cell containing a question mark or a
            # numbered question prefix.
            if not qtext:
                for i, c in enumerate(cells):
                    if i in {o_col, a_col, n_col} or not c:
                        continue
                    if "?" in c or QUESTION_RE.match(c):
                        qtext = c
                        break

            if not qtext:
                continue

            # Question number: use explicit number, then a prefix in question
            # text, then a stable sequential fallback.
            number = None
            if n_col is not None and n_col < len(cells):
                m = re.search(r"\d+", cells[n_col])
                if m:
                    number = int(m.group())

            qm = QUESTION_RE.match(qtext)
            if qm:
                if number is None:
                    number = int(qm.group(1))
                qtext = qm.group(2).strip()

            if number is None:
                number = next_number

            next_number = max(next_number, number + 1)

            # Options can be a single blob, multiple paragraphs, or dedicated
            # A/B/C/D columns.
            options = []
            if o_col is not None and o_col < len(cells):
                options = parse_options_from_blob(cells[o_col])

            if len(options) < 4:
                col_options = []
                for i, h in enumerate(header):
                    if h in {"a", "b", "c", "d"} and i < len(cells):
                        value = clean_text(cells[i])
                        if value:
                            col_options.append((h.upper(), value))
                if len(col_options) >= len(options):
                    options = col_options

            if len(options) < 4:
                # Last fallback: inspect every non-question/non-number cell.
                # This handles merged/reordered columns and malformed headers.
                blob = " ".join(
                    cells[i] for i in range(len(cells))
                    if i not in {q_col, a_col, n_col}
                )
                fallback_options = parse_options_from_blob(blob)
                if len(fallback_options) > len(options):
                    options = fallback_options

            answer = extract_answer_letter(raw_answer)
            answer_text = clean_text(raw_answer) or None

            if not answer and raw_answer and options:
                raw_norm = clean_text(raw_answer).lower()
                for ltr, opt in options:
                    if raw_norm == clean_text(opt).lower():
                        answer = ltr
                        break
                # Also allow an answer cell containing only an option letter
                # followed by extra punctuation/text.
                if not answer:
                    m = re.match(r"^\s*([A-Da-d])(?:\s*[\)\.\:\-])?", raw_answer)
                    if m:
                        answer = m.group(1).upper()

            # Keep only genuine MCQ rows. A row with a question is useful even
            # when options/answer are imperfect; validation can flag it later.
            table_questions.append({
                "number": number,
                "question": clean_text(qtext),
                "options": options,
                "answer": answer,
                "answer_text": answer_text,
                "warnings": [],
            })

        if table_questions:
            questions.extend(table_questions)

    if questions:
        return questions

    # Non-standard tables: flatten their cells and use the flexible text parser.
    if doc.tables:
        table_lines = []
        for table in doc.tables:
            for row in table.rows:
                cells = [clean_text(c.text) for c in row.cells]
                if cells:
                    table_lines.append(" | ".join(cells))
        table_questions = parse_plain_text("\n".join(table_lines))
        if table_questions:
            return table_questions

    # Finally support DOCX files containing ordinary paragraphs rather than
    # tables.
    text = "\n".join(p.text for p in doc.paragraphs)
    return parse_plain_text(text)


def parse_pdf(uploaded_file):
    """Extract selectable PDF text and feed it through the same flexible MCQ parser.

    This preserves Unicode mathematical symbols when the source PDF actually
    contains text glyphs. Image-only/scanned PDFs and complex vector equations
    need OCR/equation-image handling and are reported clearly instead of being
    silently mis-converted.
    """
    if fitz is None:
        raise RuntimeError("PDF support requires PyMuPDF. Run First Time Setup again to install the updated requirements.")
    data = uploaded_file.getvalue()
    doc = fitz.open(stream=data, filetype="pdf")
    try:
        page_texts = [page.get_text("text") for page in doc]
    finally:
        doc.close()
    text = "\n".join(page_texts)
    if not clean_text(text):
        raise ValueError("This PDF appears to contain no selectable text. It may be scanned/image-based. Please use a text PDF or OCR it first.")
    return parse_plain_text(text), text


def detect_docx_headings(uploaded_file):
    """Find likely topic headings in a Word document."""
    try:
        doc=Document(uploaded_file)
    except Exception:
        return []
    found=[]
    for p in doc.paragraphs:
        text=clean_text(p.text)
        if not text or QUESTION_RE.match(text) or OPTION_RE.match(text) or ANSWER_RE.match(text):
            continue
        style=(getattr(p.style,'name','') or '').lower()
        bold_runs=[r for r in p.runs if clean_text(r.text)]
        all_bold=bool(bold_runs) and all(bool(r.bold) for r in bold_runs)
        range_match=re.search(r'(\d+)\s*(?:-|–|—|to)\s*(\d+)',text,re.I)
        likely=style.startswith('heading') or all_bold or bool(range_match)
        if not likely and len(text)<=70 and len(text.split())<=8 and not text.endswith(('?',':')):
            likely=True
        if not likely: continue
        title=text
        start=end=None
        if range_match:
            start=int(range_match.group(1)); end=int(range_match.group(2))
            title=(text[:range_match.start()]+' '+text[range_match.end():]).strip(' -–—:|')
        if title and title.lower() not in {x['title'].lower() for x in found}:
            found.append({'title':title,'start':start,'end':end})
    return found


def heading_text_from_candidates(candidates):
    lines=[]
    for h in candidates:
        if h.get('start') is not None and h.get('end') is not None:
            lines.append(f"{h['start']}-{h['end']} | {h['title']}")
        else:
            lines.append(f" | {h['title']}")
    return "\n".join(lines)


def parse_heading_rules_text(text):
    rules=[]; bad=[]
    for line in str(text or '').splitlines():
        line=line.strip()
        if not line: continue
        m=re.match(r'^\s*(\d+)\s*(?:-|–|—|to)\s*(\d+)\s*\|\s*(.+?)\s*$',line,re.I)
        if not m:
            bad.append(line); continue
        rules.append({'start':int(m.group(1)),'end':int(m.group(2)),'title':m.group(3).strip()})
    return normalize_heading_rules(rules),bad


def parse_options_from_blob(blob):
    blob = clean_text(blob)
    if not blob:
        return []

    # Normalize common separators while preserving option text.
    # A plain ``B.`` also occurs inside names such as ``B.R. Ambedkar``.
    # Treat a dot as an option marker only when it is followed by whitespace;
    # ) / : / - remain valid even when the source omits that whitespace.
    option_marker = re.compile(
        r"(?<!\w)([A-Da-d])\s*(?:[\)\:\-]\s*|\.\s+(?=\S))"
    )
    matches = list(option_marker.finditer(blob))
    if not matches:
        return []

    result = []
    for i, match in enumerate(matches):
        letter = match.group(1).upper()
        start = match.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(blob)
        text_value = clean_text(blob[start:end])
        if text_value:
            result.append((letter, text_value))
    return result


# -----------------------------
# PDF generation
# -----------------------------

def make_styles():
    styles = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "CustomTitle",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=16,
            leading=20,
            alignment=TA_CENTER,
            textColor=colors.black,
            spaceAfter=10,
        ),
        "question": ParagraphStyle(
            "Question",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=13,
        ),
        "option": ParagraphStyle(
            "Option",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9,
            leading=11,
        ),
        "small": ParagraphStyle(
            "Small",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
        ),
        "heading": ParagraphStyle(
            "TopicHeading",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=12,
            leading=15,
            alignment=TA_LEFT,
            spaceBefore=4,
            spaceAfter=4,
        ),
        "first_page_header": ParagraphStyle(
            "FirstPageHeader",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=12,
            leading=15,
            alignment=TA_CENTER,
            spaceAfter=1.5,
        ),
    }


def _resolve_page_size(size_name, orientation_name):
    sizes = {
        "A3": A3, "A4": A4, "A5": A5,
        "Letter": LETTER, "Legal": LEGAL,
    }
    base = sizes.get(size_name, A4)
    return landscape(base) if orientation_name == "Landscape" else portrait(base)


class NumberedCanvas(pdfcanvas.Canvas):
    """Canvas that writes current-page/total-pages at the bottom after pagination is known."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._page_states = []

    def showPage(self):
        self._page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total = len(self._page_states)
        for state in self._page_states:
            self.__dict__.update(state)
            self._draw_page_counter(total)
            super().showPage()
        super().save()

    def _draw_page_counter(self, total):
        self.saveState()
        self.setFillColor(colors.HexColor("#718096"))
        self.setFont("Helvetica", 7.5)
        self.drawRightString(self._pagesize[0] - 8 * mm, 5.5 * mm, f"{self._pageNumber}/{total}")
        self.restoreState()


def _watermark_callback(text, pagesize):
    # Preserve line breaks so users can enter a multi-line watermark.
    lines = [line.strip() for line in str(text or "").splitlines() if line.strip()]

    def callback(canvas, doc):
        if not lines:
            return
        canvas.saveState()
        try:
            canvas.setFillAlpha(0.008)
        except Exception:
            pass
        canvas.setFillColor(colors.HexColor("#C8D0DB"))
        canvas.translate(pagesize[0] / 2, pagesize[1] / 2)
        canvas.rotate(35)
        canvas.setFont("Helvetica-Bold", 24)
        line_gap = 34
        start_y = ((len(lines) - 1) * line_gap) / 2
        for idx, line in enumerate(lines[:5]):
            canvas.drawCentredString(0, start_y - idx * line_gap, line[:90])
        canvas.restoreState()
    return callback


def apply_numbering(questions, mode, start_value=1):
    """Create output numbering while preserving the original/source number."""
    out=[]
    for i,q in enumerate(questions):
        nq=dict(q)
        nq["source_number"] = q.get("source_number", q["number"])
        if mode == "Use source numbering":
            nq["number"] = q["number"]
        elif mode == "Start from 1":
            nq["number"] = i+1
        else:
            nq["number"] = int(start_value)+i
        out.append(nq)
    return out


def detect_heading_lines(text):
    """Detect explicit heading/range lines in pasted TXT content."""
    candidates=[]
    pending=[]
    for raw in str(text or '').splitlines():
        line=clean_text(raw)
        if not line: continue
        qm=QUESTION_RE.match(line)
        if qm:
            n=int(qm.group(1))
            for h in pending:
                candidates.append({"title":h, "start":n, "end":n, "confidence":"Possible"})
            pending=[]
            continue
        range_match=re.match(r'^\s*(\d+)\s*(?:-|–|—|to)\s*(\d+)\s*(?:\||:|-)\s*(.+?)\s*$',line,re.I)
        if range_match:
            candidates.append({"title":clean_text(range_match.group(3)),"start":int(range_match.group(1)),"end":int(range_match.group(2)),"confidence":"Explicit range"})
            continue
        if OPTION_RE.match(line) or ANSWER_RE.match(line):
            continue
        if re.match(r'^[-_=]{3,}$', line):
            continue
        if len(line)<=80 and len(line.split())<=10 and not line.endswith(('?', ':')):
            pending.append(line)
    return candidates


def parse_pdf(uploaded_file):
    """Extract selectable PDF text and feed it through the same flexible MCQ parser.

    This preserves Unicode mathematical symbols when the source PDF actually
    contains text glyphs. Image-only/scanned PDFs and complex vector equations
    need OCR/equation-image handling and are reported clearly instead of being
    silently mis-converted.
    """
    if fitz is None:
        raise RuntimeError("PDF support requires PyMuPDF. Run First Time Setup again to install the updated requirements.")
    data = uploaded_file.getvalue()
    doc = fitz.open(stream=data, filetype="pdf")
    try:
        page_texts = [page.get_text("text") for page in doc]
    finally:
        doc.close()
    text = "\n".join(page_texts)
    if not clean_text(text):
        raise ValueError("This PDF appears to contain no selectable text. It may be scanned/image-based. Please use a text PDF or OCR it first.")
    return parse_plain_text(text), text


def detect_docx_headings(uploaded_file):
    """Find explicit/strongly formatted topic headings in Word, including table rows."""
    try:
        uploaded_file.seek(0)
        doc=Document(uploaded_file)
        uploaded_file.seek(0)
    except Exception:
        return []
    found=[]
    def add(title,start=None,end=None,confidence="Possible"):
        title=clean_text(title)
        if not title: return
        key=title.lower()
        if key not in {x['title'].lower() for x in found}:
            found.append({'title':title,'start':start,'end':end,'confidence':confidence})

    # Paragraph headings
    for p in doc.paragraphs:
        text=clean_text(p.text)
        if not text or QUESTION_RE.match(text) or OPTION_RE.match(text) or ANSWER_RE.match(text):
            continue
        style=(getattr(p.style,'name','') or '').lower()
        runs=[r for r in p.runs if clean_text(r.text)]
        all_bold=bool(runs) and all(bool(r.bold) for r in runs)
        rm=re.search(r'(\d+)\s*(?:-|–|—|to)\s*(\d+)',text,re.I)
        likely=style.startswith('heading') or all_bold or bool(rm)
        if not likely and len(text)<=70 and len(text.split())<=8 and not text.endswith(('?',':')):
            likely=True
        if likely:
            if rm:
                title=(text[:rm.start()]+' '+text[rm.end():]).strip(' -–—:|')
                add(title,int(rm.group(1)),int(rm.group(2)),"Explicit range")
            else:
                add(text,None,None,"Word heading")

    # A table row with a single non-empty cell is a common topic-heading layout.
    for table in doc.tables:
        for row in table.rows:
            vals=[clean_text(c.text) for c in row.cells]
            nonempty=[v for v in vals if v]
            if len(nonempty)==1:
                text=nonempty[0]
                if QUESTION_RE.match(text) or OPTION_RE.match(text) or ANSWER_RE.match(text):
                    continue
                rm=re.search(r'(\d+)\s*(?:-|–|—|to)\s*(\d+)',text,re.I)
                if rm:
                    title=(text[:rm.start()]+' '+text[rm.end():]).strip(' -–—:|')
                    add(title,int(rm.group(1)),int(rm.group(2)),"Table range")
                elif len(text)<=80:
                    add(text,None,None,"Table heading")
    return found


def heading_text_from_candidates(candidates):
    lines=[]
    for h in candidates:
        if h.get('start') is not None and h.get('end') is not None:
            lines.append(f"{h['start']}\t{h['end']}\t{h['title']}")
        else:
            lines.append(f"\t\t{h['title']}")
    return "\n".join(lines)


def parse_heading_rules_text(text):
    """Backward-compatible parser for tab or pipe separated topic rows."""
    rules=[]; bad=[]
    for line in str(text or '').splitlines():
        line=line.strip()
        if not line: continue
        m=re.match(r'^\s*(\d+)\s*(?:-|–|—|to)\s*(\d+)\s*(?:\||\t)\s*(.+?)\s*$',line,re.I)
        if not m:
            # Also accept: Topic | 200-225
            m2=re.match(r'^\s*(.+?)\s*\|\s*(\d+)\s*(?:-|–|—|to)\s*(\d+)\s*$',line,re.I)
            if m2:
                rules.append({'start':int(m2.group(2)),'end':int(m2.group(3)),'title':m2.group(1).strip()}); continue
            bad.append(line); continue
        rules.append({'start':int(m.group(1)),'end':int(m.group(2)),'title':m.group(3).strip()})
    return normalize_heading_rules(rules),bad


def normalize_heading_rules(rules):
    out=[]
    for rule in rules or []:
        try:
            a=int(rule.get('start')); b=int(rule.get('end')); title=clean_text(rule.get('title',''))
        except Exception:
            continue
        if not title: continue
        if a>b: a,b=b,a
        out.append({"start":a,"end":b,"title":title})
    return sorted(out,key=lambda x:(x['start'],x['end']))


def heading_for_question(q,rules,basis="output"):
    n=q.get('number') if basis=="output" else q.get('source_number',q.get('number'))
    for r in rules:
        if r['start']<=n<=r['end']:
            return r['title']
    return None


def group_by_headings(questions,rules,basis="output"):
    if not rules: return [(None,questions)]
    groups=[]; bucket=[]; current=None
    for q in questions:
        title=heading_for_question(q,rules,basis)
        if bucket and title!=current:
            groups.append((current,bucket)); bucket=[]
        bucket.append(q); current=title
    if bucket: groups.append((current,bucket))
    return groups

def _build_mcq_table_pdf(questions, title, include_correct=False, watermark="", page_size="A4", orientation="Landscape", heading_rules=None, heading_basis="output", first_page_text=""):
    buffer=io.BytesIO(); styles=make_styles(); pagesize=_resolve_page_size(page_size,orientation)
    doc=SimpleDocTemplate(buffer,pagesize=pagesize,rightMargin=8*mm,leftMargin=8*mm,topMargin=10*mm,bottomMargin=12*mm,title=title)
    story=[]
    first_lines=[line.strip() for line in str(first_page_text or "").splitlines() if line.strip()]
    if first_lines:
        for line in first_lines[:8]:
            hp=Paragraph(escape_xml(line), styles["first_page_header"])
            story.append(hp)
        story.append(Spacer(1, 5*mm))
    story += [Paragraph(escape_xml(title),styles["title"]),Spacer(1,2*mm)]
    usable=pagesize[0]-16*mm
    if include_correct:
        col_widths=[usable*.055,usable*.335,usable*.105,usable*.105,usable*.105,usable*.105,usable*.190]
    else:
        col_widths=[usable*.055,usable*.385,usable*.140,usable*.140,usable*.140,usable*.140]
    groups=group_by_headings(questions,normalize_heading_rules(heading_rules),heading_basis)
    for gi,(heading,group) in enumerate(groups):
        if heading:
            story += [Paragraph(escape_xml(heading),styles["heading"]),Spacer(1,1.5*mm)]
        headers=["Q.No.","Question","A","B","C","D"]+(["Correct Answer"] if include_correct else [])
        data=[headers]
        for q in group:
            om={l:t for l,t in q["options"]}
            row=[Paragraph(escape_xml(str(q["number"])),styles["small"]),Paragraph(escape_xml(q["question"]),styles["small"])]
            row += [Paragraph(escape_xml(om.get(l,"")),styles["small"]) for l in ["A","B","C","D"]]
            if include_correct:
                row.append(Paragraph(f"<b>{escape_xml(q['answer'] or '')} — {escape_xml(om.get(q['answer'],''))}</b>",styles["small"]))
            data.append(row)
        table=Table(data,colWidths=col_widths,repeatRows=1)
        table.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#E9EEF5")),("TEXTCOLOR",(0,0),(-1,0),colors.black),("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("ALIGN",(0,0),(0,-1),"CENTER"),("ALIGN",(2,0),(-1,-1),"CENTER"),("VALIGN",(0,0),(-1,-1),"MIDDLE"),("GRID",(0,0),(-1,-1),0.5,colors.grey),("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)]))
        story.append(table)
        if gi<len(groups)-1: story.append(Spacer(1,4*mm))
    callback=_watermark_callback(watermark,pagesize)
    doc.build(story,onFirstPage=callback,onLaterPages=callback,canvasmaker=NumberedCanvas)
    return buffer.getvalue()

def build_question_pdf(questions, title, watermark="", page_size="A4", orientation="Landscape", heading_rules=None, heading_basis="output", first_page_text=""):
    return _build_mcq_table_pdf(questions,title,False,watermark,page_size,orientation,heading_rules,heading_basis,first_page_text)


def build_answer_pdf(questions, title, watermark="", page_size="A4", orientation="Landscape", heading_rules=None, heading_basis="output", first_page_text=""):
    return _build_mcq_table_pdf(questions,title,True,watermark,page_size,orientation,heading_rules,heading_basis,first_page_text)


def _set_word_watermark(section, text):
    """Insert a faint, centred, diagonal Word watermark behind document text."""
    if not text:
        return
    header = section.header
    for p in list(header._element.p_lst):
        p.getparent().remove(p)
    lines = [clean_text(x) for x in str(text).splitlines() if clean_text(x)][:5]
    if not lines:
        return
    total = len(lines)
    for i, line in enumerate(lines):
        offset = (i - (total - 1) / 2) * 34
        safe = escape_xml(line)[:120]
        xml = f"""<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\">
          <w:r><w:pict>
            <v:shape id=\"MCQForgeWatermark{i}\" type=\"#_x0000_t136\"
              style=\"position:absolute;margin-left:0;margin-top:{offset:.1f}pt;width:520pt;height:70pt;mso-position-horizontal:center;mso-position-vertical:center;rotation:-35;z-index:-251658752\"
              fillcolor=\"#D5DCE6\" stroked=\"f\" o:allowincell=\"f\">
              <v:fill opacity=\"0.08\"/>
              <v:textpath on=\"t\" style=\"font-family:Arial;font-size:24pt;font-weight:bold\" string=\"{safe}\"/>
            </v:shape>
          </w:pict></w:r>
        </w:p>"""
        header._element.append(parse_xml(xml))

def _add_word_page_counter(section):
    footer=section.footer
    p=footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    p.text=""
    r=p.add_run()
    fld_begin=OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
    instr=OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text=" PAGE "
    fld_sep=OxmlElement("w:fldChar"); fld_sep.set(qn("w:fldCharType"), "separate")
    t=OxmlElement("w:t"); t.text="1"
    fld_end=OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
    r._r.extend([fld_begin,instr,fld_sep,t,fld_end])
    r2=p.add_run("/")
    r3=p.add_run()
    fld_begin2=OxmlElement("w:fldChar"); fld_begin2.set(qn("w:fldCharType"), "begin")
    instr2=OxmlElement("w:instrText"); instr2.set(qn("xml:space"), "preserve"); instr2.text=" NUMPAGES "
    fld_sep2=OxmlElement("w:fldChar"); fld_sep2.set(qn("w:fldCharType"), "separate")
    t2=OxmlElement("w:t"); t2.text="1"
    fld_end2=OxmlElement("w:fldChar"); fld_end2.set(qn("w:fldCharType"), "end")
    r3._r.extend([fld_begin2,instr2,fld_sep2,t2,fld_end2])
    for rr in p.runs:
        rr.font.name="Arial"; rr.font.size=Pt(8); rr.font.color.rgb=__import__('docx').shared.RGBColor(113,128,150)


def build_word_doc(questions,title,include_correct=False,watermark="",page_size="A4",orientation="Landscape",heading_rules=None,heading_basis="output", first_page_text=""):
    doc=Document(); section=doc.sections[0]
    page_map={"A3":(11.69,16.54),"A4":(8.27,11.69),"A5":(5.83,8.27),"Letter":(8.5,11),"Legal":(8.5,14)}
    w,h=page_map.get(page_size,page_map["A4"])
    if orientation=="Landscape": w,h=h,w; section.orientation=WD_ORIENT.LANDSCAPE
    section.page_width=Inches(w); section.page_height=Inches(h)
    section.top_margin=Inches(.45); section.bottom_margin=Inches(.55); section.left_margin=Inches(.45); section.right_margin=Inches(.45)
    _set_word_watermark(section,watermark)
    _add_word_page_counter(section)
    first_lines=[line.strip() for line in str(first_page_text or "").splitlines() if line.strip()]
    for line in first_lines[:8]:
        hp=doc.add_paragraph(); hp.alignment=WD_ALIGN_PARAGRAPH.CENTER; hp.paragraph_format.space_after=Pt(1)
        hr=hp.add_run(line); hr.bold=True; hr.font.name="Arial"; hr.font.size=Pt(13)
    if first_lines:
        doc.add_paragraph().paragraph_format.space_after=Pt(2)
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(8); r=p.add_run(title); r.bold=True; r.font.name="Arial"; r.font.size=Pt(18)
    available=w-.90; proportions=[.055,.335,.105,.105,.105,.105,.190] if include_correct else [.055,.385,.140,.140,.140,.140]; widths=[Inches(available*x) for x in proportions]
    groups=group_by_headings(questions,normalize_heading_rules(heading_rules),heading_basis)
    for gi,(heading,group) in enumerate(groups):
        if heading:
            hp=doc.add_paragraph(); hp.paragraph_format.space_before=Pt(8); hp.paragraph_format.space_after=Pt(4); hr=hp.add_run(heading); hr.bold=True; hr.font.name="Arial"; hr.font.size=Pt(14)
        cols=7 if include_correct else 6; table=doc.add_table(rows=1,cols=cols); table.style="Table Grid"; table.autofit=False
        labels=["Q.No.","Question","A","B","C","D"]+(["Correct Answer"] if include_correct else [])
        for i,(cell,label) in enumerate(zip(table.rows[0].cells,labels)):
            cell.text=label; cell.width=widths[i]
            for run in cell.paragraphs[0].runs: run.bold=True
        for q in group:
            row=table.add_row().cells; om={l:t for l,t in q["options"]}; vals=[str(q["number"]),q["question"],om.get("A",""),om.get("B",""),om.get("C",""),om.get("D","")]
            if include_correct: vals.append(f"{q['answer'] or ''} — {om.get(q['answer'],'')}")
            for i,(cell,val) in enumerate(zip(row,vals)): cell.text=val; cell.width=widths[i]
        if gi<len(groups)-1: doc.add_paragraph()
    return doc


DEFAULT_FIRST_PAGE_TEXT = "SRUTI ACADEMY\nBARASAT ASAWANIPALLY SCHOOL ROAD\nKOLKATA-700124\nMOB NO.6290000542"



# -----------------------------
# Normal document conversion
# -----------------------------

def _insert_paragraph_at_start(doc, text, bold=True, size=13):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    r.bold = bold
    r.font.name = "Arial"
    r.font.size = Pt(size)
    body = doc._element.body
    body.remove(p._p)
    body.insert(0, p._p)
    return p


def _add_normal_docx_header(doc, first_page_text, title):
    # Insert the branding at the very beginning while keeping the source
    # document content and formatting intact as much as python-docx allows.
    lines = [x.strip() for x in str(first_page_text or "").splitlines() if x.strip()]
    insert_at = 0
    for line in lines[:8][::-1]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(line)
        r.bold = True; r.font.name = "Arial"; r.font.size = Pt(13)
        body = doc._element.body; body.remove(p._p); body.insert(0, p._p)
    if title.strip():
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(10)
        r = p.add_run(title.strip()); r.bold=True; r.font.name="Arial"; r.font.size=Pt(18)
        body = doc._element.body; body.remove(p._p); body.insert(0, p._p)


def _configure_docx_section(section, page_size="A4", orientation="Portrait"):
    page_map={"A3":(11.69,16.54),"A4":(8.27,11.69),"A5":(5.83,8.27),"Letter":(8.5,11),"Legal":(8.5,14)}
    w,h=page_map.get(page_size,page_map["A4"])
    if orientation == "Landscape":
        w,h=h,w; section.orientation=WD_ORIENT.LANDSCAPE
    else:
        section.orientation=WD_ORIENT.PORTRAIT
    section.page_width=Inches(w); section.page_height=Inches(h)
    section.top_margin=Inches(.55); section.bottom_margin=Inches(.65)
    section.left_margin=Inches(.65); section.right_margin=Inches(.65)


def build_normal_word_from_docx(data, title="", watermark="", page_size="A4", orientation="Portrait", first_page_text=""):
    doc=Document(io.BytesIO(data))
    for section in doc.sections:
        _configure_docx_section(section,page_size,orientation)
        _set_word_watermark(section,watermark)
        _add_word_page_counter(section)
    _add_normal_docx_header(doc,first_page_text,title)
    out=io.BytesIO(); doc.save(out); return out.getvalue()


def build_normal_word_from_text(text, title="", watermark="", page_size="A4", orientation="Portrait", first_page_text=""):
    doc=Document(); section=doc.sections[0]
    _configure_docx_section(section,page_size,orientation)
    _set_word_watermark(section,watermark); _add_word_page_counter(section)
    _add_normal_docx_header(doc,first_page_text,title)
    for raw in str(text or "").splitlines():
        line=raw.strip()
        if not line:
            doc.add_paragraph()
            continue
        p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(5)
        r=p.add_run(line); r.font.name="Arial"; r.font.size=Pt(11)
    out=io.BytesIO(); doc.save(out); return out.getvalue()


def build_normal_word_from_pdf(data, title="", watermark="", page_size="A4", orientation="Portrait", first_page_text=""):
    # Exact-appearance PDF -> Word conversion. First apply the same heading,
    # watermark and page-number overlay to the PDF, then place each final page
    # into Word as an image. This guarantees equations/symbols survive exactly.
    if fitz is None:
        raise RuntimeError("PDF support requires PyMuPDF.")
    branded_pdf=build_normal_pdf_from_pdf(data,title,watermark,page_size,orientation,first_page_text)
    src=fitz.open(stream=branded_pdf,filetype="pdf")
    doc=Document(); section=doc.sections[0]
    _configure_docx_section(section,page_size,orientation)
    for idx,page in enumerate(src):
        if idx>0: doc.add_page_break()
        pix=page.get_pixmap(matrix=fitz.Matrix(1.8,1.8),alpha=False)
        img=io.BytesIO(pix.tobytes("png"))
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        run=p.add_run(); run.add_picture(img,width=section.page_width-section.left_margin-section.right_margin)
    src.close()
    out=io.BytesIO(); doc.save(out); return out.getvalue()


def build_normal_pdf_from_pdf(data, title="", watermark="", page_size="A4", orientation="Portrait", first_page_text=""):
    # Keep the original PDF page rendering intact, then overlay the standard
    # branding. This is the safest route for equations, symbols, diagrams, and
    # unusual fonts because nothing is re-parsed.
    if fitz is None:
        raise RuntimeError("PDF support requires PyMuPDF.")
    src=fitz.open(stream=data,filetype="pdf")
    for i,page in enumerate(src):
        rect=page.rect
        if i==0 and title.strip():
            page.insert_textbox(fitz.Rect(24,18,rect.width-24,42),title.strip(),
                                fontname="helv",fontsize=16,color=(0.05,0.10,0.18),
                                fill_opacity=1,align=1,overlay=True)
        if i==0:
            lines=[x.strip() for x in str(first_page_text or "").splitlines() if x.strip()]
            if lines:
                y=48
                for line in lines[:6]:
                    page.insert_textbox(fitz.Rect(24,y-10,rect.width-24,y+3),line,
                                        fontname="helv",fontsize=10,color=(0.05,0.10,0.18),
                                        fill_opacity=1,align=1,overlay=True)
                    y+=13
        wm=[x.strip() for x in str(watermark or "").splitlines() if x.strip()]
        if wm:
            center=fitz.Point(rect.width/2,rect.height/2)
            # Put multiple watermark lines around the centre.
            start=-18*(len(wm)-1)/2
            for j,line in enumerate(wm[:5]):
                page.insert_text((center.x-90,center.y+start+j*18),line,fontname="helv",fontsize=24,
                                 color=(0.55,0.60,0.68),fill_opacity=0.035,rotate=35,overlay=True)
        page.insert_text((rect.width-55,rect.height-14),f"{i+1}/{len(src)}",fontname="helv",fontsize=8,
                         color=(0.44,0.50,0.59),fill_opacity=1,overlay=True)
    out=src.tobytes(garbage=4,deflate=True); src.close(); return out


def _normal_text_from_docx(data):
    doc=Document(io.BytesIO(data)); chunks=[]
    for p in doc.paragraphs:
        t=clean_text(p.text)
        if t: chunks.append(t)
    for table in doc.tables:
        for row in table.rows:
            vals=[clean_text(c.text) for c in row.cells]
            if any(vals): chunks.append("    ".join(v for v in vals if v))
    return "\n".join(chunks)


def build_normal_pdf_from_text(text, title="", watermark="", page_size="A4", orientation="Portrait", first_page_text=""):
    buffer=io.BytesIO(); pagesize=_resolve_page_size(page_size,orientation)
    styles=make_styles()
    doc=SimpleDocTemplate(buffer,pagesize=pagesize,rightMargin=18*mm,leftMargin=18*mm,topMargin=18*mm,bottomMargin=15*mm,title=title or "Formatted Document")
    story=[]
    lines=[x.rstrip() for x in str(text or "").splitlines()]
    first_lines=[x.strip() for x in str(first_page_text or "").splitlines() if x.strip()]
    for line in first_lines[:8]: story.append(Paragraph(escape_xml(line),styles["first_page_header"]))
    if first_lines: story.append(Spacer(1,5*mm))
    if title.strip(): story.append(Paragraph(escape_xml(title.strip()),styles["title"]))
    body_style=ParagraphStyle("NormalDoc",parent=getSampleStyleSheet()["BodyText"],fontName="Helvetica",fontSize=10.5,leading=14,spaceAfter=6)
    for line in lines:
        if not line.strip(): story.append(Spacer(1,3*mm)); continue
        story.append(Paragraph(escape_xml(line.strip()),body_style))
    cb=_watermark_callback(watermark,pagesize)
    doc.build(story,onFirstPage=cb,onLaterPages=cb,canvasmaker=NumberedCanvas)
    return buffer.getvalue()


def build_normal_pdf_from_docx(data, title="", watermark="", page_size="A4", orientation="Portrait", first_page_text=""):
    text=_normal_text_from_docx(data)
    return build_normal_pdf_from_text(text,title,watermark,page_size,orientation,first_page_text)


def build_normal_outputs(uploaded_file, source_ext, output_formats, title="", watermark="", page_size="A4", orientation="Portrait", first_page_text=""):
    data=uploaded_file.getvalue()
    pdf_bytes=docx_bytes=None
    if source_ext==".pdf":
        if "PDF" in output_formats: pdf_bytes=build_normal_pdf_from_pdf(data,title,watermark,page_size,orientation,first_page_text)
        if "Word (.docx)" in output_formats: docx_bytes=build_normal_word_from_pdf(data,title,watermark,page_size,orientation)
    elif source_ext==".docx":
        if "PDF" in output_formats: pdf_bytes=build_normal_pdf_from_docx(data,title,watermark,page_size,orientation,first_page_text)
        if "Word (.docx)" in output_formats: docx_bytes=build_normal_word_from_docx(data,title,watermark,page_size,orientation,first_page_text)
    else:
        text=data.decode("utf-8-sig",errors="replace")
        if "PDF" in output_formats: pdf_bytes=build_normal_pdf_from_text(text,title,watermark,page_size,orientation,first_page_text)
        if "Word (.docx)" in output_formats: docx_bytes=build_normal_word_from_text(text,title,watermark,page_size,orientation,first_page_text)
    return pdf_bytes,docx_bytes


# -----------------------------
# Streamlit UI — polished version
# -----------------------------
import base64

st.set_page_config(
    page_title="XD DOC FORGE — PDF Generator",
    page_icon="assets/logo.svg",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------- Theme ----------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    --navy: #10233f;
    --blue: #2f6df6;
    --purple: #7b4ff6;
    --green: #18a957;
    --text: #16243a;
    --muted: #6c7890;
    --line: #e4eaf3;
    --card: #ffffff;
    --soft: #f5f8fd;
}

html, body, [class*="css"] {
    font-family: Inter, system-ui, sans-serif;
}

.stApp {
    background:
      radial-gradient(circle at 88% 8%, rgba(123,79,246,.10), transparent 25%),
      radial-gradient(circle at 10% 0%, rgba(47,109,246,.09), transparent 28%),
      #f6f8fc;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #10233f 0%, #132c4f 55%, #0c1c33 100%);
    border-right: 0;
}

section[data-testid="stSidebar"] * {
    color: #eaf1ff !important;
}

.brand {
    display:flex; align-items:center; gap:12px; padding:8px 4px 20px;
}
.brand img { width:52px; height:52px; }
.brand-title { font-size:23px; font-weight:800; letter-spacing:-.5px; }
.brand-sub { font-size:11px; opacity:.72; margin-top:2px; }

.side-pill {
    padding:12px 14px; border-radius:12px; margin:7px 0;
    background:rgba(255,255,255,.055); border:1px solid rgba(255,255,255,.06);
    font-size:13px;
}
.side-pill.active {
    background:linear-gradient(90deg,#2f6df6,#5a62ee);
    box-shadow:0 10px 25px rgba(0,0,0,.18);
}

.hero {
    background:rgba(255,255,255,.86);
    border:1px solid rgba(228,234,243,.95);
    border-radius:24px;
    padding:26px 30px;
    box-shadow:0 14px 45px rgba(26,48,84,.08);
    margin-bottom:18px;
}
.hero-row { display:flex; justify-content:space-between; align-items:center; gap:20px; }
.hero h1 { margin:0; font-size:38px; color:#10233f; letter-spacing:-1.5px; }
.hero h1 span { color:#416ef2; }
.hero p { margin:8px 0 0; color:#68758b; font-size:15px; }
.support {
    background:#eef3ff; border-radius:16px; padding:14px 18px;
    color:#243b63; font-size:13px; min-width:170px;
}
.feature-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin-bottom:18px; }
.feature {
    background:rgba(255,255,255,.84); border:1px solid #e4eaf3;
    border-radius:17px; padding:15px 16px; box-shadow:0 7px 25px rgba(28,47,78,.05);
}
.feature .icon { font-size:23px; }
.feature b { display:block; margin-top:6px; color:#152845; font-size:13px; }
.feature small { color:#7b879a; }

.card {
    background:rgba(255,255,255,.92);
    border:1px solid #e3e9f2; border-radius:20px;
    padding:20px; box-shadow:0 9px 30px rgba(29,48,80,.06);
    height:100%;
}
.card-title { font-size:18px; font-weight:750; color:#132743; margin-bottom:13px; }
.step {
    display:inline-flex; align-items:center; justify-content:center;
    width:29px; height:29px; border-radius:50%;
    color:#fff; background:linear-gradient(135deg,#2f6df6,#6c54ed);
    font-weight:800; margin-right:8px; font-size:13px;
}
.drop-hint {
    border:1.5px dashed #9db8ef; border-radius:16px; padding:28px 18px;
    text-align:center; background:#f8fbff; margin:8px 0 12px;
}
.drop-hint .cloud { font-size:35px; }
.drop-hint b { display:block; color:#213657; margin-top:6px; }
.drop-hint span { color:#7c889b; font-size:12px; }

.validation-ok {
    background:#eefbf3; border:1px solid #c9efd8; color:#14763d;
    border-radius:15px; padding:13px 15px; line-height:1.75; font-size:13px;
}
.validation-error {
    background:#fff2f2; border:1px solid #ffd1d1; color:#a32a2a;
    border-radius:15px; padding:13px 15px; line-height:1.65; font-size:13px;
}
.format-box {
    background:linear-gradient(135deg,#f4f6ff,#f7f1ff);
    border:1px solid #dddffb; border-radius:15px; padding:14px;
}
.format-box b { color:#263a65; }
.mini-table {
    margin-top:12px; border:1px solid #dfe6f0; border-radius:10px;
    overflow:hidden; font-size:10px; background:white;
}
.mini-table table { width:100%; border-collapse:collapse; }
.mini-table th { background:#f1f4f9; padding:7px; }
.mini-table td { padding:7px; border-top:1px solid #e8edf4; }
.mini-table .correct { background:#dff7e8; font-weight:700; }

.generate-bar {
    margin-top:18px; padding:18px 22px; border-radius:18px;
    background:linear-gradient(100deg,#2f6df6,#7650ee);
    color:white; font-size:19px; font-weight:800; text-align:center;
    box-shadow:0 12px 30px rgba(70,83,230,.25);
}

div.stButton > button {
    border-radius:12px; font-weight:700; min-height:42px;
}
div.stDownloadButton > button {
    border-radius:12px; font-weight:700; width:100%; min-height:45px;
}

@media (max-width: 900px) {
    .feature-grid { grid-template-columns:repeat(2,1fr); }
    .hero h1 { font-size:28px; }
}

/* Streamlit Cloud/theme compatibility: keep the main workspace readable
   even when the viewer/browser is using a dark Streamlit theme. */
[data-testid="stAppViewContainer"] {
    background: #f6f8fc !important;
}
[data-testid="stAppViewContainer"] .main {
    background: transparent !important;
}
[data-testid="stAppViewContainer"] .main .block-container {
    color: #16243a !important;
}
[data-testid="stAppViewContainer"] .main p,
[data-testid="stAppViewContainer"] .main li,
[data-testid="stAppViewContainer"] .main label,
[data-testid="stAppViewContainer"] .main [data-testid="stWidgetLabel"],
[data-testid="stAppViewContainer"] .main [data-testid="stWidgetLabel"] *,
[data-testid="stAppViewContainer"] .main [data-testid="stMarkdownContainer"],
[data-testid="stAppViewContainer"] .main [data-testid="stMarkdownContainer"] *,
[data-testid="stAppViewContainer"] .main h1,
[data-testid="stAppViewContainer"] .main h2,
[data-testid="stAppViewContainer"] .main h3,
[data-testid="stAppViewContainer"] .main h4,
[data-testid="stAppViewContainer"] .main h5,
[data-testid="stAppViewContainer"] .main h6 {
    color: #16243a !important;
}
[data-testid="stAppViewContainer"] .main .stCaption,
[data-testid="stAppViewContainer"] .main [data-testid="stCaptionContainer"],
[data-testid="stAppViewContainer"] .main [data-testid="stCaptionContainer"] * {
    color: #68758b !important;
}

/* Form controls */
[data-testid="stAppViewContainer"] .main input,
[data-testid="stAppViewContainer"] .main textarea,
[data-testid="stAppViewContainer"] .main [role="combobox"],
[data-testid="stAppViewContainer"] .main [data-baseweb="select"] > div {
    color: #16243a !important;
    background-color: #ffffff !important;
}
[data-testid="stAppViewContainer"] .main input::placeholder,
[data-testid="stAppViewContainer"] .main textarea::placeholder {
    color: #8a95a8 !important;
    opacity: 1 !important;
}
[data-testid="stAppViewContainer"] .main [data-baseweb="select"] *,
[data-testid="stAppViewContainer"] .main [role="option"] {
    color: #16243a !important;
}
[data-testid="stAppViewContainer"] .main [role="radiogroup"] label,
[data-testid="stAppViewContainer"] .main [role="radiogroup"] label * {
    color: #16243a !important;
}
[data-testid="stAppViewContainer"] .main [data-testid="stFileUploader"] *,
[data-testid="stAppViewContainer"] .main [data-testid="stFileUploaderDropzone"] * {
    color: #16243a !important;
}
[data-testid="stAppViewContainer"] .main .stButton button,
[data-testid="stAppViewContainer"] .main .stDownloadButton button {
    color: #ffffff !important;
}
[data-testid="stAppViewContainer"] .main [data-testid="stExpander"] summary,
[data-testid="stAppViewContainer"] .main [data-testid="stExpander"] summary * {
    color: #16243a !important;
}

.stTextArea textarea {
    border-radius: 14px !important;
    border: 1px solid #cfd9e8 !important;
    background: #fbfdff !important;
    font-family: Consolas, "Courier New", monospace !important;
    font-size: 13px !important;
}

.validate-panel {
    background: linear-gradient(90deg, rgba(55,146,255,.10), rgba(116,88,255,.10));
    border: 1px solid rgba(80,120,220,.25);
    border-radius: 12px;
    padding: 12px 14px;
    margin: 8px 0 10px;
}
.validate-panel span { color:#667085; font-size:12px; }
</style>
""", unsafe_allow_html=True)

if "generated_question_pdf" not in st.session_state:
    st.session_state.generated_question_pdf = None
if "generated_answer_pdf" not in st.session_state:
    st.session_state.generated_answer_pdf = None
if "generated_question_filename" not in st.session_state:
    st.session_state.generated_question_filename = ""
if "generated_answer_filename" not in st.session_state:
    st.session_state.generated_answer_filename = ""
if "generated_question_docx" not in st.session_state:
    st.session_state.generated_question_docx = None
if "generated_answer_docx" not in st.session_state:
    st.session_state.generated_answer_docx = None
if "generated_question_docx_filename" not in st.session_state:
    st.session_state.generated_question_docx_filename = ""
if "generated_answer_docx_filename" not in st.session_state:
    st.session_state.generated_answer_docx_filename = ""
if "detected_headings" not in st.session_state:
    st.session_state.detected_headings = []
if "heading_rules" not in st.session_state:
    st.session_state.heading_rules = []
if "heading_text" not in st.session_state:
    st.session_state.heading_text = ""
if "show_heading_dialog" not in st.session_state:
    st.session_state.show_heading_dialog = False
if "pending_generate" not in st.session_state:
    st.session_state.pending_generate = False
if "numbering_mode" not in st.session_state:
    st.session_state.numbering_mode = "Use source numbering"
if "numbering_start" not in st.session_state:
    st.session_state.numbering_start = 1
if "heading_basis" not in st.session_state:
    st.session_state.heading_basis = "Output numbering"
if "topic_editor_rows" not in st.session_state:
    st.session_state.topic_editor_rows = []
if "normal_pdf" not in st.session_state:
    st.session_state.normal_pdf = None
if "normal_docx" not in st.session_state:
    st.session_state.normal_docx = None
if "normal_pdf_name" not in st.session_state:
    st.session_state.normal_pdf_name = ""
if "normal_docx_name" not in st.session_state:
    st.session_state.normal_docx_name = ""

# Sidebar
logo_b64 = base64.b64encode(Path("assets/logo.svg").read_bytes()).decode()
st.sidebar.markdown(
    f"""
    <div class="brand">
      <img src="data:image/svg+xml;base64,{logo_b64}">
      <div><div class="brand-title">XD DOC FORGE</div>
      <div class="brand-sub">Questions + Documents → Perfect Files</div></div>
    </div>
    <div class="side-pill active">⌂ &nbsp; Home</div>
    <div class="side-pill">▣ &nbsp; Generate / Convert</div>
    <div class="side-pill">☷ &nbsp; Preview & Validate</div>
    <div class="side-pill">⚙ &nbsp; PDF & Word Settings</div>
    """,
    unsafe_allow_html=True,
)
st.sidebar.markdown("---")
st.sidebar.caption("Local processing • Your files stay on this PC")
st.sidebar.caption("XD DOC FORGE")

# Hero
st.markdown("""
<div class="hero">
  <div class="hero-row">
    <div>
      <h1>XD DOC <span>FORGE</span></h1>
      <p>Create polished Question Sets or format ordinary Word, PDF and text documents.</p>
    </div>
    <div class="support"><b>Supports</b><br>📘 .docx &nbsp;&nbsp; 📄 .pdf &nbsp;&nbsp; 📄 .txt<br>📝 PDF + Word output • MCQ + Normal</div>
  </div>
</div>
<div style="background:#eef4ff;border:1px solid #d7e3fb;border-radius:14px;padding:11px 15px;margin:-6px 0 17px;color:#405777;font-size:12px;">
<b>Workflow:</b> Input → Validate → Preview → Name → Generate → Download &nbsp; • &nbsp; One-line options and answers supported.
</div>
<div class="feature-grid">
  <div class="feature"><div class="icon">🛡️</div><b>Strict Validation</b><small>No silent guessing</small></div>
  <div class="feature"><div class="icon">⚡</div><b>Automatic Extraction</b><small>Questions, options & answers</small></div>
  <div class="feature"><div class="icon">📑</div><b>Beautiful PDFs</b><small>Question + Answer Set</small></div>
  <div class="feature"><div class="icon">🎛️</div><b>Customizable</b><small>Format and layout controls</small></div>
</div>
""", unsafe_allow_html=True)

# Choose the workflow before the MCQ-specific controls. Normal Document mode
# intentionally bypasses MCQ validation and simply formats ordinary documents.
conversion_mode = st.radio(
    "What are you converting?",
    ["MCQ / Question Set", "Normal Document"],
    horizontal=True,
    index=0,
    help="Use MCQ mode for question/option/answer tables. Use Normal Document for ordinary Word, PDF or TXT files.",
)

if conversion_mode == "Normal Document":
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title"><span class="step">1</span>Normal Document Converter</div>', unsafe_allow_html=True)
    normal_file = st.file_uploader(
        "Choose a normal Word, PDF or text file",
        type=["docx", "pdf", "txt"],
        label_visibility="collapsed",
        key="normal_document_upload",
    )
    st.markdown(
        '<div class="drop-hint"><div class="cloud">📄</div><b>Drop an ordinary document here</b>'
        '<span>.docx / .pdf / .txt • no MCQ structure required</span></div>',
        unsafe_allow_html=True,
    )
    if normal_file:
        st.success(f"📄 {normal_file.name} • Ready to format")
        ext=Path(normal_file.name).suffix.lower()
        if ext==".pdf":
            st.info("🧮 PDF detected: PDF output can preserve the original page exactly, including equations, symbols, diagrams and unusual fonts. PDF → Word uses page images for visual fidelity, so the contents are not individually editable.")
        elif ext==".docx":
            st.info("📝 Word detected: the original Word formatting and tables are retained as far as the Word format allows, with the standard branding added.")
        else:
            st.info("📄 Text detected: the text will be laid out cleanly using the same page, heading and watermark configuration.")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card" style="margin-top:16px;">', unsafe_allow_html=True)
    st.markdown('<div class="card-title"><span class="step">2</span>Normal Document Settings</div>', unsafe_allow_html=True)
    normal_base=st.text_input(
        "Base name", value="", placeholder="Optional — leave blank to use the source filename",
        key="normal_base_name", help="If blank, the uploaded filename is used automatically."
    )
    normal_title=st.text_input(
        "Document title", value="", placeholder="Optional — leave blank for no extra title",
        key="normal_title", help="If entered, this is shown centered and bold at the top of the first page."
    )
    normal_first_page=st.text_area(
        "First-page heading / contact text", value=DEFAULT_FIRST_PAGE_TEXT, height=120,
        key="normal_first_page_text",
        help="Default heading. It appears at the beginning of the first page only."
    )
    normal_watermark=st.text_area(
        "Watermark", value="SRUTI ACADEMY\n6290000542", height=90,
        key="normal_watermark_text",
        help="Default watermark. It is applied lightly to every page."
    )
    nc1,nc2,nc3=st.columns(3)
    with nc1:
        normal_page_size=st.selectbox("Page size",["A4","A3","A5","Letter","Legal"],index=0,key="normal_page_size")
    with nc2:
        normal_orientation=st.radio("Orientation",["Portrait","Landscape"],index=0,horizontal=True,key="normal_orientation")
    with nc3:
        normal_formats=st.multiselect("Output format",["PDF","Word (.docx)"],default=["PDF"],key="normal_formats")
    st.caption("Portrait is selected by default. Page numbers are added at the bottom as 1/6, 2/6, 3/6… and the heading appears only at the top of page 1.")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="generate-bar">📄 &nbsp; Ready to Format Your Document</div>', unsafe_allow_html=True)
    st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)
    normal_generate=st.button("✨  FORMAT + GENERATE DOCUMENT",type="primary",use_container_width=True,key="generate_normal_document")
    if normal_generate:
        if not normal_file:
            st.error("❌ Please upload a PDF, Word or text file first.")
        elif not normal_formats:
            st.error("❌ Select at least one output format.")
        else:
            base=clean_text(normal_base) or Path(normal_file.name).stem
            ext=Path(normal_file.name).suffix.lower()
            with st.spinner("Formatting your document…"):
                try:
                    npdf,ndocx=build_normal_outputs(normal_file,ext,normal_formats,
                        title=normal_title.strip(),watermark=normal_watermark.strip(),
                        page_size=normal_page_size,orientation=normal_orientation,
                        first_page_text=normal_first_page.strip())
                    st.session_state.normal_pdf=npdf; st.session_state.normal_docx=ndocx
                    st.session_state.normal_pdf_name=f"{base} - Formatted.pdf"
                    st.session_state.normal_docx_name=f"{base} - Formatted.docx"
                    st.success("🎉 Your formatted document is ready.")
                except Exception as exc:
                    st.error(f"❌ Could not format the document: {exc}")

    if st.session_state.get("normal_pdf") or st.session_state.get("normal_docx"):
        st.markdown("### 📥 Your formatted file is ready")
        nd1,nd2=st.columns(2)
        with nd1:
            if st.session_state.get("normal_pdf"):
                st.download_button("⬇️ Download PDF",st.session_state.normal_pdf,
                    file_name=st.session_state.normal_pdf_name,mime="application/pdf",
                    use_container_width=True,key="download_normal_pdf")
        with nd2:
            if st.session_state.get("normal_docx"):
                st.download_button("⬇️ Download Word",st.session_state.normal_docx,
                    file_name=st.session_state.normal_docx_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,key="download_normal_docx")
    st.markdown(
        "<div style='text-align:center;color:#8a95a8;font-size:11px;margin:25px 0 5px;'>"
        "XD DOC FORGE • Normal Document Converter • Local-first"
        "</div>", unsafe_allow_html=True
    )
    st.stop()

left, right = st.columns([1, 1.12], gap="large")

with left:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title"><span class="step">1</span>Upload Your File</div>', unsafe_allow_html=True)
    source_mode = st.radio(
        "Input source",
        ["Upload Word/Text file", "Paste questions"],
        horizontal=True,
        label_visibility="collapsed",
    )

    uploaded = None
    pasted_text = ""

    if source_mode == "Upload Word/Text file":
        uploaded = st.file_uploader(
            "Choose a Word, PDF or text file",
            type=["docx", "pdf", "txt"],
            label_visibility="collapsed",
        )
        st.markdown(
            '<div class="drop-hint"><div class="cloud">☁️</div><b>Drop your file here</b>'
            '<span>or use the upload button above • .docx / .pdf / .txt</span></div>',
            unsafe_allow_html=True,
        )
        if uploaded:
            st.success(f"📄 {uploaded.name}  •  Ready to process")
    else:
        pasted_text = st.text_area(
            "Paste your MCQs here",
            height=260,
            placeholder=(
                "Paste your questions here, for example:\n\n"
                "1. Who was the last ruler of the Lodi dynasty?\n"
                "A) Sikandar Lodi\n"
                "B) Bahlul Lodi\n"
                "C) Ibrahim Lodi\n"
                "D) Daulat Khan Lodi\n"
                "Answer: C"
            ),
            label_visibility="collapsed",
        )
        st.caption("You can paste directly with Ctrl+V. The same validation rules are applied.")

current_source_signature = (
    source_mode,
    uploaded.name if uploaded else "",
    pasted_text if source_mode == "Paste questions" else "",
)

# Keep validation controls visible on every Streamlit rerun. Previously this
# entire section was accidentally nested under the "source changed" block,
# which meant the Validate button disappeared immediately after the first
# interaction.
if st.session_state.get("_source_signature") != current_source_signature:
    st.session_state.generated_question_pdf = None
    st.session_state.generated_answer_pdf = None
    st.session_state.generated_question_filename = ""
    st.session_state.generated_answer_filename = ""
    st.session_state.questions = None
    st.session_state.errors = None
    st.session_state.detected_headings = []
    st.session_state.heading_rules = []
    st.session_state.heading_text = ""
    st.session_state.topic_editor_rows = []
    st.session_state._source_signature = current_source_signature

# Sample template is always available.
st.download_button(
    "📋 Download Sample TXT Template",
    data="""1. Sample question goes here?\nA) Option A\nB) Option B\nC) Option C\nD) Option D\nAnswer: A\n\n2. Another question?\nA) Option A\nB) Option B\nC) Option C\nD) Option D\nAnswer: C\n""",
    file_name="MCQ_Template.txt",
    mime="text/plain",
    use_container_width=True,
)
st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
st.markdown('<div class="card">', unsafe_allow_html=True)
st.markdown('<div class="card-title"><span class="step">3</span>Validate & Preview</div>', unsafe_allow_html=True)

st.markdown(
    '<div class="validate-panel"><b>Step 2 — Check your questions before generating</b>'
    '<br><span>Click Validate & Preview. Any problem will be shown here in red with the exact question number.</span></div>',
    unsafe_allow_html=True,
)
validate_clicked = st.button(
    "🔎  VALIDATE & PREVIEW QUESTIONS",
    use_container_width=True,
    type="primary",
    key="validate_questions_button",
)

if "questions" not in st.session_state:
    st.session_state.questions = None
if "errors" not in st.session_state:
    st.session_state.errors = None

if validate_clicked:
    has_input = bool(uploaded) if source_mode == "Upload Word/Text file" else bool(pasted_text.strip())
    if not has_input:
        st.session_state.questions = None
        st.session_state.errors = ["Please upload a file or paste your questions first."]
    else:
        try:
            if source_mode == "Paste questions":
                qs = parse_plain_text(pasted_text)
                detected = detect_heading_lines(pasted_text)
            elif uploaded.name.lower().endswith(".docx"):
                qs = parse_docx(uploaded)
                detected = detect_docx_headings(uploaded)
            elif uploaded.name.lower().endswith(".pdf"):
                qs, raw_text = parse_pdf(uploaded)
                detected = detect_heading_lines(raw_text)
                st.info("📄 PDF text extracted. Mathematical symbols that are stored as selectable PDF text are preserved where possible. Complex/vector/image equations may need equation-image preservation in a future dedicated mode.")
            else:
                raw_text = uploaded.getvalue().decode("utf-8-sig", errors="replace")
                qs = parse_plain_text(raw_text)
                detected = detect_heading_lines(raw_text)
            st.session_state.detected_headings = detected
            st.session_state.heading_text = heading_text_from_candidates(detected)
            errs = validate_questions(qs)
            if not qs:
                errs = ["No questions were detected. Check the input format and try again."]
            st.session_state.questions = qs
            st.session_state.errors = errs
        except Exception as exc:
            st.session_state.questions = []
            st.session_state.errors = [f"Could not process input: {exc}"]

qs = st.session_state.questions
errs = st.session_state.errors

if qs and not errs:
    st.markdown(
        f"""<div class="validation-ok">
        <b>✓ {len(qs)} questions detected</b><br>
        ✓ All questions have 4 options<br>
        ✓ {len(qs)} answers found<br>
        ✓ No validation errors
        </div>""",
        unsafe_allow_html=True,
    )
elif errs:
    st.markdown(
        '<div class="validation-error"><b>Validation failed</b><br>'
        + "<br>".join("❌ " + escape_xml(e) for e in errs[:20])
        + '</div>',
        unsafe_allow_html=True,
    )
else:
    st.info("Upload your file and click **VALIDATE & PREVIEW QUESTIONS**.")
st.markdown('</div>', unsafe_allow_html=True)

with right:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown('<div class="card-title"><span class="step">2</span>Output Settings</div>', unsafe_allow_html=True)

    output_name = st.text_input(
        "Base name for PDFs",
        value="",
        placeholder="Optional — leave blank for Question Set / Answer Set",
        help="Optional. If left blank, files are saved as Question Set.pdf / Answer Set.pdf (or .docx)."
    )

    st.markdown('<div class="format-box"><b>Format 3 — Table PDFs</b><br>'
                '<span style="color:#6f7890;font-size:12px;">Question Set: Q.No. + Question + A/B/C/D. '
                'Answer Set: the same table plus a dedicated Correct Answer column.</span>'
                '<div class="mini-table"><table><tr><th>Q.No.</th><th>Question</th><th>A</th><th>B</th><th>C</th><th>D</th><th>Correct</th></tr>'
                '<tr><td>1</td><td>Sample question?</td><td>Option A</td><td>Option B</td><td>Option C</td><td>Option D</td><td><b>B — Option B</b></td></tr></table></div>'
                '</div>', unsafe_allow_html=True)

    st.markdown("#### 🖨️ Page & Output Settings")
    page_size = st.selectbox("Page size", ["A4", "A3", "A5", "Letter", "Legal"], index=0)
    orientation = st.radio("Orientation", ["Landscape", "Portrait"], horizontal=True, index=1)
    output_formats = st.multiselect("Download format", ["PDF", "Word (.docx)"], default=["PDF"])

    st.markdown("#### 📝 First-page heading / contact text")
    first_page_text = st.text_area(
        "Show this at the top of page 1 only",
        value="SRUTI ACADEMY\nBARASAT ASAWANIPALLY SCHOOL ROAD\nKOLKATA-700124\nMOB NO.6290000542",
        placeholder="e.g.\nSRUTI ACADEMY\nBARASAT ASAWANIPALLY SCHOOL ROAD\nKOLKATA-700124\nMOB NO.6290000542",
        height=120,
        help="This is separate from the diagonal watermark. It appears only at the top of the first page of each Question Set and Answer Set.",
        key="first_page_text",
    )

    st.markdown("#### 🔢 Question Numbering")
    numbering_mode = st.radio(
        "Numbering",
        ["Use source numbering", "Start from 1", "Start from custom number"],
        index=0,
        help="Keep the original numbers, restart at 1, or renumber from any number such as 100 or 501."
    )
    numbering_start = 1
    if numbering_mode == "Start from custom number":
        numbering_start = st.number_input("First output question number", min_value=1, value=100, step=1)
    st.session_state.numbering_mode = numbering_mode
    st.session_state.numbering_start = int(numbering_start)

    st.markdown("#### 🧩 Topic Headings")
    if st.session_state.get("detected_headings"):
        st.success(f"Detected {len(st.session_state.detected_headings)} possible topic heading(s). A confirmation popup will appear before generation.")
    else:
        st.caption("No obvious topic headings were detected. You can still add custom headings in the generation popup.")

    st.markdown("#### 💧 Watermark")
    watermark = st.text_area(
        "Watermark text",
        value="SRUTI ACADEMY\n6290000542",
        height=100,
        help="Multiline watermark. The default is SRUTI ACADEMY / 6290000542. Press Enter or Shift+Enter for a new line. It is printed very lightly on every PDF page and included on every Word page.",
        key="watermark_text",
    )
    if watermark.strip():
        import html
        preview_text = html.escape(watermark.strip(), quote=False).replace("\n", "<br>")
        st.markdown(
            '<div style="height:150px;border:1px solid #dce4ef;border-radius:16px;'
            'background:linear-gradient(135deg,#f8fbff,#ffffff);display:flex;'
            'align-items:center;justify-content:center;overflow:hidden;position:relative;">'
            '<div style="transform:rotate(-18deg);font-size:22px;font-weight:700;'
            'line-height:1.35;letter-spacing:1.5px;text-align:center;'
            'color:rgba(65,78,105,.08);white-space:nowrap;">'
            + preview_text +
            '</div><div style="position:absolute;bottom:8px;right:10px;'
            'font-size:10px;color:#98a3b4;">Watermark preview • very light</div></div>',
            unsafe_allow_html=True,
        )
    else:
        st.caption("Enter watermark text above to see a live preview. You can use multiple lines.")

    st.markdown('</div>', unsafe_allow_html=True)

# Preview
qs = st.session_state.questions
errs = st.session_state.errors
if qs and not errs:
    st.markdown('<div style="height:14px"></div>', unsafe_allow_html=True)
    st.markdown('<div class="card">', unsafe_allow_html=True)
    pleft, pright = st.columns([4, 1])
    with pleft:
        st.markdown('<div class="card-title"><span class="step">4</span>Preview</div>', unsafe_allow_html=True)
    with pright:
        show_n = st.selectbox("Show", [5, 10, 20, "All"], index=0, label_visibility="collapsed")
    n = len(qs) if show_n == "All" else int(show_n)

    display_qs = apply_numbering(qs, st.session_state.numbering_mode, st.session_state.numbering_start)
    preview_rows = []
    for q in display_qs[:n]:
        option_map = {l: t for l, t in q["options"]}
        preview_rows.append({
            "Q.No.": q["number"],
            "Question": q["question"],
            "A": option_map.get("A", ""),
            "B": option_map.get("B", ""),
            "C": option_map.get("C", ""),
            "D": option_map.get("D", ""),
            "Answer": q["answer"],
        })
    st.dataframe(preview_rows, use_container_width=True, hide_index=True)
    st.caption(f"Showing {min(n, len(display_qs))} of {len(display_qs)} questions.")
    st.markdown('</div>', unsafe_allow_html=True)

# Generate
st.markdown('<div class="generate-bar">📄 &nbsp; Ready to Generate Your PDFs</div>', unsafe_allow_html=True)
st.markdown('<div style="height:10px"></div>', unsafe_allow_html=True)

can_generate = bool(qs and not errs)

# Keep the button clickable at all times. The app explains the exact reason
# instead of silently disabling the control.
generate_clicked = st.button(
    "✨  GENERATE QUESTION SET + ANSWER SET",
    type="primary",
    use_container_width=True,
    key="generate_pdfs_button",
)

if not can_generate:
    if not qs:
        st.info("ℹ️ Nothing has been validated yet. Upload/paste your MCQs, then click **VALIDATE & PREVIEW QUESTIONS**.")
    elif errs:
        st.error("❌ Cannot generate yet — validation found errors. See the red **Validation failed** box above.")
    elif not output_name.strip():
        st.warning("⚠️ Enter an output name first. The button is still clickable and will tell you what is missing.")

if generate_clicked:
    if not qs:
        st.error("❌ Please upload/paste the MCQs and click **VALIDATE & PREVIEW QUESTIONS** first.")
    elif errs:
        st.error("❌ PDF generation stopped because validation failed. Fix the errors shown above, then validate again.")
    elif not output_name.strip():
        st.error("❌ Please enter the output name in **Output Settings**.")
    elif not output_formats:
        st.error("❌ Select at least one download format: PDF or Word (.docx).")
    else:
        st.session_state.show_heading_dialog = True
        st.rerun()

if st.session_state.get("show_heading_dialog"):
    @st.dialog("Topic headings — interactive setup", width="large")
    def heading_dialog():
        st.markdown("### 🧩 Add topic headings")
        st.caption("Each row creates a separate table section. You can use the detected headings, edit them, or add your own.")

        detected=st.session_state.get("detected_headings",[])
        if detected:
            st.success(f"✨ {len(detected)} possible heading(s) detected from your source.")
            with st.expander("Show detected headings", expanded=True):
                for h in detected:
                    rng=(f"{h['start']}–{h['end']}" if h.get('start') is not None else "range not known")
                    st.write(f"• **{h['title']}** — {rng} — {h.get('confidence','Possible')}")
        else:
            st.info("No explicit headings were detected. You can add them manually below.")

        use_headings=st.radio("Topic headings", ["No headings", "Use headings"], horizontal=True, index=0, key="topic_mode_dialog")
        if use_headings == "No headings":
            st.session_state.heading_rules=[]
        else:
            basis_label=st.radio(
                "Ranges refer to",
                ["Output numbering", "Source numbering"],
                horizontal=True,
                index=0,
                help="Choose Output numbering if you renumber questions to 1/100/etc. Choose Source numbering if the topic ranges refer to the original question numbers."
            )
            basis = "output" if basis_label == "Output numbering" else "source"

            # Build editable rows from detected headings on first open.
            if not st.session_state.get("topic_editor_rows"):
                rows=[]
                for h in detected:
                    rows.append({"Use": True, "Topic": h.get("title",""), "Start": h.get("start") or "", "End": h.get("end") or ""})
                if not rows:
                    rows=[{"Use":True,"Topic":"","Start":"","End":""}]
                st.session_state.topic_editor_rows=rows

            st.markdown("#### Edit / add topics")
            df=pd.DataFrame(st.session_state.topic_editor_rows, columns=["Use","Topic","Start","End"])
            edited=st.data_editor(
                df,
                num_rows="dynamic",
                hide_index=True,
                use_container_width=True,
                column_config={
                    "Use": st.column_config.CheckboxColumn("✓", default=True, width="small"),
                    "Topic": st.column_config.TextColumn("Topic name", help="Example: India"),
                    "Start": st.column_config.NumberColumn("From Q.No.", min_value=1, step=1),
                    "End": st.column_config.NumberColumn("To Q.No.", min_value=1, step=1),
                },
                key="topic_editor",
            )
            st.session_state.topic_editor_rows=edited.to_dict("records")
            st.caption("Tip: for 200–225, enter Topic = India, From = 200, To = 225. Add as many rows as you need.")

            # Preview the sections before committing.
            preview_rules=[]; invalid=[]
            for idx,row in enumerate(st.session_state.topic_editor_rows,1):
                if not bool(row.get("Use",True)): continue
                title=clean_text(row.get("Topic",""))
                a=row.get("Start"); b=row.get("End")
                if not title and a in (None,"") and b in (None,""): continue
                try:
                    a=int(a); b=int(b)
                    if not title or a<1 or b<1:
                        raise ValueError
                    preview_rules.append({"start":a,"end":b,"title":title})
                except Exception:
                    invalid.append(idx)
            if invalid:
                st.warning("Complete the Topic / From / To fields in row(s): " + ", ".join(map(str,invalid)))
            elif preview_rules:
                st.markdown("**Section preview:**")
                for r in normalize_heading_rules(preview_rules):
                    st.markdown(f"**{r['title']}**  ·  Questions {r['start']}–{r['end']}")
            st.session_state.heading_basis=basis

        c1,c2=st.columns(2)
        with c1:
            if st.button("Cancel",use_container_width=True,key="topic_cancel"):
                st.session_state.show_heading_dialog=False; st.rerun()
        with c2:
            if st.button("Generate with these settings",type="primary",use_container_width=True,key="topic_generate"):
                if use_headings == "No headings":
                    st.session_state.heading_rules=[]
                else:
                    rules=[]; bad=[]
                    for idx,row in enumerate(st.session_state.topic_editor_rows,1):
                        if not bool(row.get("Use",True)): continue
                        title=clean_text(row.get("Topic","")); a=row.get("Start"); b=row.get("End")
                        if not title and a in (None,"") and b in (None,""): continue
                        try:
                            a=int(a); b=int(b)
                            if not title or a<1 or b<1: raise ValueError
                            rules.append({'start':a,'end':b,'title':title})
                        except Exception:
                            bad.append(idx)
                    if bad:
                        st.error("Please complete topic row(s): " + ", ".join(map(str,bad)))
                        return
                    st.session_state.heading_rules=normalize_heading_rules(rules)
                st.session_state.show_heading_dialog=False
                st.session_state.pending_generate=True
                st.rerun()
    heading_dialog()

if st.session_state.get("pending_generate"):
    st.session_state.pending_generate=False
    display_qs=apply_numbering(qs,st.session_state.numbering_mode,st.session_state.numbering_start)
    base=output_name.strip() or "MCQ"
    question_filename_pdf=f"{base} - Question Set.pdf"; answer_filename_pdf=f"{base} - Answer Set.pdf"
    question_filename_docx=f"{base} - Question Set.docx"; answer_filename_docx=f"{base} - Answer Set.docx"
    rules=st.session_state.get("heading_rules",[])
    with st.spinner("Building your files…"):
        if "PDF" in output_formats:
            question_pdf=build_question_pdf(display_qs,f"{base} - Question Set",watermark=watermark.strip(),page_size=page_size,orientation=orientation,heading_rules=rules,heading_basis=st.session_state.get("heading_basis","output"),first_page_text=first_page_text.strip())
            answer_pdf=build_answer_pdf(display_qs,f"{base} - Answer Set",watermark=watermark.strip(),page_size=page_size,orientation=orientation,heading_rules=rules,heading_basis=st.session_state.get("heading_basis","output"),first_page_text=first_page_text.strip())
        else:
            question_pdf=answer_pdf=None
        if "Word (.docx)" in output_formats:
            qdoc=build_word_doc(display_qs,f"{base} - Question Set",False,watermark.strip(),page_size,orientation,rules,st.session_state.get("heading_basis","output"),first_page_text.strip())
            adoc=build_word_doc(display_qs,f"{base} - Answer Set",True,watermark.strip(),page_size,orientation,rules,st.session_state.get("heading_basis","output"),first_page_text.strip())
            qbuf,abuf=io.BytesIO(),io.BytesIO(); qdoc.save(qbuf); adoc.save(abuf)
            question_docx,answer_docx=qbuf.getvalue(),abuf.getvalue()
        else:
            question_docx=answer_docx=None
    st.session_state.generated_question_pdf=question_pdf; st.session_state.generated_answer_pdf=answer_pdf
    st.session_state.generated_question_docx=question_docx; st.session_state.generated_answer_docx=answer_docx
    st.session_state.generated_question_filename=question_filename_pdf; st.session_state.generated_answer_filename=answer_filename_pdf
    st.session_state.generated_question_docx_filename=question_filename_docx; st.session_state.generated_answer_docx_filename=answer_filename_docx
    st.success("🎉 Your Question Set and Answer Set have been generated successfully.")

if st.session_state.get("generated_question_pdf") or st.session_state.get("generated_question_docx"):
    st.markdown("### 📥 Your files are ready")
    d1, d2 = st.columns(2)
    with d1:
        if st.session_state.get("generated_question_pdf"):
            st.download_button("⬇️ Question Set — PDF", data=st.session_state.generated_question_pdf,
                file_name=st.session_state.generated_question_filename, mime="application/pdf",
                use_container_width=True, key="download_question_pdf")
        if st.session_state.get("generated_question_docx"):
            st.download_button("⬇️ Question Set — Word", data=st.session_state.generated_question_docx,
                file_name=st.session_state.generated_question_docx_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, key="download_question_docx")
    with d2:
        if st.session_state.get("generated_answer_pdf"):
            st.download_button("⬇️ Answer Set — PDF", data=st.session_state.generated_answer_pdf,
                file_name=st.session_state.generated_answer_filename, mime="application/pdf",
                use_container_width=True, key="download_answer_pdf")
        if st.session_state.get("generated_answer_docx"):
            st.download_button("⬇️ Answer Set — Word", data=st.session_state.generated_answer_docx,
                file_name=st.session_state.generated_answer_docx_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, key="download_answer_docx")

st.markdown(
    "<div style='text-align:center;color:#8a95a8;font-size:11px;margin:25px 0 5px;'>"
    "XD DOC FORGE • Local-first • Designed for reliable question paper generation"
    "</div>",
    unsafe_allow_html=True,
)
