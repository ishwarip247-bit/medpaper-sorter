
import io
import json
import re
import os

import streamlit as st
import fitz
import pytesseract
from PIL import Image
from google import genai
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from xml.sax.saxutils import escape

# Windows Tesseract path

if os.name == "nt":
    TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(TESSERACT_PATH):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH

st.set_page_config(
    page_title="MedPaper Sorter",
    page_icon="📚",
    layout="wide"
)

st.title("📚 MedPaper Sorter")
st.write(
    "Upload your syllabus and question papers to organize "
    "questions by syllabus topic."
)

st.info(
    "Scanned PDFs are processed using OCR. Please review extracted "
    "text and AI classifications before relying on the results."
)


def extract_pdf_text(uploaded_file):
    """Extract selectable PDF text, using OCR when needed."""
    doc = fitz.open(stream=uploaded_file.getvalue(), filetype="pdf")
    pages = []

    for page_number, page in enumerate(doc, start=1):
        text = page.get_text("text").strip()

        if len(text) < 30:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            image = Image.open(io.BytesIO(pix.tobytes("png")))
            text = pytesseract.image_to_string(image, lang="eng").strip()

        pages.append(f"\n[Page {page_number}]\n{text}")

    doc.close()
    return "\n".join(pages)


def get_client():
    api_key = st.secrets.get("GEMINI_API_KEY", "")
    if not api_key:
        raise ValueError(
            "GEMINI_API_KEY is missing. Add it to "
            ".streamlit/secrets.toml."
        )
    return genai.Client(api_key=api_key)


def ask_gemini(client, prompt):
    """Ask Gemini for a JSON response."""
    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=prompt,
        config={"response_mime_type": "application/json"}
    )
    return json.loads(response.text)


def split_questions(text):
    """Split OCR text into approximate question-sized sections."""
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    questions = []
    current = []

    pattern = re.compile(
        r"^(?:Q(?:UESTION)?\s*)?\d{1,2}\s*[\.\):\-]"
        r"|^(?:\([a-zA-Z0-9]+\))"
        r"|^\d{1,2}\s*[.)]\s+"
    )

    for line in lines:
        if pattern.match(line) and current:
            questions.append(" ".join(current))
            current = [line]
        else:
            current.append(line)

    if current:
        questions.append(" ".join(current))

    return [q for q in questions if len(q) >= 8]


def classify_questions(client, syllabus, paper_texts):
    """Use uploaded syllabus as the source of topic labels."""
    prompt = f"""
You are a careful medical examination question organizer.

Use ONLY the syllabus supplied below to identify units, topics,
and subtopics. Do not invent syllabus topics.

SYLLABUS:
{syllabus[:30000]}

QUESTION PAPER TEXTS:
{json.dumps(paper_texts, ensure_ascii=False)[:45000]}

Tasks:
1. Identify the units and topics stated in the syllabus.
2. Extract individual exam questions from each paper.
3. Preserve each question's wording as closely as possible.
4. Capture marks, question number, and paper filename when available.
5. Assign the best matching syllabus unit, topic, and subtopic.
6. If a question is unclear or doesn't match confidently, use
   "Needs Review" as the topic and explain why.
7. Do not invent missing marks, years, or question numbers.
8. Keep related subparts together unless they clearly test
   different topics.

Return valid JSON in this format:
{{
  "questions": [
    {{
      "paper": "filename",
      "question_number": "as printed or Unknown",
      "question": "question text",
      "marks": "marks or Unknown",
      "unit": "syllabus unit or Unknown",
      "topic": "syllabus topic or Needs Review",
      "subtopic": "syllabus subtopic or Unknown",
      "confidence": "High, Medium, or Low",
      "reason": "short explanation"
    }}
  ]
}}
"""
    return ask_gemini(client, prompt)


def create_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sorted Questions"

    headers = [
        "Paper", "Question No.", "Question", "Marks",
        "Unit", "Topic", "Subtopic", "Confidence", "Review Notes"
    ]
    ws.append(headers)

    for q in rows:
        ws.append([
            q.get("paper", ""),
            q.get("question_number", ""),
            q.get("question", ""),
            q.get("marks", ""),
            q.get("unit", ""),
            q.get("topic", ""),
            q.get("subtopic", ""),
            q.get("confidence", ""),
            q.get("reason", "")
        ])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for column, width in {
        "A": 25, "B": 16, "C": 65, "D": 12,
        "E": 24, "F": 30, "G": 30, "H": 14, "I": 40
    }.items():
        ws.column_dimensions[column].width = width

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


def create_pdf(rows):
    output = io.BytesIO()
    doc = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=38,
        leftMargin=38,
        topMargin=38,
        bottomMargin=38
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="QuestionText",
        parent=styles["BodyText"],
        fontSize=9,
        leading=12,
        spaceAfter=5
    ))

    story = [
        Paragraph("MedPaper Sorter — Topic-wise Questions", styles["Title"]),
        Spacer(1, 14)
    ]

    grouped = {}
    for q in rows:
        unit = str(q.get("unit", "Unknown"))
        topic = str(q.get("topic", "Needs Review"))
        grouped.setdefault((unit, topic), []).append(q)

    for (unit, topic), questions in sorted(grouped.items()):
        story.append(Paragraph(
            f"{escape(unit)} — {escape(topic)}",
            styles["Heading2"]
        ))

        for q in questions:
            details = (
                f"<b>{escape(str(q.get('question_number', 'Unknown')))}</b>"
                f" | Marks: {escape(str(q.get('marks', 'Unknown')))}"
                f" | Paper: {escape(str(q.get('paper', 'Unknown')))}"
            )
            body = escape(str(q.get("question", ""))).replace("\n", "<br/>")
            story.append(KeepTogether([
                Paragraph(details, styles["QuestionText"]),
                Paragraph(body, styles["QuestionText"]),
                Spacer(1, 5)
            ]))

        story.append(Spacer(1, 8))

    doc.build(story)
    return output.getvalue()


syllabus_file = st.file_uploader(
    "1. Upload your syllabus PDF",
    type=["pdf"],
    key="syllabus"
)

paper_files = st.file_uploader(
    "2. Upload one or more question-paper PDFs",
    type=["pdf"],
    accept_multiple_files=True,
    key="papers"
)

if syllabus_file:
    if st.button("Extract syllabus text"):
        with st.spinner("Reading syllabus..."):
            try:
                st.session_state["syllabus_text"] = extract_pdf_text(
                    syllabus_file
                )
                st.success("Syllabus text extracted.")
            except Exception as exc:
                st.error(f"Could not read syllabus: {exc}")

if "syllabus_text" in st.session_state:
    with st.expander("Review extracted syllabus text", expanded=False):
        st.text_area(
            "Syllabus text",
            st.session_state["syllabus_text"],
            height=250
        )

if paper_files:
    if st.button("Extract question-paper text"):
        extracted = {}
        with st.spinner("Reading question papers..."):
            try:
                for paper in paper_files:
                    extracted[paper.name] = extract_pdf_text(paper)
                st.session_state["paper_texts"] = extracted
                st.success(f"Extracted {len(extracted)} paper(s).")
            except Exception as exc:
                st.error(f"Could not read a paper: {exc}")

if "paper_texts" in st.session_state:
    with st.expander("Review extracted question-paper text", expanded=False):
        for filename, text in st.session_state["paper_texts"].items():
            st.markdown(f"**{filename}**")
            st.text_area(
                f"Extracted text — {filename}",
                text,
                height=180,
                key=f"text_{filename}"
            )

if st.button("✨ Sort questions by syllabus topics", type="primary"):
    syllabus_text = st.session_state.get("syllabus_text", "")
    paper_texts = st.session_state.get("paper_texts", {})

    if not syllabus_text:
        st.warning("Upload and extract the syllabus first.")
    elif not paper_texts:
        st.warning("Upload and extract at least one question paper.")
    else:
        try:
            client = get_client()
            with st.spinner(
                "Gemini is identifying topics and sorting questions..."
            ):
                result = classify_questions(
                    client, syllabus_text, paper_texts
                )
                rows = result.get("questions", [])

            if not rows:
                st.warning("No questions were returned. Try clearer PDFs.")
            else:
                st.session_state["sorted_questions"] = rows
                st.success(f"Processed {len(rows)} question(s).")
        except Exception as exc:
            st.error(
                "Sorting failed. Check your API key, model access, "
                "internet connection, and API quota. Details: " + str(exc)
            )

if "sorted_questions" in st.session_state:
    st.subheader("3. Review and correct the classifications")

    edited = st.data_editor(
        st.session_state["sorted_questions"],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "confidence": st.column_config.SelectboxColumn(
                "Confidence",
                options=["High", "Medium", "Low"]
            )
        }
    )
    st.session_state["sorted_questions"] = edited

    st.download_button(
        "⬇️ Download Excel",
        data=create_excel(edited),
        file_name="medpaper_sorted_questions.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        "⬇️ Download topic-wise PDF",
        data=create_pdf(edited),
        file_name="medpaper_sorted_questions.pdf",
        mime="application/pdf"
    )
