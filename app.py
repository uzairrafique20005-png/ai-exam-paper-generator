"""
ExamCraft AI – Exam Paper Generator
Powered by Groq API (llama-3.1-8b-instant)
Full redesign: light theme, animations, PDF export, proper Bloom's taxonomy formatting
"""

import os
import io
import re
import streamlit as st
from groq import Groq
import PyPDF2
from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch, cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, HRFlowable,
    Table, TableStyle, PageBreak
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY

# ─────────────────────────────────────────────
# COLOUR PALETTE (shared across app + exports)
# ─────────────────────────────────────────────
TEAL        = "#1A7A6E"
TEAL_LIGHT  = "#E8F5F3"
TEAL_MID    = "#2AA090"
AMBER       = "#E07B39"
AMBER_LIGHT = "#FDF3EC"
SLATE       = "#2D3748"
SLATE_LIGHT = "#718096"
WHITE       = "#FFFFFF"
OFF_WHITE   = "#F7FAFA"
BORDER      = "#D1E8E4"

# ReportLab colour objects
RL_TEAL        = colors.HexColor(TEAL)
RL_TEAL_LIGHT  = colors.HexColor(TEAL_LIGHT)
RL_AMBER       = colors.HexColor(AMBER)
RL_AMBER_LIGHT = colors.HexColor(AMBER_LIGHT)
RL_SLATE       = colors.HexColor(SLATE)
RL_SLATE_LIGHT = colors.HexColor(SLATE_LIGHT)
RL_WHITE       = colors.white
RL_BORDER      = colors.HexColor(BORDER)

# python-docx RGBColor objects
D_TEAL  = RGBColor(0x1A, 0x7A, 0x6E)
D_AMBER = RGBColor(0xE0, 0x7B, 0x39)
D_SLATE = RGBColor(0x2D, 0x37, 0x48)
D_GRAY  = RGBColor(0x71, 0x80, 0x96)

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="ExamCraft AI",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CUSTOM CSS – light theme, animations, always readable
# ─────────────────────────────────────────────
st.markdown(f"""
<style>
/* ── Force light base ── */
html, body, [data-testid="stAppViewContainer"], .stApp {{
    background-color: {OFF_WHITE} !important;
    color: {SLATE} !important;
}}

/* ── Sidebar: teal gradient ── */
section[data-testid="stSidebar"] {{
    background: linear-gradient(160deg, {TEAL} 0%, #145f55 100%) !important;
}}

/* White text for structural sidebar elements ONLY — never override inline-styled spans */
section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] h4,
section[data-testid="stSidebar"] > div p,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stSelectbox label,
section[data-testid="stSidebar"] .stNumberInput label,
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] > p,
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] > h2,
section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] > h3 {{
    color: {WHITE} !important;
}}

/* Radio button labels in sidebar */
section[data-testid="stSidebar"] .stRadio label p,
section[data-testid="stSidebar"] .stRadio label span:not([style]) {{
    color: {WHITE} !important;
}}

/* Caption / small text in sidebar */
section[data-testid="stSidebar"] .stCaption,
section[data-testid="stSidebar"] .stCaption p {{
    color: rgba(255,255,255,0.75) !important;
}}

/* Sidebar selectbox + number input */
section[data-testid="stSidebar"] [data-baseweb="select"] > div,
section[data-testid="stSidebar"] [data-baseweb="input"] > div {{
    background: rgba(255,255,255,0.15) !important;
    border-color: rgba(255,255,255,0.3) !important;
    color: {WHITE} !important;
    border-radius: 8px !important;
}}

/* ── Main content – always dark readable text ── */
.main .block-container p,
.main .block-container li,
.main .block-container label,
.stMarkdown p, .stMarkdown li {{
    color: {SLATE} !important;
}}
h1, h2, h3, h4 {{ color: {SLATE} !important; }}

/* ── Radio button labels – main area ── */
.stRadio label,
.stRadio label p,
.stRadio label span,
div[data-testid="stRadio"] label,
div[data-testid="stRadio"] label p {{
    color: {SLATE} !important;
    font-weight: 600 !important;
}}

/* ── Radio question text ── */
div[data-testid="stRadio"] > label > div > p {{
    color: {SLATE} !important;
}}

/* ── Metric labels ── */
[data-testid="stMetricLabel"]  {{ color: {SLATE_LIGHT} !important; }}
[data-testid="stMetricValue"]  {{ color: {TEAL} !important; }}

/* ── Animated hero banner ── */
.hero-banner {{
    background: linear-gradient(135deg, {TEAL} 0%, {TEAL_MID} 60%, #1d9c8e 100%);
    border-radius: 16px;
    padding: 2.5rem 2rem 2rem 2rem;
    margin-bottom: 1.5rem;
    animation: fadeSlideDown 0.7s ease forwards;
    box-shadow: 0 8px 32px rgba(26,122,110,0.18);
}}
.hero-title {{
    font-size: 2.4rem;
    font-weight: 800;
    color: {WHITE} !important;
    margin: 0 0 0.3rem 0;
    letter-spacing: -0.5px;
}}
.hero-sub {{
    color: rgba(255,255,255,0.88) !important;
    font-size: 1.05rem;
    margin: 0;
}}

/* ── Section cards ── */
.section-card {{
    background: {WHITE};
    border: 1px solid {BORDER};
    border-radius: 12px;
    padding: 1.5rem 1.6rem;
    margin-bottom: 1.2rem;
    box-shadow: 0 2px 12px rgba(26,122,110,0.06);
    animation: fadeIn 0.5s ease forwards;
}}
.section-label {{
    font-size: 0.75rem;
    font-weight: 700;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    color: {TEAL} !important;
    margin-bottom: 0.4rem;
}}
.section-title {{
    font-size: 1.15rem;
    font-weight: 700;
    color: {SLATE} !important;
    margin: 0 0 1rem 0;
}}

/* ── Exam output card ── */
.exam-output {{
    background: {WHITE};
    border-left: 5px solid {TEAL};
    border-radius: 10px;
    padding: 1.8rem 2rem;
    margin: 1rem 0;
    box-shadow: 0 4px 20px rgba(26,122,110,0.08);
    line-height: 1.8;
    color: {SLATE} !important;
}}
.exam-output * {{ color: {SLATE} !important; }}

/* ── Primary generate button ── */
.stButton > button {{
    background: linear-gradient(135deg, {TEAL} 0%, {TEAL_MID} 100%) !important;
    color: {WHITE} !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.65rem 1.8rem !important;
    font-size: 1rem !important;
    font-weight: 700 !important;
    box-shadow: 0 4px 14px rgba(26,122,110,0.35) !important;
    transition: all 0.25s ease !important;
}}
.stButton > button:hover {{
    background: linear-gradient(135deg, #145f55 0%, {TEAL} 100%) !important;
    box-shadow: 0 6px 20px rgba(26,122,110,0.45) !important;
    transform: translateY(-2px) !important;
}}
.stButton > button:active {{ transform: translateY(0) !important; }}

/* ── Download buttons ── */
.stDownloadButton > button {{
    border-radius: 10px !important;
    font-weight: 700 !important;
    transition: all 0.22s ease !important;
    border: 2px solid {TEAL} !important;
    color: {TEAL} !important;
    background: {WHITE} !important;
}}
.stDownloadButton > button:hover {{
    background: {TEAL} !important;
    color: {WHITE} !important;
    transform: translateY(-2px) !important;
    box-shadow: 0 5px 15px rgba(26,122,110,0.3) !important;
}}

/* ── Radio ── */
.stRadio label {{ color: {SLATE} !important; font-weight: 600 !important; }}
.stRadio label p, .stRadio label span {{ color: {SLATE} !important; }}

/* ── Text areas ── */
.stTextArea textarea {{
    border-radius: 8px !important;
    border-color: {BORDER} !important;
    background: {WHITE} !important;
    color: {SLATE} !important;
}}
.stTextArea textarea:focus {{
    border-color: {TEAL} !important;
    box-shadow: 0 0 0 2px {TEAL_LIGHT} !important;
}}

/* ── Expander ── */
details summary {{
    background: {TEAL_LIGHT} !important;
    border-radius: 8px !important;
    padding: 0.8rem 1rem !important;
    font-weight: 700 !important;
    color: {TEAL} !important;
    cursor: pointer !important;
}}

/* ── Alerts ── */
.stAlert {{ border-radius: 10px !important; }}

/* ── Animations ── */
@keyframes fadeSlideDown {{
    from {{ opacity: 0; transform: translateY(-20px); }}
    to   {{ opacity: 1; transform: translateY(0); }}
}}
@keyframes fadeIn {{
    from {{ opacity: 0; transform: translateY(10px); }}
    to   {{ opacity: 1; transform: translateY(0); }}
}}

/* ── Hide Streamlit chrome ── */
#MainMenu, footer {{ visibility: hidden; }}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# API KEY LOADER
# ─────────────────────────────────────────────
def load_api_key() -> str:
    try:
        return st.secrets["GROQ_API_KEY"]
    except (KeyError, FileNotFoundError):
        key = os.environ.get("GROQ_API_KEY", "")
        if not key:
            st.error("⚠️ Groq API key not found. Add to `.streamlit/secrets.toml` or set GROQ_API_KEY env var.")
            st.stop()
        return key


# ─────────────────────────────────────────────
# FILE EXTRACTOR
# ─────────────────────────────────────────────
MAX_CONTEXT_CHARS = 3000

def extract_file_content(uploaded_file) -> str:
    try:
        if uploaded_file.type == "application/pdf":
            reader = PyPDF2.PdfReader(uploaded_file)
            pages = [page.extract_text() or "" for page in reader.pages]
            raw_text = "\n".join(pages).strip()
        else:
            raw_text = uploaded_file.read().decode("utf-8", errors="ignore").strip()
    except Exception as e:
        st.warning(f"File reading error: {e}.")
        return ""
    if len(raw_text) > MAX_CONTEXT_CHARS:
        st.warning("📄 Document is large. Only the first ~500 words were used to stay within API limits.")
        return raw_text[:MAX_CONTEXT_CHARS]
    return raw_text


# ─────────────────────────────────────────────
# PROMPT – enforces Bloom's tags + clean delimiters
# ─────────────────────────────────────────────
def build_prompt(topic_context, grade, difficulty, duration, total_marks, source_type):
    context_instruction = (
        f"Base the entire exam STRICTLY on this uploaded source material:\n\n{topic_context}"
        if source_type == "Upload Source Material"
        else f"Cover the following subject/topics:\n\n{topic_context}"
    )
    return f"""You are an expert curriculum designer. Produce a complete exam paper and answer key.

CONFIGURATION:
- Grade/Semester: {grade}
- Difficulty: {difficulty}
- Duration: {duration} minutes
- Total Marks: {total_marks}
- {context_instruction}

OUTPUT FORMAT — follow this structure EXACTLY:

---EXAM START---

## EXAM PAPER

### SECTION A — MULTIPLE CHOICE QUESTIONS (5 Questions)

Q1. [BLOOM: Knowledge] (2 Marks)
Question text here?
   A) Option one
   B) Option two
   C) Option three
   D) Option four

Q2. [BLOOM: Comprehension] (2 Marks)
...

Q3. [BLOOM: Application] (2 Marks)
...

Q4. [BLOOM: Analysis] (2 Marks)
...

Q5. [BLOOM: Evaluation] (2 Marks)
...

### SECTION B — SHORT ANSWER QUESTIONS (3 Questions)

Q6. [BLOOM: Application] (5 Marks)
Question text here?

Q7. [BLOOM: Comprehension] (5 Marks)
Question text here?

Q8. [BLOOM: Analysis] (10 Marks)
Question text here?

### SECTION C — ESSAY QUESTION (1 Question)

Q9. [BLOOM: Synthesis] (20 Marks)
Essay question text here?

---ANSWER KEY---

## ANSWER KEY & GRADING RUBRIC

### MCQ Answers
Q1. Correct: B — Explanation.
Q2. Correct: A — Explanation.
Q3. Correct: C — Explanation.
Q4. Correct: D — Explanation.
Q5. Correct: B — Explanation.

### Short Answer — Model Answers
Q6. Model answer in 3-5 sentences.
Q7. Model answer in 3-5 sentences.
Q8. Model answer in 3-5 sentences.

### Essay — Grading Rubric
| Criterion | Excellent (5) | Good (3-4) | Needs Work (1-2) |
|---|---|---|---|
| Argument / Thesis | Clear original thesis | Adequate thesis | Weak or missing |
| Evidence & Examples | 3+ relevant examples | 1-2 examples | Vague or missing |
| Structure & Clarity | Logical flow | Mostly clear | Disorganised |
| Critical Thinking | Deep insight | Some analysis | Surface level |

RULES:
- EVERY question MUST have [BLOOM: Level] tag exactly as shown.
- Use ---EXAM START--- and ---ANSWER KEY--- as the only section separators.
- Distribute {total_marks} marks proportionally.
- No extra commentary — output only the exam and answer key.
"""


# ─────────────────────────────────────────────
# GROQ API
# ─────────────────────────────────────────────
def generate_exam(prompt: str, api_key: str) -> str:
    client = Groq(api_key=api_key)
    response = client.chat.completions.create(
        model="llama-3.1-8b-instant",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7,
        max_tokens=4096,
    )
    return response.choices[0].message.content


# ─────────────────────────────────────────────
# SPLIT TEXT
# ─────────────────────────────────────────────
BLOOM_RE = re.compile(r'\[BLOOM:\s*([A-Za-z]+)\]')

def split_exam_and_key(full_text: str):
    marker = "---ANSWER KEY---"
    if marker in full_text:
        idx = full_text.find(marker)
        return full_text[:idx].strip(), full_text[idx:].strip()
    for kw in ["ANSWER KEY", "═══"]:
        if kw in full_text.upper():
            idx = full_text.upper().find(kw)
            ls  = full_text.rfind("\n", 0, idx)
            return full_text[:ls].strip(), full_text[ls:].strip()
    return full_text.strip(), ""


# ─────────────────────────────────────────────
# HTML RENDERER for on-screen display
# ─────────────────────────────────────────────
BLOOM_COLOURS = {
    "knowledge":     ("#EEF2FF", "#3730A3"),
    "comprehension": ("#F0FDF4", "#166534"),
    "application":   ("#FFF7ED", "#C2410C"),
    "analysis":      ("#FDF4FF", "#7E22CE"),
    "synthesis":     ("#FFF1F2", "#BE123C"),
    "evaluation":    ("#ECFEFF", "#0E7490"),
}

def bloom_badge_html(level: str) -> str:
    key = level.lower()
    bg, fg = BLOOM_COLOURS.get(key, ("#F3F4F6", "#374151"))
    return (
        f'<span style="display:inline-block;padding:2px 10px;border-radius:20px;'
        f'font-size:0.72rem;font-weight:700;background:{bg};color:{fg};'
        f'margin-left:6px;vertical-align:middle;">🧠 {level}</span>'
    )

def render_exam_html(text: str, is_key=False) -> str:
    accent = AMBER if is_key else TEAL
    lines  = text.split("\n")
    out    = []
    for line in lines:
        s = line.strip()
        if not s or s in ("---EXAM START---", "---ANSWER KEY---"):
            out.append("<br>")
            continue

        def replace_bloom(m):
            return bloom_badge_html(m.group(1))
        s_rendered = BLOOM_RE.sub(replace_bloom, s)

        if s.startswith("## "):
            t = s[3:].strip()
            out.append(
                f'<h2 style="color:{accent};border-bottom:2px solid {TEAL_LIGHT};'
                f'padding-bottom:6px;margin-top:1.4rem;">{BLOOM_RE.sub(replace_bloom, t)}</h2>'
            )
        elif s.startswith("### "):
            t = s[4:].strip()
            out.append(
                f'<h3 style="color:{SLATE};background:{TEAL_LIGHT};padding:8px 14px;'
                f'border-radius:6px;border-left:4px solid {accent};margin-top:1.2rem;">'
                f'{BLOOM_RE.sub(replace_bloom, t)}</h3>'
            )
        elif s.startswith("#### "):
            t = s[5:].strip()
            out.append(f'<h4 style="color:{AMBER};margin-top:1rem;">{BLOOM_RE.sub(replace_bloom, t)}</h4>')
        elif re.match(r'^[A-D]\)', s):
            out.append(
                f'<div style="padding:3px 0 3px 2rem;color:{SLATE_LIGHT};font-size:0.95rem;">{s}</div>'
            )
        elif re.match(r'^Q\d+\.', s):
            out.append(
                f'<p style="font-weight:700;color:{SLATE};margin:1.1rem 0 0.3rem 0;'
                f'font-size:1rem;">{s_rendered}</p>'
            )
        elif s.startswith("|"):
            cells = [c.strip() for c in s.strip("|").split("|")]
            if all(set(c) <= set("-: ") for c in cells):
                continue
            row_html = "".join(
                f'<td style="padding:6px 10px;border:1px solid {BORDER};color:{SLATE};">{c}</td>'
                for c in cells
            )
            bg = TEAL_LIGHT if not is_key else AMBER_LIGHT
            out.append(f'<table style="width:100%;border-collapse:collapse;margin:6px 0;">'
                       f'<tr style="background:{bg};">{row_html}</tr></table>')
        elif s.startswith("---"):
            out.append(f'<hr style="border:1px solid {BORDER};margin:1rem 0;">')
        else:
            out.append(f'<p style="color:{SLATE};margin:0.3rem 0;">{s_rendered}</p>')
    return "\n".join(out)


# ─────────────────────────────────────────────
# PDF EXPORT
# ─────────────────────────────────────────────
def _rl_styles():
    S = {}
    S["title"]    = ParagraphStyle("T",  fontName="Helvetica-Bold",   fontSize=22, textColor=RL_TEAL,       alignment=TA_CENTER, spaceAfter=6)
    S["subtitle"] = ParagraphStyle("ST", fontName="Helvetica",        fontSize=10, textColor=RL_SLATE_LIGHT, alignment=TA_CENTER, spaceAfter=18)
    S["h2"]       = ParagraphStyle("H2", fontName="Helvetica-Bold",   fontSize=14, textColor=RL_TEAL,       spaceBefore=14, spaceAfter=5)
    S["h2key"]    = ParagraphStyle("H2K",fontName="Helvetica-Bold",   fontSize=14, textColor=RL_AMBER,      spaceBefore=14, spaceAfter=5)
    S["h3"]       = ParagraphStyle("H3", fontName="Helvetica-Bold",   fontSize=11, textColor=RL_SLATE,      spaceBefore=10, spaceAfter=4, backColor=RL_TEAL_LIGHT, leftIndent=8, borderPad=6)
    S["h4"]       = ParagraphStyle("H4", fontName="Helvetica-Bold",   fontSize=10, textColor=RL_AMBER,      spaceBefore=8, spaceAfter=3)
    S["bloom"]    = ParagraphStyle("BL", fontName="Helvetica-BoldOblique", fontSize=8, textColor=RL_TEAL, spaceAfter=1)
    S["question"] = ParagraphStyle("Q",  fontName="Helvetica-Bold",   fontSize=10, textColor=RL_SLATE,      spaceBefore=8, spaceAfter=3, leading=14)
    S["option"]   = ParagraphStyle("OP", fontName="Helvetica",        fontSize=10, textColor=RL_SLATE_LIGHT, leftIndent=22, spaceAfter=2, leading=13)
    S["body"]     = ParagraphStyle("BD", fontName="Helvetica",        fontSize=10, textColor=RL_SLATE,      spaceAfter=4, leading=15, alignment=TA_JUSTIFY)
    S["keybody"]  = ParagraphStyle("KB", fontName="Helvetica",        fontSize=10, textColor=RL_SLATE,      spaceAfter=4, leading=14, backColor=RL_AMBER_LIGHT, leftIndent=8, borderPad=4)
    return S

def build_pdf(full_text, title, grade, difficulty, duration, total_marks) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=2*cm, rightMargin=2*cm,
                            topMargin=2.2*cm, bottomMargin=2*cm)
    S     = _rl_styles()
    story = []

    # Header
    story.append(Paragraph("ExamCraft AI", S["title"]))
    story.append(Paragraph(title, S["subtitle"]))

    # Meta table
    meta = [["Level", grade.split("(")[0].strip(), "Difficulty", difficulty],
            ["Duration", f"{duration} mins", "Total Marks", str(total_marks)]]
    mt = Table(meta, colWidths=[2.5*cm, 6.5*cm, 2.5*cm, 3.5*cm])
    mt.setStyle(TableStyle([
        ("BACKGROUND",   (0,0), (0,-1), RL_TEAL),
        ("BACKGROUND",   (2,0), (2,-1), RL_TEAL),
        ("TEXTCOLOR",    (0,0), (0,-1), RL_WHITE),
        ("TEXTCOLOR",    (2,0), (2,-1), RL_WHITE),
        ("BACKGROUND",   (1,0), (1,-1), RL_TEAL_LIGHT),
        ("BACKGROUND",   (3,0), (3,-1), RL_TEAL_LIGHT),
        ("TEXTCOLOR",    (1,0), (3,-1), RL_SLATE),
        ("FONTNAME",     (0,0), (-1,-1), "Helvetica"),
        ("FONTNAME",     (0,0), (0,-1), "Helvetica-Bold"),
        ("FONTNAME",     (2,0), (2,-1), "Helvetica-Bold"),
        ("FONTSIZE",     (0,0), (-1,-1), 9),
        ("ALIGN",        (0,0), (-1,-1), "CENTER"),
        ("VALIGN",       (0,0), (-1,-1), "MIDDLE"),
        ("BOX",          (0,0), (-1,-1), 0.5, RL_BORDER),
        ("INNERGRID",    (0,0), (-1,-1), 0.3, RL_BORDER),
        ("TOPPADDING",   (0,0), (-1,-1), 6),
        ("BOTTOMPADDING",(0,0), (-1,-1), 6),
    ]))
    story.append(mt)
    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", thickness=2, color=RL_TEAL))
    story.append(Spacer(1, 8))

    exam_body, answer_key = split_exam_and_key(full_text)

    def render(text, is_key=False):
        for line in text.split("\n"):
            s = line.strip()
            if not s or s in ("---EXAM START---", "---ANSWER KEY---"):
                story.append(Spacer(1, 4))
                continue

            bloom_m = BLOOM_RE.search(s)
            bloom_l = bloom_m.group(1) if bloom_m else None
            s_clean = BLOOM_RE.sub("", s).strip()

            if s.startswith("## "):
                style = S["h2key"] if is_key else S["h2"]
                story.append(Paragraph(s[3:].strip(), style))
                story.append(HRFlowable(width="100%", thickness=1,
                                        color=RL_AMBER if is_key else RL_TEAL))
            elif s.startswith("### "):
                story.append(Paragraph(s[4:].strip(), S["h3"]))
            elif s.startswith("#### "):
                story.append(Paragraph(s[5:].strip(), S["h4"]))
            elif re.match(r'^[A-D]\)', s_clean):
                story.append(Paragraph(s_clean, S["option"]))
            elif re.match(r'^Q\d+\.', s_clean):
                if bloom_l:
                    story.append(Paragraph(f"Bloom's Taxonomy: {bloom_l}", S["bloom"]))
                story.append(Paragraph(s_clean, S["question"]))
            elif s.startswith("|"):
                cells = [c.strip() for c in s.strip("|").split("|")]
                if all(set(c) <= set("-: ") for c in cells):
                    continue
                tbl = Table([cells])
                bg  = RL_AMBER_LIGHT if is_key else RL_TEAL_LIGHT
                tbl.setStyle(TableStyle([
                    ("BACKGROUND",   (0,0), (-1,-1), bg),
                    ("TEXTCOLOR",    (0,0), (-1,-1), RL_SLATE),
                    ("FONTNAME",     (0,0), (-1,-1), "Helvetica"),
                    ("FONTSIZE",     (0,0), (-1,-1), 8),
                    ("BOX",          (0,0), (-1,-1), 0.4, RL_BORDER),
                    ("INNERGRID",    (0,0), (-1,-1), 0.2, RL_BORDER),
                    ("TOPPADDING",   (0,0), (-1,-1), 4),
                    ("BOTTOMPADDING",(0,0), (-1,-1), 4),
                    ("ALIGN",        (0,0), (-1,-1), "LEFT"),
                ]))
                story.append(tbl)
                story.append(Spacer(1, 2))
            elif s.startswith("---"):
                story.append(Spacer(1, 4))
            else:
                style = S["keybody"] if is_key else S["body"]
                story.append(Paragraph(s_clean or s, style))

    render(exam_body, is_key=False)
    if answer_key:
        story.append(PageBreak())
        story.append(Paragraph("ANSWER KEY & GRADING RUBRIC", S["h2key"]))
        story.append(HRFlowable(width="100%", thickness=2, color=RL_AMBER))
        story.append(Spacer(1, 8))
        render(answer_key, is_key=True)

    doc.build(story)
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────
# DOCX EXPORT
# ─────────────────────────────────────────────
def build_docx(full_text, title) -> bytes:
    doc = DocxDocument()
    for sec in doc.sections:
        sec.top_margin    = Inches(1)
        sec.bottom_margin = Inches(1)
        sec.left_margin   = Inches(1.2)
        sec.right_margin  = Inches(1.2)

    tp = doc.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r  = tp.add_run("ExamCraft AI — " + title)
    r.bold = True; r.font.size = Pt(18); r.font.color.rgb = D_TEAL
    doc.add_paragraph()

    exam_body, answer_key = split_exam_and_key(full_text)

    def add_sec(text, is_key=False):
        accent = D_AMBER if is_key else D_TEAL
        for line in text.split("\n"):
            s = line.strip()
            if not s or s in ("---EXAM START---", "---ANSWER KEY---"):
                doc.add_paragraph(); continue

            bloom_m = BLOOM_RE.search(s)
            bloom_l = bloom_m.group(1) if bloom_m else None
            s_clean = BLOOM_RE.sub("", s).strip()

            if s.startswith("## "):
                h = doc.add_heading(s[3:].strip(), level=1)
                for r in h.runs: r.font.color.rgb = accent
            elif s.startswith("### "):
                h = doc.add_heading(s[4:].strip(), level=2)
                for r in h.runs: r.font.color.rgb = accent
            elif s.startswith("#### "):
                p = doc.add_paragraph()
                r = p.add_run(s[5:].strip())
                r.bold = True; r.font.size = Pt(11); r.font.color.rgb = D_AMBER
            elif re.match(r'^[A-D]\)', s_clean):
                p = doc.add_paragraph(style="List Bullet")
                r = p.add_run(s_clean); r.font.color.rgb = D_GRAY
            elif re.match(r'^Q\d+\.', s_clean):
                if bloom_l:
                    bp = doc.add_paragraph()
                    br = bp.add_run(f"  Bloom's Taxonomy: {bloom_l}")
                    br.italic = True; br.font.size = Pt(8); br.font.color.rgb = D_TEAL
                p = doc.add_paragraph()
                r = p.add_run(s_clean)
                r.bold = True; r.font.size = Pt(11); r.font.color.rgb = D_SLATE
            elif s.startswith("|"):
                cells = [c.strip() for c in s.strip("|").split("|")]
                if all(set(c) <= set("-: ") for c in cells): continue
                p = doc.add_paragraph("  |  ".join(cells))
                p.paragraph_format.left_indent = Inches(0.15)
            elif s.startswith("---"):
                doc.add_paragraph()
            else:
                p = doc.add_paragraph(s_clean or s)
                if is_key: p.paragraph_format.left_indent = Inches(0.15)

    add_sec(exam_body, is_key=False)
    if answer_key:
        doc.add_page_break()
        h = doc.add_heading("ANSWER KEY & GRADING RUBRIC", level=1)
        for r in h.runs: r.font.color.rgb = D_AMBER
        add_sec(answer_key, is_key=True)

    buf = io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.read()


# ═══════════════════════════════════════════════
# ══  UI  ═══════════════════════════════════════
# ═══════════════════════════════════════════════

# ── Sidebar ──
with st.sidebar:
    st.markdown("## ⚙️ Exam Configuration")
    st.markdown("---")
    grade = st.selectbox("🎓 Target Grade / Semester", [
        "Grade 6–8 (Middle School)", "Grade 9–10 (Secondary)",
        "Grade 11–12 (Pre-University)", "Semester 1 (University)",
        "Semester 2 (University)", "Semester 3–4 (University)", "Postgraduate Level",
    ])
    difficulty = st.selectbox("📊 Exam Difficulty", ["Beginner", "Intermediate", "Advanced", "Expert"], index=1)
    c1, c2 = st.columns(2)
    with c1: duration    = st.number_input("⏱️ Duration", min_value=15, max_value=300, value=60,  step=15)
    with c2: total_marks = st.number_input("🏆 Marks",    min_value=10, max_value=200, value=100, step=5)

    st.markdown("---")
    st.markdown("### 🧠 Bloom's Taxonomy")
    # Render all badges inside a white-backed container so they show on the teal sidebar
    badges_html = "".join(
        f'<span style="background:{bg};color:{fg};padding:3px 12px;border-radius:20px;'
        f'font-size:0.74rem;font-weight:700;display:inline-block;margin:3px 3px;">'
        f'{lvl.capitalize()}</span>'
        for lvl, (bg, fg) in BLOOM_COLOURS.items()
    )
    st.markdown(
        f'<div style="background:rgba(255,255,255,0.15);border-radius:10px;'
        f'padding:10px 10px 6px 10px;line-height:2.2;">{badges_html}</div>',
        unsafe_allow_html=True,
    )
    st.markdown("---")
    st.caption("ExamCraft AI · Groq LLaMA 3.1")


# ── Hero ──
st.markdown(f"""
<div class="hero-banner">
  <div class="hero-title">📝 ExamCraft AI</div>
  <div class="hero-sub">AI-powered exam generator for teachers — structured, Bloom's-tagged, ready to export in PDF, DOCX & TXT.</div>
</div>
""", unsafe_allow_html=True)


# ── Step 1 ──
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Step 1</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">📂 Provide Exam Content</div>', unsafe_allow_html=True)

source_type   = st.radio("How would you like to provide the topic?",
                         ["Enter Topic Manually", "Upload Source Material"], horizontal=True)
topic_context = ""

if source_type == "Enter Topic Manually":
    topic_context = st.text_area(
        "Subject & Topics",
        placeholder="e.g., Subject: Physics\nTopics: Thermodynamics, Laws of Entropy, Heat Transfer",
        height=140,
    )
else:
    uploaded_file = st.file_uploader("Upload syllabus, notes, or chapter (.pdf or .txt)", type=["pdf","txt"])
    if uploaded_file:
        with st.spinner("Extracting content…"):
            topic_context = extract_file_content(uploaded_file)
        if topic_context:
            st.success(f"✅ Extracted **{len(topic_context.split())} words** from `{uploaded_file.name}`")
            with st.expander("Preview extracted content"):
                st.text(topic_context[:1500] + ("…" if len(topic_context) > 1500 else ""))
        else:
            st.warning("Could not extract text from this file.")
st.markdown('</div>', unsafe_allow_html=True)


# ── Step 2 ──
st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-label">Step 2</div>', unsafe_allow_html=True)
st.markdown('<div class="section-title">🚀 Generate Your Exam</div>', unsafe_allow_html=True)
gen_col, _ = st.columns([1, 3])
with gen_col:
    generate_btn = st.button("⚡ Generate Exam Paper", use_container_width=True)
st.markdown('</div>', unsafe_allow_html=True)


# ── Generation ──
if generate_btn:
    if not topic_context.strip():
        st.warning("⚠️ Please enter a topic or upload a source file first.")
        st.stop()
    api_key = load_api_key()
    prompt  = build_prompt(topic_context, grade, difficulty, duration, total_marks, source_type)
    with st.spinner("🧠 Crafting your exam paper — this may take 20–40 seconds…"):
        try:
            full_output = generate_exam(prompt, api_key)
            st.session_state.update({
                "full_output": full_output, "exam_grade": grade,
                "exam_diff": difficulty, "exam_dur": duration, "exam_marks": total_marks,
            })
        except Exception as e:
            st.error(f"❌ Generation failed: {e}")
            st.stop()


# ── Results ──
if "full_output" in st.session_state:
    full_output = st.session_state["full_output"]
    grade_label = st.session_state.get("exam_grade",  grade)
    diff_label  = st.session_state.get("exam_diff",   difficulty)
    dur_val     = st.session_state.get("exam_dur",    duration)
    marks_val   = st.session_state.get("exam_marks",  total_marks)

    exam_body, answer_key = split_exam_and_key(full_output)

    st.success("✅ Exam generated successfully!")

    mc1, mc2, mc3, mc4 = st.columns(4)
    mc1.metric("Level",       grade_label.split("(")[0].strip())
    mc2.metric("Difficulty",  diff_label)
    mc3.metric("Duration",    f"{dur_val} mins")
    mc4.metric("Total Marks", str(marks_val))

    st.markdown("---")
    st.markdown("## 📄 Generated Exam Paper")
    st.markdown(
        f'<div class="exam-output">{render_exam_html(exam_body, is_key=False)}</div>',
        unsafe_allow_html=True,
    )

    if answer_key:
        st.markdown("---")
        with st.expander("🔑 View Answer Key & Grading Rubric (Teacher Only)", expanded=False):
            st.info("📌 Keep this section hidden when distributing to students.")
            st.markdown(
                f'<div class="exam-output" style="border-left-color:{AMBER};">'
                f'{render_exam_html(answer_key, is_key=True)}</div>',
                unsafe_allow_html=True,
            )

    # ── Step 3: Export ──
    st.markdown("---")
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-label">Step 3</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-title">💾 Export Your Exam</div>', unsafe_allow_html=True)

    exam_title  = f"Exam — {grade_label} — {diff_label}"
    txt_content = f"{exam_title}\n{'='*60}\n\n{full_output}"

    with st.spinner("Preparing download files…"):
        docx_bytes = build_docx(full_output, exam_title)
        pdf_bytes  = build_pdf(full_output, exam_title, grade_label, diff_label, dur_val, marks_val)

    dl1, dl2, dl3 = st.columns(3)
    with dl1:
        st.download_button("📄 Download .txt", data=txt_content.encode("utf-8"),
                           file_name="exam_paper.txt", mime="text/plain",
                           use_container_width=True)
    with dl2:
        st.download_button("📝 Download .docx", data=docx_bytes,
                           file_name="exam_paper.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True)
    with dl3:
        st.download_button("🖨️ Download .pdf", data=pdf_bytes,
                           file_name="exam_paper.pdf", mime="application/pdf",
                           use_container_width=True)

    st.caption("All formats include the exam + answer key. PDF and DOCX use teal/amber theme with Bloom's labels.")
    st.markdown('</div>', unsafe_allow_html=True)
