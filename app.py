import streamlit as st
import google.generativeai as genai
import fitz  # PyMuPDF
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import json
import os
import io
import base64
import re
from dotenv import load_dotenv

# ── Load environment variables (works locally) ──────────────────────────────
load_dotenv()

# ── Read config: Streamlit Secrets first, then .env fallback ────────────────
def get_secret(key, default=""):
    try:
        return st.secrets[key]
    except Exception:
        return os.getenv(key, default)

GEMINI_API_KEY  = get_secret("GEMINI_API_KEY")
GOOGLE_SHEET_ID = get_secret("GOOGLE_SHEET_ID")

SHEET_HEADERS = ["Date", "Name", "Topic", "Assessment", "Marks", "Percentage", "Feedback/Remarks"]

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Assessment Grader",
    page_icon="📝",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Custom CSS ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d6a9f 100%);
        padding: 2rem;
        border-radius: 12px;
        text-align: center;
        color: white;
        margin-bottom: 2rem;
    }
    .main-header h1 { font-size: 2.2rem; margin: 0; }
    .main-header p  { font-size: 1rem; margin: 0.5rem 0 0; opacity: 0.85; }
    .success-card {
        background: #f0fff4;
        border-left: 4px solid #28a745;
        border-radius: 8px;
        padding: 1rem 1.5rem;
        margin: 0.8rem 0;
    }
    .warning-card {
        background: #fffbf0;
        border-left: 4px solid #ffc107;
        border-radius: 8px;
        padding: 1rem 1.5rem;
        margin: 0.8rem 0;
    }
    .error-card {
        background: #fff5f5;
        border-left: 4px solid #dc3545;
        border-radius: 8px;
        padding: 1rem 1.5rem;
        margin: 0.8rem 0;
    }
    .stButton > button {
        width: 100%;
        background: linear-gradient(135deg, #1e3a5f, #2d6a9f);
        color: white;
        border: none;
        padding: 0.6rem 1.5rem;
        border-radius: 8px;
        font-size: 1rem;
        font-weight: 600;
    }
    .stButton > button:hover { opacity: 0.9; }
</style>
""", unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════
#  HELPERS
# ════════════════════════════════════════════════════════════════════

def pdf_to_images(pdf_bytes: bytes) -> list:
    """Convert every PDF page to PNG bytes for Gemini."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for page in doc:
        mat = fitz.Matrix(2.0, 2.0)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        images.append(pix.tobytes("png"))
    doc.close()
    return images


def extract_and_grade_with_gemini(image_bytes_list: list, api_key: str) -> dict:
    """Send all PDF pages to Gemini Vision. Gemini extracts metadata AND grades answers."""
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")

    prompt = """You are an expert technical assessor for an IT training organization.

You will receive images of a scanned intern assessment answer sheet (could be handwritten or printed).

Your tasks:
1. Extract the following metadata from the document:
   - date (written on the paper — use exactly as written, format as DD-MM-YYYY if possible)
   - intern_name (full name of the intern)
   - topic (e.g., "MSSQL", "Python", "SQL", etc.)
   - assessment_number (e.g., "Assessment 1", "Assessment 2", etc.)

2. Extract every question and the intern's answer.

3. For each question:
   - Determine the correct answer using your knowledge
   - Compare with the intern's answer
   - Award marks: 1 mark per question (partial credit: 0.5 if partially correct)
   - Write brief feedback

4. Calculate:
   - total_marks: sum of marks awarded
   - max_marks: total number of questions
   - percentage: (total_marks / max_marks) * 100

5. Write overall_feedback summarizing performance (2-3 sentences).

Respond ONLY with a valid JSON object in this exact format (no markdown, no extra text):
{
  "date": "DD-MM-YYYY",
  "intern_name": "Full Name",
  "topic": "Topic Name",
  "assessment_number": "Assessment N",
  "questions": [
    {
      "q_no": 1,
      "question": "Question text",
      "intern_answer": "What intern wrote",
      "correct_answer": "Correct answer",
      "marks_awarded": 1,
      "max_marks": 1,
      "feedback": "Brief comment"
    }
  ],
  "total_marks": 0,
  "max_marks": 0,
  "percentage": 0.0,
  "overall_feedback": "Overall performance summary."
}

If you cannot read a field, use "Unknown" for strings and 0 for numbers.
Return ONLY the JSON. No markdown. No explanation. No code fences."""

    # Build parts: images + text prompt
    parts = []
    for img_bytes in image_bytes_list:
        parts.append({
            "inline_data": {
                "mime_type": "image/png",
                "data": base64.b64encode(img_bytes).decode("utf-8")
            }
        })
    parts.append({"text": prompt})

    response = model.generate_content({"parts": parts})

    raw = response.text.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw)
    raw = re.sub(r"\s*```$", "", raw)
    return json.loads(raw)


# ── Google Sheets helpers ────────────────────────────────────────────────────

def get_gsheet_from_secret(sheet_id: str):
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    try:
        creds_dict = dict(st.secrets["GOOGLE_CREDENTIALS"])
        if "private_key" in creds_dict:
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    except Exception:
        creds_path = st.session_state.get("creds_path", "")
        if not creds_path or not os.path.exists(creds_path):
            raise ValueError("Google credentials not found. Please upload credentials.json in the sidebar.")
        creds = Credentials.from_service_account_file(creds_path, scopes=scopes)

    gc = gspread.authorize(creds)
    return gc.open_by_key(sheet_id).sheet1


def ensure_headers(ws):
    existing = ws.row_values(1)
    if existing != SHEET_HEADERS:
        ws.clear()
        ws.append_row(SHEET_HEADERS)
        ws.format("A1:G1", {
            "textFormat": {"bold": True},
            "backgroundColor": {"red": 0.18, "green": 0.23, "blue": 0.37}
        })


def is_duplicate(ws, name: str, assessment: str) -> bool:
    records = ws.get_all_records()
    for row in records:
        if (str(row.get("Name", "")).strip().lower() == name.strip().lower() and
                str(row.get("Assessment", "")).strip().lower() == assessment.strip().lower()):
            return True
    return False


def append_to_sheet(ws, row_data: dict):
    ws.append_row([
        row_data["date"],
        row_data["intern_name"],
        row_data["topic"],
        row_data["assessment_number"],
        f"{row_data['total_marks']}/{row_data['max_marks']}",
        f"{row_data['percentage']:.1f}%",
        row_data["overall_feedback"],
    ])


# ── Excel download helper ────────────────────────────────────────────────────

def build_excel(result: dict) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Assessment Result"

    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="1E3A5F")
    center      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin        = Side(border_style="thin", color="AAAAAA")
    border      = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col, h in enumerate(SHEET_HEADERS, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font; cell.fill = header_fill
        cell.alignment = center; cell.border = border

    data_row = [
        result["date"], result["intern_name"], result["topic"],
        result["assessment_number"],
        f"{result['total_marks']}/{result['max_marks']}",
        f"{result['percentage']:.1f}%", result["overall_feedback"],
    ]
    for col, val in enumerate(data_row, 1):
        cell = ws.cell(row=2, column=col, value=val)
        cell.alignment = center; cell.border = border

    ws.cell(row=4, column=1, value="Question-wise Breakdown").font = Font(bold=True, size=12)
    q_headers = ["Q.No", "Question", "Intern's Answer", "Correct Answer", "Marks", "Feedback"]
    for col, h in enumerate(q_headers, 1):
        cell = ws.cell(row=5, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2D6A9F")
        cell.alignment = center; cell.border = border

    for i, q in enumerate(result.get("questions", []), 6):
        awarded = q.get("marks_awarded", 0)
        maximum = q.get("max_marks", 1)
        bg = "F0FFF4" if awarded >= maximum else "FFF5F5" if awarded == 0 else "FFFBF0"
        for col, val in enumerate([q.get("q_no", i-5), q.get("question",""),
                                    q.get("intern_answer",""), q.get("correct_answer",""),
                                    f"{awarded}/{maximum}", q.get("feedback","")], 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border
            cell.fill = PatternFill("solid", fgColor=bg)

    for col, w in enumerate([8, 40, 35, 35, 10, 40], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ════════════════════════════════════════════════════════════════════

with st.sidebar:
    st.image("https://img.icons8.com/color/96/graduation-cap.png", width=80)
    st.title("⚙️ Configuration")

    st.subheader("🔑 Gemini API Key")
    api_key_input = st.text_input(
        "Google Gemini API Key",
        value=GEMINI_API_KEY,
        type="password",
        help="Get your key at: aistudio.google.com/app/apikey",
    )

    st.subheader("📊 Google Sheets")
    sheet_id_input = st.text_input(
        "Google Sheet ID",
        value=GOOGLE_SHEET_ID,
        help="Found in the Sheet URL: /spreadsheets/d/<SHEET_ID>/",
    )

    st.info("💡 On Streamlit Cloud, credentials are read from **Secrets** automatically.")
    creds_file = st.file_uploader("Upload credentials.json (local only)", type="json")
    if creds_file:
        creds_path = "/tmp/uploaded_credentials.json"
        with open(creds_path, "wb") as f:
            f.write(creds_file.read())
        st.session_state["creds_path"] = creds_path
        st.success("✅ credentials.json loaded")

    st.divider()
    st.caption("Built with ❤️ using Gemini AI + Streamlit")


# ════════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════════

st.markdown("""
<div class="main-header">
    <h1>📝 Assessment Auto-Grader</h1>
    <p>Upload a scanned PDF → Gemini AI extracts & grades → Saves to Google Sheets</p>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([2, 1])
with col1:
    uploaded_pdf = st.file_uploader("📂 Upload Scanned Assessment PDF", type=["pdf"])
with col2:
    st.markdown("### ℹ️ Requirements")
    st.info("PDF can be scanned/handwritten. Should contain **Name**, **Date**, **Topic**, **Assessment No.** and **Q&A**. Duplicates are blocked automatically.")

if uploaded_pdf:
    st.markdown("---")
    if st.button("🚀 Grade This Assessment", use_container_width=True):

        missing = []
        if not api_key_input:  missing.append("Gemini API Key")
        if not sheet_id_input: missing.append("Google Sheet ID")
        if missing:
            st.markdown(f'<div class="error-card">❌ Missing: <b>{", ".join(missing)}</b></div>', unsafe_allow_html=True)
            st.stop()

        pdf_bytes = uploaded_pdf.read()

        with st.spinner("📄 Converting PDF pages to images…"):
            try:
                image_bytes_list = pdf_to_images(pdf_bytes)
                st.success(f"✅ Converted {len(image_bytes_list)} page(s)")
            except Exception as e:
                st.error(f"Failed to read PDF: {e}"); st.stop()

        with st.spinner("🤖 Gemini AI is reading and grading… (20–40 sec)"):
            try:
                result = extract_and_grade_with_gemini(image_bytes_list, api_key_input)
            except json.JSONDecodeError as e:
                st.error(f"AI returned invalid JSON. Try a clearer scan. Details: {e}"); st.stop()
            except Exception as e:
                st.error(f"Gemini API error: {e}"); st.stop()

        st.markdown("## 📋 Extracted Information")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("👤 Name",       result.get("intern_name", "Unknown"))
        m2.metric("📅 Date",       result.get("date", "Unknown"))
        m3.metric("📚 Topic",      result.get("topic", "Unknown"))
        m4.metric("🔢 Assessment", result.get("assessment_number", "Unknown"))

        st.markdown("## 📊 Score")
        s1, s2, s3 = st.columns(3)
        s1.metric("Marks Awarded", f"{result.get('total_marks',0)} / {result.get('max_marks',0)}")
        s2.metric("Percentage",    f"{result.get('percentage',0):.1f}%")
        pct   = result.get("percentage", 0)
        grade = "A+" if pct>=90 else "A" if pct>=80 else "B" if pct>=70 else "C" if pct>=60 else "D" if pct>=50 else "F"
        s3.metric("Grade", grade)

        st.markdown("### 💬 Overall Feedback")
        st.info(result.get("overall_feedback", "No feedback generated."))

        st.markdown("## 📝 Question-wise Breakdown")
        for q in result.get("questions", []):
            awarded = q.get("marks_awarded", 0)
            maximum = q.get("max_marks", 1)
            icon = "✅" if awarded >= maximum else "⚠️" if awarded > 0 else "❌"
            with st.expander(f"{icon} Q{q.get('q_no','?')} — {str(q.get('question',''))[:80]}... | {awarded}/{maximum} marks"):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("**Intern's Answer:**"); st.write(q.get("intern_answer","—"))
                with c2:
                    st.markdown("**Correct Answer:**");  st.write(q.get("correct_answer","—"))
                st.markdown(f"**Feedback:** {q.get('feedback','—')}")

        st.markdown("## 💾 Saving to Google Sheets")
        try:
            ws = get_gsheet_from_secret(sheet_id_input)
            ensure_headers(ws)
            name       = result.get("intern_name", "Unknown")
            assessment = result.get("assessment_number", "Unknown")
            if is_duplicate(ws, name, assessment):
                st.markdown(f'<div class="warning-card">⚠️ <b>Duplicate!</b> {name} — {assessment} already exists. Not saved.</div>', unsafe_allow_html=True)
            else:
                append_to_sheet(ws, result)
                st.markdown('<div class="success-card">✅ Saved to Google Sheets!</div>', unsafe_allow_html=True)
        except Exception as e:
            st.warning(f"⚠️ Google Sheets skipped: {e}")

        st.markdown("## ⬇️ Download Excel Report")
        excel_bytes = build_excel(result)
        safe_name   = re.sub(r"[^a-zA-Z0-9_]", "_", result.get("intern_name", "assessment"))
        filename    = f"{safe_name}_{result.get('assessment_number','result').replace(' ','_')}.xlsx"
        st.download_button("📥 Download Detailed Excel Report", data=excel_bytes,
                           file_name=filename,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)

st.markdown("---")
st.markdown("## 📊 View All Records")
if st.button("🔄 Load Records from Google Sheets", use_container_width=True):
    try:
        ws      = get_gsheet_from_secret(sheet_id_input)
        records = ws.get_all_records()
        if records:
            df = pd.DataFrame(records)
            st.dataframe(df, use_container_width=True, height=400)
            st.caption(f"Total records: {len(df)}")
            out = io.BytesIO()
            df.to_excel(out, index=False)
            st.download_button("📥 Download All as Excel", data=out.getvalue(),
                               file_name="all_assessments.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.info("No records yet.")
    except Exception as e:
        st.error(f"Could not load records: {e}")

st.markdown("<p style='text-align:center;color:#888;font-size:0.85rem;'>Assessment Auto-Grader · Powered by Gemini AI · Built with Streamlit</p>", unsafe_allow_html=True)
