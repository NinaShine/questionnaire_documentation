"""
Documentation Survey - application Streamlit (anonyme, ~6 minutes)
Les réponses sont ajoutées dans un fichier Excel (responses.xlsx).

Lancer :  streamlit run app.py
Admin  :  http://localhost:8501/?admin=1   (mot de passe : variable ADMIN_PASSWORD)
"""

import os
import threading
import uuid
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------
EXCEL_PATH = Path(os.environ.get("SURVEY_EXCEL_PATH", "responses.xlsx"))
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "change-me")
MIN_RESPONSES_FOR_STATS = 5  # seuil d'anonymat : pas de statistiques en dessous
WRITE_LOCK = threading.Lock()

# Palette inspirée de l'identité Sanofi (violet + rose)
PURPLE = "#7A00E6"
DARK_PURPLE = "#23004C"
PINK = "#FF3EB5"
LILAC = "#F3E8FF"

st.set_page_config(
    page_title="Documentation Survey",
    page_icon="📄",
    layout="centered",
)

st.markdown(
    f"""
    <style>
      .stApp {{ background-color: #FBF8FF; }}
      html, body, [class*="css"] {{ font-family: "Segoe UI", Arial, sans-serif; }}
      .banner {{
          background: linear-gradient(135deg, {DARK_PURPLE} 0%, {PURPLE} 100%);
          color: white; padding: 28px 30px; border-radius: 16px; margin-bottom: 18px;
      }}
      .banner h1 {{ color: white; margin: 0 0 6px 0; font-size: 1.9rem; }}
      .banner p  {{ color: #EBDDFF; margin: 0; font-size: 1rem; }}
      .accent    {{ color: {PINK}; }}
      .section-title {{
          color: {DARK_PURPLE}; border-left: 6px solid {PINK};
          padding-left: 12px; margin: 28px 0 8px 0; font-size: 1.35rem; font-weight: 700;
      }}
      .qcard {{
          background: white; border: 1px solid #E4D3FA; border-radius: 12px;
          padding: 4px 18px 6px 18px; margin-bottom: 4px;
      }}
      .note {{
          background: {LILAC}; border-radius: 10px; padding: 12px 16px;
          color: {DARK_PURPLE}; font-size: 0.92rem;
      }}
      div.stButton > button, div.stDownloadButton > button {{
          background-color: {PURPLE}; color: white; border: none;
          border-radius: 10px; padding: 0.6rem 1.6rem; font-weight: 600;
      }}
      div.stButton > button:hover, div.stDownloadButton > button:hover {{
          background-color: {DARK_PURPLE}; color: white;
      }}
      .stProgress > div > div > div > div {{ background-color: {PINK}; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------------------------------
# Contenu du questionnaire
# --------------------------------------------------------------------------
ROLES = [
    "Scrum Master", "Product Owner", "Technical System Owner", "Developer",
    "DevOps", "Quality", "Regulatory / CSV / Validation", "Business", "Other",
]
DOCUMENTS = [
    "User Stories (Jira)", "Functional Specification", "Technical Specification",
    "Functional Risk Assessment (FRA)", "Traceability Matrix", "Validation Strategy",
    "Release Plan", "Release Report", "Release Notes",
    "Test Cases / Test Reports (UATC, IQTC...)",
    "Installation evidence / report (IQ / OQ)", "Authorization for Use (AFU)",
    "Site Impact Analysis", "User / Operator Manual", "Other",
]
TIME_OPTIONS = ["Less than 2 hours", "Half a day", "1 day", "2 days", "More than 2 days"]
FREQ_OPTIONS = ["Per sprint", "Per release", "Per month", "Per change request", "Rarely"]
SOURCES = [
    "Jira", "Confluence", "Previous version of this same document",
    "Other existing documents", "My team's standards or rules",
    "Colleagues or meetings", "My own knowledge",
]
APPROVERS = [
    "Product Owner (PO)", "Business System Owner (BSO)", "Business Owner (BO)",
    "Business Quality Representative (BQR)", "Technical System Owner (TSO)",
    "No formal approval", "Don't know",
]
TIME_SINKS = [
    "Finding and reading the source information",
    "Setting up structure, headings, numbering",
    "Writing genuinely new content",
    "Re-entering content that already exists elsewhere",
    "Formatting and tables",
    "Inserting diagrams or images",
    "Chasing people for information or decisions",
    "Reworking after review",
    "Waiting for approvals",
]
EXISTING_PCT = ["Less than 10%", "10-25%", "25-50%", "More than 50%"]
TOOLS = [
    "Jira", "Confluence", "Word / Excel / PowerPoint", "SharePoint", "Teams",
    "Concierge", "Rovo", "Claude", "Cursor", "Copilot", "Other",
]
AI_USED = [
    "Yes, mostly (AI wrote most of the first draft)",
    "Yes, partly (AI helped with some sections)",
    "No, but I would like to",
    "No",
]
GXP_ACTIVITIES = [
    "Routing for signature and chasing approvers",
    "Version control (which copy is current)",
    "Change control",
    "Checking which other documents a change affects",
    "Traceability between related documents",
    "Periodic review",
    "Audit or inspection requests",
    "None of these",
]
USE_TOOL = [
    "Yes, straight away",
    "Yes, but only if a human reviews it before approval",
    "Unsure",
    "No",
]
HESITATIONS = [
    "Accuracy of the output",
    "Accountability if the AI is wrong",
    "No traceability to the source requirement",
    "Not approved by Quality / Compliance for GxP use",
    "Data must stay inside Sanofi (approved tool only)",
    "AI-generated content is not clearly labelled as such",
    "I prefer to keep full control",
    "Other",
]
FOLLOWUP = ["Yes", "No"]

COLUMNS = [
    "response_id", "timestamp", "q1_role", "q2_document", "q2_other",
    "q3_time", "q3_frequency", "q3_sources", "q3_approvers",
    "q4_time_sinks", "q4_pct_existing", "q5_tools", "q5_ai_used",
    "q6_gxp_activities", "q7_use_tool", "q7_hesitations", "q8_one_problem",
    "q8_followup",
]

# --------------------------------------------------------------------------
# Excel
# --------------------------------------------------------------------------
def _new_workbook() -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = "Responses"
    ws.append(COLUMNS)
    header_fill = PatternFill("solid", start_color=PURPLE.lstrip("#"))
    for cell in ws[1]:
        cell.font = Font(name="Arial", bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    return wb


def save_response(row: dict) -> None:
    """Ajoute une ligne dans le fichier Excel (créé si besoin)."""
    with WRITE_LOCK:
        wb = load_workbook(EXCEL_PATH) if EXCEL_PATH.exists() else _new_workbook()
        ws = wb["Responses"]
        ws.append([row.get(col, "") for col in COLUMNS])
        for col_idx, col in enumerate(COLUMNS, start=1):
            ws.column_dimensions[get_column_letter(col_idx)].width = (
                14 if col in ("response_id", "timestamp") else 32
            )
        for cell in ws[ws.max_row]:
            cell.font = Font(name="Arial")
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        wb.save(EXCEL_PATH)


def load_responses() -> pd.DataFrame:
    if not EXCEL_PATH.exists():
        return pd.DataFrame(columns=COLUMNS)
    return pd.read_excel(EXCEL_PATH, sheet_name="Responses")


# --------------------------------------------------------------------------
# Page administrateur
# --------------------------------------------------------------------------
def admin_page() -> None:
    st.markdown(
        '<div class="banner"><h1>Admin <span class="accent">·</span> Results</h1>'
        "<p>Download the Excel file and see the aggregated results.</p></div>",
        unsafe_allow_html=True,
    )
    pwd = st.text_input("Admin password", type="password")
    if pwd != ADMIN_PASSWORD:
        if pwd:
            st.error("Wrong password.")
        return

    df = load_responses()
    st.metric("Responses collected", len(df))
    if EXCEL_PATH.exists():
        st.download_button(
            "⬇️ Download responses.xlsx",
            data=EXCEL_PATH.read_bytes(),
            file_name="responses.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if len(df) < MIN_RESPONSES_FOR_STATS:
        st.info(
            f"Aggregated charts are hidden until there are at least "
            f"{MIN_RESPONSES_FOR_STATS} responses (anonymity threshold)."
        )
        return

    st.markdown('<div class="section-title">Overview</div>', unsafe_allow_html=True)
    for col, label in [
        ("q2_document", "Document that takes the most time"),
        ("q3_time", "Time to produce one"),
        ("q7_use_tool", "Would use an AI-generated first draft"),
    ]:
        st.subheader(label)
        st.bar_chart(df[col].value_counts())

    st.subheader("Where the time goes (multi-select)")
    counts = df["q4_time_sinks"].dropna().str.split(" | ", regex=False).explode().value_counts()
    st.bar_chart(counts)


# --------------------------------------------------------------------------
# Questionnaire
# --------------------------------------------------------------------------
def reset_form() -> None:
    for key in list(st.session_state.keys()):
        if key.startswith("f_") or key == "submitted":
            del st.session_state[key]


def survey_page() -> None:
    if st.session_state.get("submitted"):
        st.markdown(
            '<div class="banner"><h1>Thank you! <span class="accent">✔</span></h1>'
            "<p>Your answers were recorded anonymously. We will share a short summary "
            "of the results with everyone who takes part.</p></div>",
            unsafe_allow_html=True,
        )
        if st.button("Submit another response"):
            reset_form()
            st.rerun()
        return

    st.markdown(
        """
        <div class="banner">
          <h1>Help us spend less time on documents<span class="accent">.</span></h1>
          <p>7 short questions · about 6 minutes · anonymous</p>
        </div>
        <div class="note">
          This survey is anonymous: we do not collect your name or email, and no
          individual responses will be shared, only aggregated themes.
          Participation is voluntary.<br>
        </div>
        """,
        unsafe_allow_html=True,
    )

    ss = st.session_state
    progress_slot = st.empty()

    # ---------------- Part 1 ----------------
    st.markdown('<div class="section-title">Part 1 · Your documents</div>', unsafe_allow_html=True)

    st.markdown('<div class="qcard">', unsafe_allow_html=True)
    q1 = st.radio("**Q1. Which role describes you best?**", ROLES, index=None, key="f_q1")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="qcard">', unsafe_allow_html=True)
    q2 = st.radio(
        "**Q2. Which ONE document takes the most of your time overall?**",
        DOCUMENTS, index=None, key="f_q2",
    )
    q2_other = ""
    if q2 == "Other":
        q2_other = st.text_input("Please specify", key="f_q2_other")
    st.markdown("</div>", unsafe_allow_html=True)

    doc_label = q2 if q2 and q2 != "Other" else "this document"
    st.caption(f"The next questions are about: **{doc_label}**")

    st.markdown('<div class="qcard">', unsafe_allow_html=True)
    st.markdown("**Q3. How is this document produced?**")
    q3_time = st.radio(
        "Time to produce one (drafting only, not including review and approval)",
        TIME_OPTIONS, index=None, key="f_q3_time", horizontal=True,
    )
    q3_freq = st.radio(
        "How often do you produce it?", FREQ_OPTIONS, index=None,
        key="f_q3_freq", horizontal=True,
    )
    q3_sources = st.multiselect(
        "Where does the information come from? (all that apply)", SOURCES, key="f_q3_sources"
    )
    q3_approvers = st.multiselect(
        "Who formally approves it? (all that apply)", APPROVERS, key="f_q3_approvers"
    )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="qcard">', unsafe_allow_html=True)
    q4 = st.multiselect(
        "**Q4. Where does the time go?** (choose up to 2)",
        TIME_SINKS, max_selections=2, key="f_q4",
    )
    q4_pct = st.radio(
        "How much of this document already exists elsewhere?",
        EXISTING_PCT, index=None, key="f_q4_pct", horizontal=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="qcard">', unsafe_allow_html=True)
    st.markdown("**Q5. Tools and AI**")
    q5_tools = st.multiselect(
        "Tools you use for documents (all that apply)", TOOLS, key="f_q5_tools"
    )
    q5_ai = st.radio("Was AI used to draft this document?", AI_USED, index=None, key="f_q5_ai")
    st.markdown("</div>", unsafe_allow_html=True)

    # ---------------- Part 2 ----------------
    st.markdown(
        '<div class="section-title">Part 2 · Regulated documents (GxP)</div>',
        unsafe_allow_html=True,
    )

    st.markdown('<div class="qcard">', unsafe_allow_html=True)
    q6 = st.multiselect(
        "**Q6. Which controlled-document activities take a disproportionate amount of time?** "
        "(choose up to 2)",
        GXP_ACTIVITIES, max_selections=2, key="f_q6",
    )
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="qcard">', unsafe_allow_html=True)
    q7 = st.radio(
        "**Q7. Would you use a tool that generates a first draft of this document?**",
        USE_TOOL, index=None, key="f_q7",
    )
    q7_hes = []
    show_hesitation = q7 is not None and q7 != "Yes, straight away"
    if show_hesitation:
        q7_hes = st.multiselect(
            "What would make you hesitate? (choose up to 3)",
            HESITATIONS, max_selections=3, key="f_q7_hes",
        )
    st.markdown("</div>", unsafe_allow_html=True)

    # ---------------- Closing ----------------
    st.markdown('<div class="section-title">Closing (optional)</div>', unsafe_allow_html=True)
    st.markdown('<div class="qcard">', unsafe_allow_html=True)
    q8 = st.text_area(
        "**Q8. If you could fix ONE documentation problem this year, what would it be?**",
        max_chars=300, key="f_q8", height=90,
    )
    q8_follow = st.radio(
        "Would you take part in a short follow-up discussion?",
        FOLLOWUP, index=None, horizontal=True, key="f_q8_follow",
    )
    if q8_follow == "Yes":
        st.info(
            "Thank you! To keep your answers anonymous, please register through the "
            "separate follow-up link shared with this survey (do not write your name here)."
        )
    st.markdown("</div>", unsafe_allow_html=True)

    # ---------------- Progression ----------------
    required = {
        "Q1": q1, "Q2": q2, "Q3 time": q3_time, "Q3 frequency": q3_freq,
        "Q3 sources": q3_sources, "Q3 approvers": q3_approvers, "Q4": q4,
        "Q4 % existing": q4_pct, "Q5 tools": q5_tools, "Q5 AI": q5_ai,
        "Q6": q6, "Q7": q7,
    }
    if show_hesitation:
        required["Q7 hesitation"] = q7_hes
    if q2 == "Other":
        required["Q2 specify"] = q2_other.strip()
    missing = [name for name, val in required.items() if not val]
    done = len(required) - len(missing)
    progress_slot.progress(done / len(required), text=f"{done}/{len(required)} required answers")

    if st.button("Submit my answers"):
        if missing:
            st.error("Please answer: " + ", ".join(missing))
            return
        row = {
            "response_id": uuid.uuid4().hex[:8],
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "q1_role": q1,
            "q2_document": q2,
            "q2_other": q2_other.strip(),
            "q3_time": q3_time,
            "q3_frequency": q3_freq,
            "q3_sources": " | ".join(q3_sources),
            "q3_approvers": " | ".join(q3_approvers),
            "q4_time_sinks": " | ".join(q4),
            "q4_pct_existing": q4_pct,
            "q5_tools": " | ".join(q5_tools),
            "q5_ai_used": q5_ai,
            "q6_gxp_activities": " | ".join(q6),
            "q7_use_tool": q7,
            "q7_hesitations": " | ".join(q7_hes),
            "q8_one_problem": q8.strip(),
            "q8_followup": q8_follow or "",
        }
        save_response(row)
        ss["submitted"] = True
        st.rerun()


# --------------------------------------------------------------------------
# Routage
# --------------------------------------------------------------------------
if st.query_params.get("admin") == "1":
    admin_page()
else:
    survey_page()