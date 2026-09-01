import streamlit as st
import pandas as pd
from datetime import datetime
import io
import uuid
import os
import smtplib
from email.message import EmailMessage
import gspread
from google.oauth2.service_account import Credentials

# Optional PDF dependency
try:
    from fpdf import FPDF
    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False

# Optional Excel styling dependency
try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# =========================================================
# APP CONFIGURATION
# =========================================================
st.set_page_config(page_title="Weekly Project Report", page_icon="📈", layout="wide")

# =========================================================
# SECURITY LOCK SCREEN
# =========================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.markdown("## 🔒 App Locked")
    st.write("Please enter the team password to access the Weekly Project Report.")
    entered_pwd = st.text_input("Password", type="password")
    if st.button("Unlock"):
        if entered_pwd == st.secrets["APP_PASSWORD"]:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop() # This entirely stops the rest of the app from loading!
    
# =========================================================
# GOOGLE SHEETS BACKEND CONFIGURATION (FIXED)
# =========================================================
@st.cache_resource
def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)

def load_data_from_gsheets():
    try:
        client = get_gspread_client()
        sheet = client.open("Weekly Notes Database").sheet1
        return sheet.get_all_records()
    except Exception as e:
        st.error(f"Failed to connect to Google Sheets. Error: {e}")
        return []

def save_new_notes_to_gsheets(new_entries):
    client = get_gspread_client()
    sheet = client.open("Weekly Notes Database").sheet1
    headers = sheet.row_values(1)
    rows_to_append = [[str(entry.get(col, "")) for col in headers] for entry in new_entries]
    sheet.append_rows(rows_to_append)

def update_existing_note_in_gsheets(updated_entry):
    client = get_gspread_client()
    sheet = client.open("Weekly Notes Database").sheet1
    records = sheet.get_all_records()
    headers = sheet.row_values(1)
    for idx, row in enumerate(records):
        if str(row.get("id")) == str(updated_entry.get("id")):
            row_idx = idx + 2 
            update_values = [[str(updated_entry.get(col, "")) for col in headers]]
            sheet.update(f"A{row_idx}:{get_column_letter(len(headers))}{row_idx}", update_values)
            break

# =========================================================
# PREMIUM UI SYSTEM
# =========================================================
st.markdown("""
<style>
:root {
    --app-bg: #f4f7fb; --surface: #ffffff; --primary: #16324f; --accent: #2563eb; --border: #dfe5ec;
    --text-main: #0f172a; --muted: #64748b; --radius-sm: 8px; --radius-md: 12px;
}
.stApp { background: var(--app-bg); }
.block-container { max-width: 1440px !important; margin: 0 auto !important; padding-top: 1.5rem !important; padding-bottom: 2.5rem !important; }
.stMarkdown, .stText, .stCaption, .stSelectbox label, .stMultiSelect label, .stTextInput label, .stTextArea label, .stNumberInput label, .stDateInput label { font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif !important; }
.stSelectbox label p, .stMultiSelect label p, .stTextInput label p, .stTextArea label p, .stNumberInput label p, .stDateInput label p, .stSelectbox label, .stMultiSelect label, .stTextInput label, .stTextArea label, .stNumberInput label, .stDateInput label { font-size: 15px !important; font-weight: bold !important; color: #16324f !important; letter-spacing: .25px !important; margin-bottom: 4px !important; }
[data-baseweb="select"] > div { background: linear-gradient(to bottom, #ffffff, #f8fafc) !important; border: 1px solid #cbd5e1 !important; box-shadow: 0 2px 4px rgba(0,0,0,0.03), 0 1px 2px rgba(0,0,0,0.04) !important; border-radius: 8px !important; min-height: 44px !important; transition: all 0.15s ease !important; }
[data-baseweb="select"] > div:hover { border-color: #94a3b8 !important; box-shadow: 0 4px 6px rgba(0,0,0,0.06), 0 2px 4px rgba(0,0,0,0.05) !important; }
.stSelectbox [data-baseweb="select"] span, .stMultiSelect [data-baseweb="select"] span { font-size: 15px !important; }
.stMultiSelect [data-baseweb="tag"] { font-size: 14px !important; font-weight: bold !important; }
.app-hero { background: linear-gradient(135deg, #132d47 0%, #1c4569 68%, #235c85 100%); border-radius: 18px; padding: 28px 32px; min-height: 110px; box-sizing: border-box; color: #ffffff !important; box-shadow: 0 12px 30px rgba(22, 50, 79, .18); margin-bottom: 16px; position: relative; overflow: hidden; isolation: isolate; }
.app-hero:after { content: ""; position: absolute; width: 260px; height: 260px; right: -80px; top: -120px; border-radius: 50%; background: rgba(255,255,255,.08); }
.app-hero > * { position: relative; z-index: 2; }
.app-hero-eyebrow { font-size: 13px; font-weight: bold; letter-spacing: 1.7px; text-transform: uppercase; color: #9ed8ff; margin-bottom: 6px; }
.app-hero-title { font-size: 34px; line-height: 1.15; font-weight: bold; letter-spacing: -.6px; margin: 0; color: #fff !important; }
.app-hero-subtitle { margin-top: 8px; font-size: 15px; color: #d6e4f1 !important; }
.stButton > button, .stDownloadButton > button { background: linear-gradient(to bottom, #ffffff, #f1f5f9) !important; border: 1px solid #cbd5e1 !important; box-shadow: 0 3px 0 #cbd5e1, 0 4px 6px rgba(0,0,0,0.05) !important; border-radius: 10px !important; color: #334155 !important; font-weight: bold !important; transition: all 0.1s ease-out !important; transform: translateY(0) !important; }
.stButton > button:active, .stDownloadButton > button:active { transform: translateY(3px) !important; box-shadow: 0 0px 0 transparent, 0 0px 0px rgba(0,0,0,0) !important; }
.stButton > button[kind="primary"] { background: linear-gradient(to bottom, var(--primary), #112840) !important; border-color: #0b1a2a !important; box-shadow: 0 3px 0 #0b1a2a, 0 5px 12px rgba(22, 50, 79, 0.3) !important; color: white !important; }
div.element-container:has(.marker-nav1-inactive) + div.element-container button { background: linear-gradient(to bottom, #f8fafc, #eff6ff) !important; border-color: #bfdbfe !important; box-shadow: 0 3px 0 #bfdbfe, 0 4px 6px rgba(0,0,0,0.05) !important; color: #1e40af !important; }
div.element-container:has(.marker-nav1-active) + div.element-container button { background: linear-gradient(to bottom, #3b82f6, #1d4ed8) !important; border-color: #1e40af !important; box-shadow: 0 3px 0 #1e40af, 0 4px 8px rgba(37,99,235,0.3) !important; color: #ffffff !important; }
div.element-container:has(.marker-nav2-inactive) + div.element-container button { background: linear-gradient(to bottom, #f8fafc, #f0fdf4) !important; border-color: #bbf7d0 !important; box-shadow: 0 3px 0 #bbf7d0, 0 4px 6px rgba(0,0,0,0.05) !important; color: #166534 !important; }
div.element-container:has(.marker-nav2-active) + div.element-container button { background: linear-gradient(to bottom, #10b981, #047857) !important; border-color: #064e3b !important; box-shadow: 0 3px 0 #064e3b, 0 4px 8px rgba(16,185,129,0.3) !important; color: #ffffff !important; }
div.element-container:has(.marker-nav3-inactive) + div.element-container button { background: linear-gradient(to bottom, #f8fafc, #fffbeb) !important; border-color: #fde68a !important; box-shadow: 0 3px 0 #fde68a, 0 4px 6px rgba(0,0,0,0.05) !important; color: #b45309 !important; }
div.element-container:has(.marker-nav3-active) + div.element-container button { background: linear-gradient(to bottom, #f59e0b, #b45309) !important; border-color: #78350f !important; box-shadow: 0 3px 0 #78350f, 0 4px 8px rgba(245,158,11,0.3) !important; color: #ffffff !important; }
div.element-container:has(.marker-send) + div.element-container button { background: linear-gradient(to bottom, #f5f3ff, #ede9fe) !important; border-color: #c4b5fd !important; box-shadow: 0 3px 0 #c4b5fd, 0 4px 6px rgba(0,0,0,0.05) !important; color: #5b21b6 !important; }
div.element-container:has(.marker-pdf) + div.element-container button { background: linear-gradient(to bottom, #fff1f2, #ffe4e6) !important; border-color: #fecdd3 !important; box-shadow: 0 3px 0 #fecdd3, 0 4px 6px rgba(0,0,0,0.05) !important; color: #9f1239 !important; }
div.element-container:has(.marker-excel) + div.element-container button { background: linear-gradient(to bottom, #ecfdf5, #d1fae5) !important; border-color: #a7f3d0 !important; box-shadow: 0 3px 0 #a7f3d0, 0 4px 6px rgba(0,0,0,0.05) !important; color: #065f46 !important; }
div.element-container:has([class^="marker-"]) + div.element-container button:active { transform: translateY(3px) !important; box-shadow: 0 0px 0 transparent, 0 0px 0px rgba(0,0,0,0) !important; }
.page-head { background: var(--surface); border: 1px solid var(--border); border-radius: 14px; padding: 16px 20px; box-shadow: var(--shadow-sm); margin-bottom: 15px; min-height: 80px; display: flex; flex-direction: column; justify-content: center; }
.page-title { font-size: 28px; line-height: 1.2; font-weight: bold; color: var(--primary) !important; margin: 0; }
.page-subtitle { color: var(--muted) !important; font-size: 15px; margin-top: 5px; }
.section-header { font-size: 22px; font-weight: bold; color: var(--primary) !important; margin: 18px 0 10px 0; display: flex; align-items: center; gap: 8px; }
.section-header:before { content: ""; width: 5px; height: 22px; background: var(--accent); border-radius: 3px; }
[data-testid="stVerticalBlockBorderWrapper"] { border-radius: 12px !important; border-color: var(--border) !important; box-shadow: var(--shadow-sm) !important; background: var(--surface) !important; }
.card-title { font-size: 17px; font-weight: bold; color: var(--primary) !important; margin-bottom: 3px; }
.kpi-card { background: var(--surface); border: 1px solid var(--border); border-radius: 13px; padding: 15px 18px; box-shadow: var(--shadow-sm); min-height: 85px; border-top: 4px solid var(--blue); transition: transform .18s ease, box-shadow .18s ease; }
.kpi-card:hover { transform: translateY(-2px); box-shadow: var(--shadow-md); }
.kpi-label { font-size: 12px; color: var(--muted) !important; font-weight: bold; letter-spacing: .8px; }
.kpi-value { margin-top: 5px; font-size: 32px; line-height: 1; font-weight: bold; color: var(--primary) !important; }
.custom-table { width: 100%; border-collapse: separate; border-spacing: 0; background: #fff; border: 1px solid var(--border); border-radius: 12px; overflow: hidden; font-size: 15px; box-shadow: var(--shadow-sm); }
.custom-table th { background: #eef3f8; color: #344054; padding: 13px 16px; text-align: left; font-size: 13.5px; font-weight: bold; text-transform: uppercase; letter-spacing: .45px; border-bottom: 1px solid var(--border); }
.custom-table td { padding: 14px 16px; border-bottom: 1px solid #edf1f5; color: #344054; vertical-align: middle; }
.custom-table tr:hover td { background: #f8fbff; }
.proj-name { font-size: 15px; font-weight: bold; color: var(--primary); display:block; }
.work-type { font-size: 13px; color: var(--muted); display:block; margin-top:2px; }
.status-pill { display: inline-flex; align-items: center; padding: 5px 9px; border-radius: 999px; font-size: 12px; font-weight: bold; border: 1px solid transparent; }
.status-on-track { background:#ecfdf5; color:#16805b; border-color:#a7f3d0; }
.status-in-progress { background:#fffbeb; color:#b7791f; border-color:#fde68a; }
.status-blocked { background:#fef2f2; color:#c2413d; border-color:#fecaca; }
.status-completed { background:#eff6ff; color:#2563eb; border-color:#bfdbfe; }
.status-not-started { background:#f2f4f7; color:#667085; border-color:#d0d5dd; }
.summary-box { background: #f8fafc; border-radius: 8px; color: #334155; font-size: 1.15rem; line-height: 1.6; margin-bottom: 2.5rem; padding: 1.25rem; border: 1px solid #e2e8f0; }
.risk-box { background-color: #fef2f2; border-left: 5px solid #dc2626; padding: 1.25rem; margin-bottom: 1rem; border-radius: 6px; }
.risk-title { color: #0f172a; font-size: 1.05rem; font-weight: 700; margin-bottom: 0.25rem; }
.risk-desc { color: #334155; font-size: 0.95rem; }
.acc-header { background: linear-gradient(90deg,#173654,#214f77); color:#fff !important; padding:9px 14px; border-radius:10px; font-size:17px; font-weight:bold; margin:13px 0 7px 0; }
.team-header { color:var(--primary) !important; font-size:15px; font-weight:bold; padding:5px 2px; border-bottom:2px solid #e6ebf1; margin:7px 0 5px; }
.edit-banner { background:#eff6ff; color:#2563eb !important; border:1px solid #bfdbfe; padding:8px 11px; border-radius:9px; font-size:13px; font-weight:bold; margin-bottom:10px; }
.block-container .stVerticalBlock { gap: .8rem !important; }
hr { margin: 12px 0 !important; border-color:#e6ebf1 !important; }
</style>
""", unsafe_allow_html=True)

# =========================================================
# SESSION STATE & LIVE DATA SYNC
# =========================================================
if 'data_synced' not in st.session_state:
    st.session_state.notes_db = load_data_from_gsheets()
    st.session_state.data_synced = True

if 'project_count' not in st.session_state: st.session_state.project_count = 1
if 'show_success' not in st.session_state: st.session_state.show_success = False
if 'success_message' not in st.session_state: st.session_state.success_message = ""
if 'edit_id' not in st.session_state: st.session_state.edit_id = None
if 'nav_selection' not in st.session_state: st.session_state.nav_selection = "📋 Weekly Summary"
if 'show_send_dialog' not in st.session_state: st.session_state.show_send_dialog = None

# Dynamically populate account and team options based on cloud data
st.session_state.accounts_db = {}
if st.session_state.notes_db:
    live_df = pd.DataFrame(st.session_state.notes_db)
    for acc in live_df["Account"].unique():
        st.session_state.accounts_db[acc] = live_df[live_df["Account"] == acc]["Team"].unique().tolist()

# =========================================================
# HELPERS
# =========================================================
def get_pct_decimal(pct_str):
    try: return float(str(pct_str).replace('%', '').strip()) / 100.0
    except Exception: return 0.0

def get_status_icon(status):
    icons = {"On Track": "●", "At Risk": "●", "Blocked": "●", "Completed": "●", "Not Started": "●"}
    return icons.get(status, "●")

def render_metric_card(label, value):
    return f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div></div>'

def format_badge_text(selected_list, all_options):
    if not selected_list: return "None"
    if len(selected_list) == len(all_options): return "All"
    if len(selected_list) <= 2: return ", ".join(map(str, selected_list))
    return f"{len(selected_list)} selected"

def parse_date(date_str):
    if not date_str or str(date_str).strip() in ["N/A", "None", ""]: return None
    try: return datetime.strptime(str(date_str), "%Y-%m-%d").date()
    except Exception: return None

def safe_pdf_text(value):
    text = "" if value is None else str(value)
    replacements = {"–":"-", "—":"-", "’":"'", "“":"\"", "”":"\"", "•":"-", "✓":"OK", "⚠":"!", "×":"x"}
    for old, new in replacements.items(): text = text.replace(old, new)
    return text.encode("latin-1", "replace").decode("latin-1")

def create_pdf_report(dataframe, header_title):
    if not FPDF_AVAILABLE: return None
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=17)
    pdf.add_page()
    pdf.set_fill_color(22, 50, 79)
    pdf.rect(0, 0, 210, 44, "F")
    pdf.set_text_color(158, 216, 255)
    pdf.set_font("Arial", "B", 8)
    pdf.set_xy(15, 9)
    pdf.cell(170, 4, safe_pdf_text("WEEKLY PROJECT MANAGEMENT"), ln=True)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", "B", 21)
    pdf.set_xy(15, 15)
    pdf.cell(170, 9, safe_pdf_text("WEEKLY PROJECT REPORT"), ln=True)
    pdf.set_text_color(218, 231, 242)
    pdf.set_font("Arial", "", 9.5)
    pdf.set_xy(15, 27)
    pdf.cell(170, 5, safe_pdf_text(header_title), ln=True)
    pdf.set_font("Arial", "", 7.5)
    pdf.set_xy(15, 35)
    pdf.cell(170, 4, safe_pdf_text(f"Generated: {datetime.now().strftime('%d %b %Y, %I:%M %p')}"), ln=True)

    total_projects = len(dataframe)
    total_resources = dataframe["Resource"].nunique() if not dataframe.empty and "Resource" in dataframe.columns else 0
    delivered = int((dataframe["This Week Delivered"].fillna("").astype(str).str.strip() != "").sum()) if not dataframe.empty else 0
    avg_comp = int(dataframe["Completion %"].apply(lambda x: get_pct_decimal(x) * 100).mean()) if not dataframe.empty else 0
    avg_util = int(dataframe["Utilization %"].apply(lambda x: get_pct_decimal(x) * 100).mean()) if not dataframe.empty else 0
    risks = int(((dataframe["Status"].isin(["At Risk", "Blocked"])) | (dataframe["Blocker"].fillna("").astype(str).str.strip() != "")).sum()) if not dataframe.empty else 0

    kpis = [("PROJECTS", total_projects), ("RESOURCES", total_resources), ("UPDATES", delivered), ("COMPLETION", f"{avg_comp}%"), ("RISKS", risks)]
    box_w = 34.5; gap = 2.7; y = 50
    for i, (label, value) in enumerate(kpis):
        x = 15 + i * (box_w + gap)
        pdf.set_fill_color(247, 249, 252); pdf.set_draw_color(221, 228, 236)
        pdf.rect(x, y, box_w, 24, "DF")
        pdf.set_text_color(102, 112, 133); pdf.set_font("Arial", "B", 6.8); pdf.set_xy(x + 2, y + 4)
        pdf.cell(box_w - 4, 4, safe_pdf_text(label), align="C")
        pdf.set_text_color(22, 50, 79); pdf.set_font("Arial", "B", 15); pdf.set_xy(x + 2, y + 10)
        pdf.cell(box_w - 4, 8, safe_pdf_text(value), align="C")
    pdf.set_y(82)

    def section(title):
        pdf.ln(5); y0 = pdf.get_y(); pdf.set_fill_color(37, 99, 235); pdf.rect(15, y0 + 1, 2.5, 7, "F")
        pdf.set_xy(21, y0); pdf.set_text_color(22, 50, 79); pdf.set_font("Arial", "B", 13)
        pdf.cell(170, 8, safe_pdf_text(title), ln=True); pdf.ln(1)

    section("Executive Summary")
    pdf.set_text_color(52, 64, 84); pdf.set_font("Arial", "", 9.5)
    summary = (f"The selected report contains {total_projects} project update(s) across {total_resources} resource(s). "
               f"{delivered} project(s) include a weekly delivery update. Average completion is {avg_comp}% and average utilization is {avg_util}%. "
               f"There are {risks} item(s) requiring risk or blocker attention.")
    pdf.multi_cell(180, 5.8, safe_pdf_text(summary))

    section("Project Status")
    headers = [("Project", 60), ("Resource", 28), ("Status", 28), ("Comp.", 17), ("Util.", 17), ("Expected", 30)]
    pdf.set_fill_color(231, 238, 246); pdf.set_text_color(37, 54, 74); pdf.set_font("Arial", "B", 7.8)
    for label, width in headers: pdf.cell(width, 8, safe_pdf_text(label), border=1, fill=True)
    pdf.ln(); pdf.set_font("Arial", "", 8.2)
    for ridx, (_, row) in enumerate(dataframe.iterrows()):
        if ridx % 2 == 1: pdf.set_fill_color(249, 251, 253)
        else: pdf.set_fill_color(255, 255, 255)
        project = safe_pdf_text(row.get("Project / Dashboard", "N/A"))[:40]; resource = safe_pdf_text(row.get("Resource", "N/A"))[:17]
        status = safe_pdf_text(row.get("Status", "N/A"))[:15]; comp = safe_pdf_text(row.get("Completion %", "N/A"))
        util = safe_pdf_text(row.get("Utilization %", "N/A")); expected = safe_pdf_text(row.get("Updated Expected Delivery date", "N/A"))
        pdf.set_text_color(52, 64, 84); pdf.cell(60, 8, project, border=1, fill=True); pdf.cell(28, 8, resource, border=1, fill=True)
        if status in ("At Risk", "Blocked"): pdf.set_text_color(160, 55, 55)
        elif status == "Completed": pdf.set_text_color(22, 128, 91)
        pdf.cell(28, 8, status, border=1, fill=True); pdf.set_text_color(52, 64, 84)
        pdf.cell(17, 8, comp, border=1, fill=True, align="C"); pdf.cell(17, 8, util, border=1, fill=True, align="C")
        pdf.cell(30, 8, expected, border=1, fill=True); pdf.ln()

    section("Weekly Highlights")
    for _, row in dataframe.iterrows():
        project = safe_pdf_text(row.get("Project / Dashboard", "N/A")); delivered_text = safe_pdf_text(row.get("This Week Delivered", "None reported."))
        pdf.set_text_color(22, 50, 79); pdf.set_font("Arial", "B", 9); pdf.multi_cell(180, 5, project)
        pdf.set_text_color(71, 84, 103); pdf.set_font("Arial", "", 9); pdf.multi_cell(180, 5, f"Delivered: {delivered_text or 'None reported.'}"); pdf.ln(1.5)

    risk_df = dataframe[(dataframe["Status"].isin(["At Risk", "Blocked"])) | (dataframe["Blocker"].fillna("").astype(str).str.strip() != "")]
    section("Risks & Blockers")
    if risk_df.empty:
        pdf.set_text_color(22, 128, 91); pdf.set_font("Arial", "B", 9); pdf.multi_cell(180, 6, safe_pdf_text("No active blockers or risks reported for this period."))
    else:
        for _, row in risk_df.iterrows():
            pdf.set_text_color(194, 65, 61); pdf.set_font("Arial", "B", 9); pdf.multi_cell(180, 5, safe_pdf_text(row.get("Project / Dashboard", "N/A")))
            pdf.set_text_color(71, 84, 103); pdf.set_font("Arial", "", 9)
            blocker = row.get("Blocker", "") or f"Status marked as: {row.get('Status', 'N/A')}"
            pdf.multi_cell(180, 5, safe_pdf_text(blocker)); pdf.ln(1.5)

    page_count = pdf.page_no()
    for page in range(1, page_count + 1):
        pdf.page = page; pdf.set_y(-12); pdf.set_draw_color(221, 228, 236); pdf.line(15, 286, 195, 286)
        pdf.set_text_color(120, 130, 145); pdf.set_font("Arial", "", 7); pdf.cell(90, 5, "Weekly Project Report", align="L"); pdf.cell(90, 5, f"Page {page} of {page_count}", align="R")

    try: output = pdf.output(dest="S"); return output.encode("latin-1") if isinstance(output, str) else bytes(output)
    except Exception: return bytes(pdf.output())

def create_excel_report(dataframe, report_title):
    buffer = io.BytesIO()
    if not OPENPYXL_AVAILABLE:
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer: dataframe.to_excel(writer, index=False, sheet_name="Weekly Notes")
        return buffer.getvalue()

    wb = Workbook(); ws = wb.active; ws.title = "Weekly Notes"
    cols = ["Resource", "Business Owner", "Project / Dashboard", "Start date", "Initial Delivery date", "Updated Expected Delivery date", "Work Type", "This Week Delivered", "Next Week Priority", "Completion %", "Utilization %", "Status", "Blocker"]
    date_str = report_title.split(": ")[-1] if ": " in report_title else report_title

    ws["A1"] = date_str; ws["A1"].fill = PatternFill("solid", fgColor="FFD966"); ws["A1"].font = Font(bold=True); ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws["B1"] = f"Weekly Notes ({date_str})"; ws["B1"].fill = PatternFill("solid", fgColor="004B87"); ws["B1"].font = Font(bold=True, color="FFFFFF"); ws["B1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=len(cols)); ws.row_dimensions[1].height = 25

    header_fill = PatternFill("solid", fgColor="DDEBF7"); header_font = Font(bold=True); thin_border = Border(left=Side(style='thin', color='000000'), right=Side(style='thin', color='000000'), top=Side(style='thin', color='000000'), bottom=Side(style='thin', color='000000'))

    for c_idx, col_name in enumerate(cols, 1):
        cell = ws.cell(row=2, column=c_idx, value=col_name); cell.fill = header_fill; cell.font = header_font; cell.border = thin_border; cell.alignment = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
    ws.row_dimensions[2].height = 35

    resource_fill = PatternFill("solid", fgColor="FCE4D6"); owner_fill = PatternFill("solid", fgColor="E2EFDA")    

    for r_idx, row in enumerate(dataframe.to_dict('records'), 3):
        for c_idx, col_name in enumerate(cols, 1):
            val = row.get(col_name, ""); cell = ws.cell(row=r_idx, column=c_idx, value=val); cell.border = thin_border; cell.alignment = Alignment(vertical="top", wrap_text=True)
            if col_name == "Resource": cell.fill = resource_fill; cell.font = Font(bold=True, color="004B87")
            elif col_name == "Business Owner": cell.fill = owner_fill; cell.font = Font(bold=True, color="004B87")
            elif col_name == "Status": cell.font = Font(bold=True)

    widths = {"A": 16, "B": 14, "C": 30, "D": 11, "E": 11, "F": 13, "G": 22, "H": 38, "I": 35, "J": 12, "K": 12, "L": 14, "M": 38}
    for col_letter, width in widths.items(): ws.column_dimensions[col_letter].width = width

    ws.freeze_panes = "A3"; wb.save(buffer)
    return buffer.getvalue()

def send_report_email(pdf_bytes, report_title, recipients, cc_recipients=None):
    recipients = [x.strip() for x in recipients if x.strip()]; cc_recipients = [x.strip() for x in (cc_recipients or []) if x.strip()]
    if not recipients: return False, "Please enter at least one recipient email address."
    cfg = st.secrets if hasattr(st, "secrets") else {}; host = cfg.get("SMTP_HOST", os.getenv("SMTP_HOST", "")); port = int(cfg.get("SMTP_PORT", os.getenv("SMTP_PORT", "587"))); username = cfg.get("SMTP_USERNAME", os.getenv("SMTP_USERNAME", "")); password = cfg.get("SMTP_PASSWORD", os.getenv("SMTP_PASSWORD", "")); sender = cfg.get("SMTP_FROM", os.getenv("SMTP_FROM", username)); use_tls = str(cfg.get("SMTP_USE_TLS", os.getenv("SMTP_USE_TLS", "true"))).lower() == "true"
    if not host or not sender: return False, "Email sending is not configured (SMTP secrets missing)."
    msg = EmailMessage(); msg["Subject"] = report_title; msg["From"] = sender; msg["To"] = ", ".join(recipients)
    if cc_recipients: msg["Cc"] = ", ".join(cc_recipients)
    msg.set_content(f"Please find attached the {report_title}.\n\nThis report was generated from the Weekly Project Report application.")
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename="Weekly_Project_Report.pdf")
    try:
        with smtplib.SMTP(host, port, timeout=20) as server:
            if use_tls: server.starttls()
            if username and password: server.login(username, password)
            server.send_message(msg)
        return True, f"Report sent successfully to {', '.join(recipients)}."
    except Exception as exc: return False, f"Unable to send report: {exc}"

# =========================================================
# DATA + GLOBAL TOP AREA
# =========================================================
df = pd.DataFrame(st.session_state.notes_db)
filtered_df = pd.DataFrame()
reporting_week = "All Weeks"
reporting_month = "All Months"

st.markdown("""
<div class="app-hero">
    <div class="app-hero-eyebrow">Weekly Project Management</div>
    <div class="app-hero-title">Weekly Project Report</div>
    <div class="app-hero-subtitle">Track progress, priorities, utilization and delivery across your portfolio.</div>
</div>
""", unsafe_allow_html=True)

# Navigation
labels = ["✍️ Enter Notes", "📋 Weekly Summary", "📈 Executive Report"]
cols = st.columns([1,1,1,2]) 
for col, label, m_idx in zip(cols[:3], labels, [1, 2, 3]):
    with col:
        state = "active" if st.session_state.nav_selection == label else "inactive"
        st.markdown(f'<div class="marker-nav{m_idx}-{state}"></div>', unsafe_allow_html=True)
        if st.button(label, key=f"nav_{label}", use_container_width=True):
            st.session_state.nav_selection = label
            st.rerun()

nav_selection = st.session_state.nav_selection

if nav_selection == "📋 Weekly Summary":
    st.markdown('<div class="page-head"><div class="page-title">📋 Weekly Summary</div><div class="page-subtitle">Review project delivery, performance, utilization and current risks.</div></div>', unsafe_allow_html=True)
elif nav_selection == "📈 Executive Report":
    st.markdown('<div class="page-head"><div class="page-title">📈 Executive Report</div><div class="page-subtitle">Executive portfolio view of delivery, ownership, utilization and risks.</div></div>', unsafe_allow_html=True)

if nav_selection != "✍️ Enter Notes" and not df.empty:
    
    # Enable the 3D CSS for Filters dynamically ONLY on Tabs 2 & 3
    st.markdown("""
    <style>
    div[data-testid="stSelectbox"], div[data-testid="stMultiSelect"] {
        background: linear-gradient(to bottom, #ffffff, #f8fafc) !important; border: 1px solid #cbd5e1 !important;
        box-shadow: 0 3px 0 #cbd5e1, 0 4px 6px rgba(0,0,0,0.03) !important; border-radius: 12px !important;
        padding: 8px 10px 14px 10px !important; margin-bottom: 8px !important; transition: all 0.15s ease-out !important;
    }
    div[data-testid="stSelectbox"]:active, div[data-testid="stMultiSelect"]:active {
        transform: translateY(3px) !important; box-shadow: 0 0px 0 transparent, 0 0px 0px rgba(0,0,0,0) !important;
    }
    div[data-testid="stSelectbox"] [data-baseweb="select"] > div, div[data-testid="stMultiSelect"] [data-baseweb="select"] > div {
        background: #ffffff !important; border: 1px solid #cbd5e1 !important; box-shadow: inset 0 2px 4px rgba(0,0,0,0.02) !important;
        border-radius: 8px !important; min-height: 40px !important; transform: none !important;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div style="font-size: 14px; font-weight: bold; color: #16324f; letter-spacing: 1px; text-transform: uppercase; border-bottom: 2px dashed #cbd5e1; padding-bottom: 8px; margin-bottom: 16px; margin-top: 8px;">🔎 REPORT FILTERS</div>', unsafe_allow_html=True)
    f_col1, f_col2, f_col3, f_col4, f_col5, f_col6 = st.columns(6)

    with f_col1:
        filter_acc = st.selectbox("Account", options=["All"] + list(df["Account"].unique()), key="filter_account")
        filtered_df = df if filter_acc == "All" else df[df["Account"] == filter_acc]
    with f_col2:
        filter_team = st.selectbox("Team", options=["All"] + list(filtered_df["Team"].unique()), key="filter_team")
        if filter_team != "All": filtered_df = filtered_df[filtered_df["Team"] == filter_team]
    with f_col3:
        year_options = list(filtered_df["Year"].unique())
        filter_year = st.multiselect("Year", options=year_options, default=year_options, key="filter_year")
        filtered_df = filtered_df[filtered_df["Year"].isin(filter_year)] if filter_year else pd.DataFrame(columns=df.columns)
    with f_col4:
        month_options = list(filtered_df["Month"].unique())
        filter_month = st.multiselect("Month", options=month_options, default=month_options, key="filter_month")
        filtered_df = filtered_df[filtered_df["Month"].isin(filter_month)] if filter_month else pd.DataFrame(columns=df.columns)
    with f_col5:
        week_options = list(filtered_df["Week"].unique())
        filter_week = st.multiselect("Week", options=week_options, default=week_options, key="filter_week")
        filtered_df = filtered_df[filtered_df["Week"].isin(filter_week)] if filter_week else pd.DataFrame(columns=df.columns)
    with f_col6:
        res_options = list(filtered_df["Resource"].unique())
        selected_resources = st.multiselect("Resource(s)", options=res_options, default=res_options, key="filter_resources")
        filtered_df = filtered_df[filtered_df["Resource"].isin(selected_resources)] if selected_resources else pd.DataFrame(columns=df.columns)

    reporting_month = format_badge_text(filter_month, month_options)
    reporting_week = format_badge_text(filter_week, week_options)

    # Action Buttons Row
    if not filtered_df.empty:
        if nav_selection == "📋 Weekly Summary":
            a1, a2, _ = st.columns([1.5, 1.5, 7])
            with a1:
                st.markdown('<div class="marker-send"></div>', unsafe_allow_html=True)
                if st.button("📤 Send Report", key="send_tab2", use_container_width=True):
                    st.session_state["show_send_dialog"] = "weekly"
            with a2:
                if FPDF_AVAILABLE:
                    pdf_bytes = create_pdf_report(filtered_df, f"Weekly Summary: {reporting_month} - {reporting_week}")
                    st.markdown('<div class="marker-pdf"></div>', unsafe_allow_html=True)
                    st.download_button("📄 Download PDF", data=pdf_bytes, file_name="Weekly_Summary.pdf", mime="application/pdf", key="pdf_tab2", use_container_width=True)
        else:
            a1, a2, a3, _ = st.columns([1.5, 1.5, 1.5, 5.5])
            with a1:
                st.markdown('<div class="marker-send"></div>', unsafe_allow_html=True)
                if st.button("📤 Send Report", key="send_tab3", use_container_width=True):
                    st.session_state["show_send_dialog"] = "executive"
            with a2:
                excel_bytes = create_excel_report(filtered_df, f"Executive Report: {reporting_month} - {reporting_week}")
                st.markdown('<div class="marker-excel"></div>', unsafe_allow_html=True)
                st.download_button("📊 Export Excel", data=excel_bytes, file_name="Weekly_Notes_Executive_Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="excel_tab3", use_container_width=True)
            with a3:
                if FPDF_AVAILABLE:
                    pdf_bytes = create_pdf_report(filtered_df, f"Executive Report: {reporting_month} - {reporting_week}")
                    st.markdown('<div class="marker-pdf"></div>', unsafe_allow_html=True)
                    st.download_button("📄 Export PDF", data=pdf_bytes, file_name="Weekly_Notes_Executive_Report.pdf", mime="application/pdf", key="pdf_tab3", use_container_width=True)
else:
    if nav_selection == "✍️ Enter Notes": filtered_df = df

# =========================================================
# SEND REPORT DIALOG
# =========================================================
if st.session_state.get("show_send_dialog"):
    report_kind = st.session_state["show_send_dialog"]
    report_title = f"{'Weekly Summary' if report_kind == 'weekly' else 'Executive Report'}: {reporting_month} - {reporting_week}"
    with st.expander("📤 Send Report", expanded=True):
        st.markdown("**Send the generated PDF report by email**")
        to_value = st.text_input("To", placeholder="name@company.com, another@company.com", key="send_to")
        cc_value = st.text_input("CC (optional)", placeholder="name@company.com", key="send_cc")
        subject_value = st.text_input("Subject", value=report_title, key="send_subject")
        b1, b2 = st.columns([1, 1])
        with b1:
            if st.button("✉️ Send Now", type="primary", use_container_width=True, key="send_now"):
                to_list = [x.strip() for x in to_value.replace(";", ",").split(",") if x.strip()]
                cc_list = [x.strip() for x in cc_value.replace(";", ",").split(",") if x.strip()]
                current_pdf = create_pdf_report(filtered_df, subject_value) if FPDF_AVAILABLE else None
                if current_pdf is None: st.error("PDF generation unavailable.")
                else:
                    ok, message = send_report_email(current_pdf, subject_value, to_list, cc_list)
                    if ok: st.success(message); st.session_state["show_send_dialog"] = None
                    else: st.error(message)
        with b2:
            if st.button("Cancel", use_container_width=True, key="cancel_send"):
                st.session_state["show_send_dialog"] = None

if st.session_state.show_success:
    st.toast(f"✅ {st.session_state.success_message}")
    st.session_state.show_success = False

# =========================================================
# VIEW 1: ENTER NOTES
# =========================================================
if nav_selection == "✍️ Enter Notes":
    st.markdown('<div class="page-head"><div class="page-title">✍️ Enter Weekly Notes</div><div class="page-subtitle">Capture your team\'s weekly progress, priorities, utilization and risks.</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-header">01 — Core Assignment</div>', unsafe_allow_html=True)
    with st.container(border=True):
        col_acc, col_team = st.columns(2)
        with col_acc:
            account_options = ["➕ Add New Account..."] + list(st.session_state.accounts_db.keys())
            selected_account = st.selectbox("Account", options=account_options)
            final_account = st.text_input("Enter New Account Name:") if selected_account == "➕ Add New Account..." else selected_account
        with col_team:
            if final_account and final_account != "➕ Add New Account...":
                team_list = st.session_state.accounts_db.get(final_account, [])
                team_options = ["➕ Add New Team..."] + team_list
                selected_team = st.selectbox("Team", options=team_options)
                final_team = st.text_input("Enter New Team Name:") if selected_team == "➕ Add New Team..." else selected_team
            else:
                st.info("Select or enter an Account first to proceed.")
                final_team = None

    if final_account and final_team:
        st.markdown('<div class="section-header">02 — Timeframe & Ownership</div>', unsafe_allow_html=True)
        with st.container(border=True):
            t_col1, t_col2, t_col3, t_col4 = st.columns(4)
            with t_col1: year = st.selectbox("Year", options=[datetime.now().year, datetime.now().year+1, datetime.now().year+2])
            with t_col2: month = st.selectbox("Month", options=["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"])
            with t_col3: week = st.selectbox("Week", options=["Week 1", "Week 2", "Week 3", "Week 4", "Week 5"])
            with t_col4: resource = st.text_input("Resource Name")

        st.markdown('<div class="section-header">03 — Weekly Accomplishments</div>', unsafe_allow_html=True)
        projects_to_save = []

        for i in range(st.session_state.project_count):
            with st.container(border=True):
                st.markdown(f'<div class="card-title">Project Entry #{i + 1}</div>', unsafe_allow_html=True)
                st.markdown('<hr>', unsafe_allow_html=True)
                r1c1, r1c2 = st.columns(2)
                project = r1c1.text_input("Project / Dashboard Name", placeholder="e.g., Q3 Sales Analysis", key=f"proj_{i}")
                business_owner = r1c2.text_input("Business Owner", placeholder="e.g., Marketing Team", key=f"owner_{i}")
                
                r2c1, r2c2, r2c3 = st.columns(3)
                start_date = r2c1.date_input("Start Date", value=None, key=f"start_{i}")
                initial_delivery = r2c2.date_input("Initial Delivery Date", value=None, key=f"init_{i}")
                updated_delivery = r2c3.date_input("Updated Expected Delivery", value=None, key=f"upd_{i}")
                
                r3c1, r3c2, r3c3, r3c4 = st.columns(4)
                wt_options = ["Development", "Enhancement", "Adhoc", "Migration", "Other"]
                wt_selection = r3c1.selectbox("Work Type", wt_options, key=f"wt_sel_{i}")
                if wt_selection == "Other": work_type = r3c1.text_input("Specify Work Type", placeholder="Enter work type...", key=f"work_{i}")
                else: work_type = wt_selection
                
                status = r3c2.selectbox("Status", ["On Track", "At Risk", "Blocked", "Completed", "Not Started"], key=f"status_{i}")
                completion_pct = r3c3.number_input("Completion %", min_value=0, max_value=100, step=5, key=f"comp_{i}")
                utilization_pct = r3c4.number_input("Utilization %", min_value=0, max_value=100, step=5, key=f"util_{i}")
                
                c_text1, c_text2 = st.columns(2)
                this_week = c_text1.text_area("This Week Delivered", placeholder="What was accomplished this week?", height=85, key=f"this_{i}")
                next_week = c_text2.text_area("Next Week Priority", placeholder="What is the focus for next week?", height=85, key=f"next_{i}")
                blocker = st.text_area("Blocker / Risks", placeholder="Detail any risks or blockers here. (Leave blank if none)", height=68, key=f"block_{i}")
                
                projects_to_save.append({
                    "id": str(uuid.uuid4()), "Project / Dashboard": project, "Business Owner": business_owner,
                    "Start date": str(start_date) if start_date else "N/A", "Initial Delivery date": str(initial_delivery) if initial_delivery else "N/A",
                    "Updated Expected Delivery date": str(updated_delivery) if updated_delivery else "N/A", "Work Type": work_type,
                    "Completion %": f"{completion_pct}%", "Utilization %": f"{utilization_pct}%", "Status": status,
                    "This Week Delivered": this_week, "Next Week Priority": next_week, "Blocker": blocker
                })

        col_add, col_rem, _ = st.columns([2, 2, 6])
        with col_add:
            if st.button("➕ Add Project", use_container_width=True):
                st.session_state.project_count += 1
                st.rerun()
        with col_rem:
            if st.session_state.project_count > 1:
                if st.button("🗑️ Remove Last", use_container_width=True):
                    st.session_state.project_count -= 1
                    st.rerun()

        st.divider()
        sub_col1, sub_col2, sub_col3 = st.columns([2, 3, 2])
        with sub_col2:
            st.markdown('<div class="marker-nav1-active"></div>', unsafe_allow_html=True)
            if st.button("💾 Save to Google Sheets", type="primary", use_container_width=True):
                if not resource: st.error("Please enter your Resource Name in Step 2 before saving.")
                else:
                    missing_names = [i + 1 for i, p in enumerate(projects_to_save) if not p["Project / Dashboard"]]
                    if missing_names: st.error(f"Please enter a Project Name for Project(s): {', '.join(map(str, missing_names))}")
                    else:
                        with st.spinner("Saving data to Google Sheets..."):
                            final_entries = []
                            for proj_data in projects_to_save:
                                full_entry = {"Account": final_account, "Team": final_team, "Year": year, "Month": month, "Week": week, "Resource": resource}
                                full_entry.update(proj_data)
                                final_entries.append(full_entry)
                            
                            save_new_notes_to_gsheets(final_entries)
                            st.session_state.data_synced = False
                            st.session_state.success_message = f"Notes successfully saved to Google Sheets."
                            st.session_state.show_success = True
                            st.session_state.project_count = 1
                            st.rerun()

# =========================================================
# VIEW 2: WEEKLY SUMMARY
# =========================================================
elif nav_selection == "📋 Weekly Summary":
    if df.empty or filtered_df.empty:
        st.info("No data available to generate the summary. Please adjust filters or enter data.")
    else:
        total_projects = len(filtered_df)
        total_resources = filtered_df["Resource"].nunique()
        delivered_count = len(filtered_df[filtered_df["This Week Delivered"].fillna("").str.strip() != ""])
        avg_comp = int((filtered_df["Completion %"].apply(lambda x: get_pct_decimal(x) * 100)).mean()) if not filtered_df.empty else 0
        avg_util = int((filtered_df["Utilization %"].apply(lambda x: get_pct_decimal(x) * 100)).mean()) if not filtered_df.empty else 0
        risk_items = filtered_df[(filtered_df["Status"].isin(["At Risk", "Blocked"])) | (filtered_df["Blocker"].fillna("").str.strip() != "")]
        risk_count = len(risk_items)

        st.markdown('<div class="section-header">Executive Summary</div>', unsafe_allow_html=True)
        st.markdown(f"""
        <div class="summary-box">
            This week's portfolio includes <strong>{total_projects}</strong> active work item(s), with <strong>{delivered_count}</strong> reporting a delivery update. 
            Average completion is <strong>{avg_comp}%</strong> and average utilization is <strong>{avg_util}%</strong>. 
            There are <strong>{risk_count}</strong> item(s) requiring risk or blocker attention.
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="section-header">Portfolio Snapshot</div>', unsafe_allow_html=True)
        s1, s2, s3, s4, s5 = st.columns(5)
        with s1: st.markdown(render_metric_card("RESOURCES", total_resources), unsafe_allow_html=True)
        with s2: st.markdown(render_metric_card("PROJECTS", total_projects), unsafe_allow_html=True)
        with s3: st.markdown(render_metric_card("DELIVERED", delivered_count), unsafe_allow_html=True)
        with s4: st.markdown(render_metric_card("COMPLETION", f"{avg_comp}%"), unsafe_allow_html=True)
        with s5: st.markdown(render_metric_card("UTILIZATION", f"{avg_util}%"), unsafe_allow_html=True)

        st.markdown('<div class="section-header">Project / Dashboard Status</div>', unsafe_allow_html=True)
        html_lines = [
            '<table class="custom-table">',
            '<thead><tr><th>Resource</th><th>Business Owner</th><th>Project / Dashboard</th><th>Expected Delivery</th><th>Completion</th><th>Utilization</th><th>Status</th></tr></thead>',
            '<tbody>'
        ]
        for _, row in filtered_df.iterrows():
            raw_status = row['Status']
            if raw_status == "On Track": pill = "status-on-track"
            elif raw_status == "At Risk": pill = "status-in-progress" 
            elif raw_status == "Blocked": pill = "status-blocked"
            elif raw_status == "Completed": pill = "status-completed"
            else: pill = "status-not-started"
            display_status = "In Progress" if raw_status == "At Risk" else raw_status
            row_html = f"<tr><td>{row['Resource']}</td><td>{row['Business Owner']}</td><td><span class='proj-name'>{row['Project / Dashboard']}</span><span class='work-type'>{row['Work Type']}</span></td><td>{row['Updated Expected Delivery date']}</td><td>{row['Completion %']}</td><td>{row['Utilization %']}</td><td><span class='status-pill {pill}'>{display_status}</span></td></tr>"
            html_lines.append(row_html)
        html_lines.append("</tbody></table>")
        st.markdown("".join(html_lines), unsafe_allow_html=True)

        st.markdown('<div class="section-header">Blockers & Risks</div>', unsafe_allow_html=True)
        if risk_count == 0: st.success("No active blockers or risks reported this week! 🎉")
        else:
            for _, row in risk_items.iterrows():
                blocker_text = row['Blocker'] if row['Blocker'] else f"Status marked as: {row['Status']}"
                st.markdown(f'<div class="risk-box"><div class="risk-title">⚠️ {row["Project / Dashboard"]}</div><div class="risk-desc">{blocker_text}</div></div>', unsafe_allow_html=True)

# =========================================================
# VIEW 3: EXECUTIVE REPORT
# =========================================================
elif nav_selection == "📈 Executive Report":
    if df.empty or filtered_df.empty:
        st.info("No data available. Go to the 'Enter Notes' tab to log your first update.")
    else:
        st.markdown('<div class="section-header">Portfolio Management View</div>', unsafe_allow_html=True)
        for account_name in filtered_df["Account"].unique():
            st.markdown(f"<div class='acc-header'>🏢 {account_name}</div>", unsafe_allow_html=True)
            account_data = filtered_df[filtered_df["Account"] == account_name]
            for team_name in account_data["Team"].unique():
                st.markdown(f"<div class='team-header'>{team_name}</div>", unsafe_allow_html=True)
                team_data = account_data[account_data["Team"] == team_name]
                for resource_name in team_data["Resource"].unique():
                    resource_data = team_data[team_data["Resource"] == resource_name]
                    with st.expander(f"👤 {resource_name} ({len(resource_data)} updates)", expanded=True):
                        for index, row in resource_data.iterrows():
                            rec_id = row.get('id')
                            if st.session_state.edit_id == rec_id:
                                with st.container(border=True):
                                    st.markdown(f"<div class='edit-banner'>✏️ Editing: {row['Project / Dashboard']}</div>", unsafe_allow_html=True)
                                    e_r1c1, e_r1c2, e_r1c3 = st.columns(3)
                                    e_proj = e_r1c1.text_input("Project Name", value=row.get('Project / Dashboard', ''), key=f"ep_{rec_id}")
                                    e_bo = e_r1c2.text_input("Business Owner", value=row.get('Business Owner', ''), key=f"ebo_{rec_id}")
                                    wt_options = ["Development", "Enhancement", "Adhoc", "Migration", "Other"]
                                    existing_wt = row.get('Work Type', '')
                                    wt_idx = wt_options.index(existing_wt) if existing_wt in wt_options else 4
                                    e_wt_sel = e_r1c3.selectbox("Work Type", wt_options, index=wt_idx, key=f"ewtsel_{rec_id}")
                                    if e_wt_sel == "Other": e_wt = e_r1c3.text_input("Specify Work Type", value=existing_wt if existing_wt not in wt_options else "", key=f"ewt_{rec_id}")
                                    else: e_wt = e_wt_sel
                                        
                                    e_r2c1, e_r2c2, e_r2c3 = st.columns(3)
                                    e_start = e_r2c1.date_input("Start Date", value=parse_date(row.get('Start date')), key=f"esd_{rec_id}")
                                    e_init = e_r2c2.date_input("Initial Delivery", value=parse_date(row.get('Initial Delivery date')), key=f"eid_{rec_id}")
                                    e_upd = e_r2c3.date_input("Updated Delivery", value=parse_date(row.get('Updated Expected Delivery date')), key=f"eud_{rec_id}")
                                    e_r3c1, e_r3c2, e_r3c3 = st.columns(3)
                                    status_opts = ["On Track", "At Risk", "Blocked", "Completed", "Not Started"]
                                    curr_status = row.get('Status', 'On Track')
                                    if curr_status not in status_opts: curr_status = "On Track"
                                    e_status = e_r3c1.selectbox("Status", status_opts, index=status_opts.index(curr_status), key=f"estat_{rec_id}")
                                    e_comp = e_r3c2.number_input("Completion %", min_value=0, max_value=100, step=5, value=int(str(row.get('Completion %', '0')).replace('%','')), key=f"ecomp_{rec_id}")
                                    e_util = e_r3c3.number_input("Utilization %", min_value=0, max_value=100, step=5, value=int(str(row.get('Utilization %', '0')).replace('%','')), key=f"eutil_{rec_id}")
                                    e_tw = st.text_area("This Week Delivered", value=row.get('This Week Delivered', ''), height=85, key=f"etw_{rec_id}")
                                    e_nw = st.text_area("Next Week Priority", value=row.get('Next Week Priority', ''), height=85, key=f"enw_{rec_id}")
                                    e_block = st.text_area("Blocker / Risks", value=row.get('Blocker', ''), height=68, key=f"eb_{rec_id}")
                                    st.divider()
                                    e_btn_c1, e_btn_c2, _ = st.columns([2, 2, 6])
                                    with e_btn_c1:
                                        if st.button("💾 Save Changes", type="primary", key=f"esave_{rec_id}", use_container_width=True):
                                            with st.spinner("Updating Google Sheets..."):
                                                updated_data = {
                                                    "id": rec_id, "Account": account_name, "Team": team_name, "Year": row.get("Year"), "Month": row.get("Month"), "Week": row.get("Week"), "Resource": resource_name,
                                                    "Project / Dashboard": e_proj, "Business Owner": e_bo, "Work Type": e_wt,
                                                    "Start date": str(e_start) if e_start else "N/A", "Initial Delivery date": str(e_init) if e_init else "N/A", "Updated Expected Delivery date": str(e_upd) if e_upd else "N/A",
                                                    "Status": e_status, "Completion %": f"{e_comp}%", "Utilization %": f"{e_util}%",
                                                    "This Week Delivered": e_tw, "Next Week Priority": e_nw, "Blocker": e_block
                                                }
                                                update_existing_note_in_gsheets(updated_data)
                                                st.session_state.data_synced = False
                                            st.session_state.edit_id = None; st.session_state.show_success = True; st.session_state.success_message = "Cloud update saved successfully."; st.rerun()
                                    with e_btn_c2:
                                        if st.button("Cancel", key=f"ecancel_{rec_id}", use_container_width=True):
                                            st.session_state.edit_id = None; st.rerun()
                            else:
                                with st.container(border=True):
                                    h_col1, h_col2 = st.columns([9, 1])
                                    with h_col1:
                                        st.markdown(f"<div class='card-title'> {row['Project / Dashboard']}</div>", unsafe_allow_html=True)
                                        st.markdown(f"<span style='font-size:14px; color:#667085;'><b>Owner:</b> {row['Business Owner']} &nbsp;|&nbsp; <b>Type:</b> {row['Work Type']}</span>", unsafe_allow_html=True)
                                    with h_col2:
                                        if st.button("✏️ Edit", key=f"edit_btn_{rec_id}", use_container_width=True):
                                            st.session_state.edit_id = rec_id; st.rerun()
                                    c1, c2, c3 = st.columns([2, 2, 1])
                                    with c1:
                                        st.markdown("<div style='font-size:13px; font-weight:800; color:#344054;'>THIS WEEK DELIVERED</div>", unsafe_allow_html=True)
                                        st.markdown(f"<div style='font-size:15px; color:#344054;'>{row['This Week Delivered'] if row['This Week Delivered'] else '<i>None reported.</i>'}</div>", unsafe_allow_html=True)
                                        st.markdown(f"<div style='font-size:13px; margin-top:7px; color:#667085;'>Start: <b>{row['Start date']}</b></div>", unsafe_allow_html=True)
                                    with c2:
                                        st.markdown("<div style='font-size:13px; font-weight:800; color:#344054;'>NEXT WEEK PRIORITY</div>", unsafe_allow_html=True)
                                        st.markdown(f"<div style='font-size:15px; color:#344054;'>{row['Next Week Priority'] if row['Next Week Priority'] else '<i>None reported.</i>'}</div>", unsafe_allow_html=True)
                                        st.markdown(f"<div style='font-size:13px; margin-top:7px; color:#667085;'>Expected Delivery: <b>{row['Updated Expected Delivery date']}</b></div>", unsafe_allow_html=True)
                                    with c3:
                                        st.markdown("<div style='font-size:13px; font-weight:800; color:#344054;'>METRICS</div>", unsafe_allow_html=True)
                                        st.markdown(f"<div style='font-size:14px; display:flex; justify-content:space-between;'><span>Completion</span><b>{row['Completion %']}</b></div>", unsafe_allow_html=True)
                                        st.progress(get_pct_decimal(row['Completion %']))
                                        st.markdown(f"<div style='font-size:14px; display:flex; justify-content:space-between; margin-top:3px;'><span>Utilization</span><b>{row['Utilization %']}</b></div>", unsafe_allow_html=True)
                                        st.progress(get_pct_decimal(row['Utilization %']))
                                    if row["Blocker"] and str(row["Blocker"]).strip().lower() not in ["none", "na", "n/a", ""]:
                                        st.markdown(f"<div class='alert-card alert-danger' style='margin-top:14px; margin-bottom:0;'><div class='alert-title'>⚠ Blocker</div><div class='alert-desc'>{row['Blocker']}</div></div>", unsafe_allow_html=True)
