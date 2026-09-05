import streamlit as st
import pandas as pd
from datetime import datetime, date, timedelta
import io
import uuid
import os
import html
import re
import smtplib
import hashlib
import hmac
import secrets
from email.message import EmailMessage
from email.utils import getaddresses
from typing import Any, Dict, List, Optional, Tuple
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

# Password verification supports the existing plain-text secrets for backward compatibility
# and also supports production PBKDF2 hashes in the form:
# pbkdf2_sha256$iterations$salt_hex$digest_hex
def hash_password(password, iterations=310000):
    salt = secrets.token_bytes(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, iterations)
    return f"pbkdf2_sha256${iterations}${salt.hex()}${digest.hex()}"


def verify_configured_password(entered_password, configured_password):
    configured = str(configured_password)
    if configured.startswith("pbkdf2_sha256$"):
        try:
            _, iterations, salt_hex, digest_hex = configured.split("$", 3)
            expected = bytes.fromhex(digest_hex)
            actual = hashlib.pbkdf2_hmac("sha256", entered_password.encode("utf-8"), bytes.fromhex(salt_hex), int(iterations))
            return hmac.compare_digest(actual, expected)
        except (ValueError, TypeError):
            return False
    # Backward-compatible comparison for existing deployments.
    return hmac.compare_digest(str(entered_password), configured)


# =========================================================
# SECURITY & ROLE-BASED LOGIN SCREEN
# =========================================================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.current_email = ""
    st.session_state.current_role = ""
    st.session_state.current_scope = "" 
    st.session_state.current_acc_scope = ""
    st.session_state.login_attempts = 0

if not st.session_state.authenticated:
    st.markdown('<div style="max-width: 450px; margin: 80px auto; padding: 30px; background: white; border-radius: 12px; border: 1px solid #dfe5ec; box-shadow: 0 4px 12px rgba(0,0,0,0.1);">', unsafe_allow_html=True)
    st.markdown("<h2 style='color: #16324f; text-align:center; margin-bottom: 5px;'>🔒 App Locked</h2>", unsafe_allow_html=True)
    st.markdown("<p style='color: #64748b; margin-bottom: 20px; text-align:center;'>Log in to access the Weekly Project Report.</p>", unsafe_allow_html=True)
    
    role = st.selectbox("Select Your Role", ["Team Member", "Team Lead", "Account Manager", "Admin"])
    user_email = st.text_input("Your Company Email", placeholder="name@company.com")
    
    user_scope = ""
    user_account_scope = ""
    
    if role == "Team Lead":
        user_account_scope = st.text_input("Account Name", placeholder="e.g. Acme Corp")
        user_scope = st.text_input("Your Team(s)", placeholder="e.g. Alpha Team, Beta Team (comma-separated)")
    elif role == "Account Manager":
        user_scope = st.text_input("Your Account(s)", placeholder="e.g. Acme Corp, Globex (comma-separated)")
        
    entered_pwd = st.text_input("Password", type="password", placeholder="Enter Role Password...")
    
    if st.button("Unlock App", type="primary", use_container_width=True):
        if st.session_state.get("login_attempts", 0) >= 5:
            st.error("Too many unsuccessful login attempts in this session. Refresh the page to retry.")
            st.stop()
        if not user_email.strip():
            st.error("Please enter your company email.")
        elif role == "Team Lead" and (not user_scope.strip() or not user_account_scope.strip()):
            st.error("Please enter both your Account and Team.")
        elif role == "Account Manager" and not user_scope.strip():
            st.error("Please enter your Account(s).")
        else:
            valid_pass = False
            password_keys = {"Team Member": "APP_PASSWORD", "Team Lead": "LEAD_PASSWORD", "Account Manager": "ACCOUNT_PASSWORD", "Admin": "ADMIN_PASSWORD"}
            password_key = password_keys[role]
            configured_password = st.secrets.get(password_key, os.getenv(password_key, ""))
            if not configured_password:
                st.error(f"Login is not configured for {role}. Please configure {password_key} in Streamlit secrets/environment variables.")
            elif verify_configured_password(entered_pwd, configured_password):
                valid_pass = True
            
            if valid_pass:
                st.session_state.authenticated = True
                st.session_state.current_email = user_email.strip().lower()
                st.session_state.current_role = role
                st.session_state.current_scope = user_scope.strip()
                st.session_state.current_acc_scope = user_account_scope.strip()
                st.session_state.login_attempts = 0
                st.rerun()
            else:
                st.session_state.login_attempts = st.session_state.get("login_attempts", 0) + 1
                remaining = max(0, 5 - st.session_state.login_attempts)
                msg = "Incorrect password."
                if remaining:
                    msg += f" {remaining} attempt(s) remaining."
                else:
                    msg += " Refresh the page before trying again."
                st.error(msg)
    st.markdown('</div>', unsafe_allow_html=True)
    st.stop()

# =========================================================
# GOOGLE SHEETS BACKEND CONFIGURATION
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
div.element-container:has(.marker-nav4-inactive) + div.element-container button { background: linear-gradient(to bottom, #f8fafc, #fff1f2) !important; border-color: #fecdd3 !important; box-shadow: 0 3px 0 #fecdd3, 0 4px 6px rgba(0,0,0,0.05) !important; color: #9f1239 !important; }
div.element-container:has(.marker-nav4-active) + div.element-container button { background: linear-gradient(to bottom, #ef4444, #b91c1c) !important; border-color: #7f1d1d !important; box-shadow: 0 3px 0 #7f1d1d, 0 4px 8px rgba(239,68,68,0.3) !important; color: #ffffff !important; }
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
.action-critical { background:#fef2f2; border:1px solid #fecaca; border-left:5px solid #dc2626; padding:12px 14px; border-radius:10px; margin:8px 0; }
.action-attention { background:#fffbeb; border:1px solid #fde68a; border-left:5px solid #d97706; padding:12px 14px; border-radius:10px; margin:8px 0; }
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
if 'last_refresh_at' not in st.session_state: st.session_state.last_refresh_at = datetime.now()
if 'quick_filter' not in st.session_state: st.session_state.quick_filter = "All"
if 'project_search' not in st.session_state: st.session_state.project_search = ""
if 'executive_compact' not in st.session_state: st.session_state.executive_compact = False
if 'copy_source_loaded' not in st.session_state: st.session_state.copy_source_loaded = False

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
    if not selected_list: return "All"
    if len(selected_list) == len(all_options): return "All"
    if len(selected_list) <= 2: return ", ".join(map(str, selected_list))
    return f"{len(selected_list)} selected"

def parse_date(date_str):
    if not date_str or str(date_str).strip() in ["N/A", "None", ""]: return None
    try: return datetime.strptime(str(date_str), "%Y-%m-%d").date()
    except Exception: return None

def pct_value(value):
    """Return percentage as a numeric 0-100 value without raising errors."""
    try:
        return max(0.0, min(100.0, float(str(value).replace("%", "").strip())))
    except Exception:
        return 0.0

def normalize_text(value):
    return str(value or "").strip().lower()

def is_blank(value):
    return normalize_text(value) in {"", "none", "na", "n/a", "nan"}

def row_project_key(row):
    """Stable project identity for week-over-week comparisons."""
    return "|".join([
        normalize_text(row.get("Account")),
        normalize_text(row.get("Team")),
        normalize_text(row.get("Resource")),
        normalize_text(row.get("Project / Dashboard")),
    ])

def date_from_row(row):
    return parse_date(row.get("Updated Expected Delivery date"))

def delivery_state(row, as_of=None):
    as_of = as_of or date.today()
    status = normalize_text(row.get("Status"))
    if status == "completed":
        return "Delivered"
    due = date_from_row(row)
    if not due:
        return "No Date"
    delta = (due - as_of).days
    if delta < 0:
        return f"Overdue {abs(delta)}d"
    if delta <= 7:
        return f"Due in {delta}d"
    return "Upcoming"

def data_quality_flags(row):
    flags = []
    status = normalize_text(row.get("Status"))
    comp = pct_value(row.get("Completion %"))
    due = date_from_row(row)
    if status == "completed" and comp < 100:
        flags.append("Completed but completion is below 100%")
    if comp >= 100 and status != "completed":
        flags.append("Completion is 100% but status is not Completed")
    if status in {"at risk", "blocked"} and is_blank(row.get("Blocker")):
        flags.append("Risk/blocked status without a blocker explanation")
    if due and status != "completed" and due < date.today():
        flags.append("Expected delivery date is overdue")
    if is_blank(row.get("This Week Delivered")):
        flags.append("No weekly delivery update reported")
    return flags

def classify_attention(row, as_of=None):
    """Return a priority label and reasons. Existing statuses remain authoritative."""
    as_of = as_of or date.today()
    status = normalize_text(row.get("Status"))
    reasons = []
    priority = "Normal"

    if status == "blocked":
        priority = "Critical"
        reasons.append("Blocked")
    elif status == "at risk":
        priority = "Critical"
        reasons.append("At Risk")
    elif status == "on hold":
        priority = "Attention"
        reasons.append("On Hold")

    due = date_from_row(row)
    if due and status != "completed":
        days = (due - as_of).days
        if days < 0:
            priority = "Critical"
            reasons.append(f"Overdue by {abs(days)} day(s)")
        elif days <= 7:
            if priority == "Normal":
                priority = "Attention"
            reasons.append(f"Due in {days} day(s)")

    dq = data_quality_flags(row)
    if dq:
        if priority == "Normal":
            priority = "Attention"
        reasons.extend(dq)

    return priority, reasons

def get_latest_prior_week(df, selected_year=None, selected_month=None, selected_week=None):
    """Find the immediately preceding reported week using the app's Year/Month/Week fields."""
    if df.empty or not {"Year", "Month", "Week"}.issubset(df.columns):
        return pd.DataFrame()

    work = df.copy()
    month_order = {m: i for i, m in enumerate([
        "January","February","March","April","May","June",
        "July","August","September","October","November","December"
    ], 1)}
    work["_year"] = pd.to_numeric(work["Year"], errors="coerce")
    work["_month_num"] = work["Month"].map(month_order)
    work["_week_num"] = pd.to_numeric(
        work["Week"].astype(str).str.extract(r"(\d+)")[0], errors="coerce"
    )
    work["_period"] = work["_year"] * 100 + work["_month_num"] * 10 + work["_week_num"]

    if selected_year is None or selected_month is None or selected_week is None:
        return pd.DataFrame()

    current_month_num = month_order.get(str(selected_month))
    current_week_num_match = re.search(r"(\d+)", str(selected_week))
    current_week_num = int(current_week_num_match.group(1)) if current_week_num_match else None
    if current_month_num is None or current_week_num is None:
        return pd.DataFrame()

    current_period = int(selected_year) * 100 + current_month_num * 10 + current_week_num
    prior = work[work["_period"] < current_period]
    if prior.empty:
        return pd.DataFrame()
    prior_period = prior["_period"].max()
    return prior[prior["_period"] == prior_period].drop(columns=["_year","_month_num","_week_num","_period"], errors="ignore")

def compare_current_to_prior(current_df, prior_df):
    """Create project-level movement signals without changing source data."""
    if current_df.empty or prior_df.empty:
        return pd.DataFrame()
    prior_map = {}
    for _, row in prior_df.iterrows():
        prior_map[row_project_key(row)] = row

    records = []
    for _, row in current_df.iterrows():
        key = row_project_key(row)
        prev = prior_map.get(key)
        if prev is None:
            movement = "New"
            comp_delta = None
            util_delta = None
            status_change = "New this period"
        else:
            comp_delta = pct_value(row.get("Completion %")) - pct_value(prev.get("Completion %"))
            util_delta = pct_value(row.get("Utilization %")) - pct_value(prev.get("Utilization %"))
            prev_status = str(prev.get("Status", ""))
            curr_status = str(row.get("Status", ""))
            status_change = f"{prev_status} → {curr_status}" if prev_status != curr_status else "No status change"
            if comp_delta >= 10:
                movement = "Improved"
            elif comp_delta <= -10:
                movement = "Declined"
            else:
                movement = "Stable"
        records.append({
            "Project / Dashboard": row.get("Project / Dashboard", ""),
            "Resource": row.get("Resource", ""),
            "Movement": movement,
            "Completion Δ": comp_delta,
            "Utilization Δ": util_delta,
            "Status Change": status_change,
        })
    return pd.DataFrame(records)

def html_escape(value):
    return html.escape("" if value is None else str(value))

def build_action_items(dataframe):
    items = []
    if dataframe.empty:
        return items
    for _, row in dataframe.iterrows():
        priority, reasons = classify_attention(row)
        if priority != "Normal":
            items.append({
                "Priority": priority,
                "Project / Dashboard": row.get("Project / Dashboard", ""),
                "Account": row.get("Account", ""),
                "Team": row.get("Team", ""),
                "Resource": row.get("Resource", ""),
                "Status": row.get("Status", ""),
                "Delivery": delivery_state(row),
                "Reason": "; ".join(dict.fromkeys(reasons)),
            })
    rank = {"Critical": 0, "Attention": 1, "Normal": 2}
    items.sort(key=lambda x: (rank.get(x["Priority"], 9), x["Project / Dashboard"]))
    return items

def refresh_cloud_data(show_message=True):
    """Explicitly refresh Google Sheets data while preserving the existing data model."""
    with st.spinner("Refreshing data from Google Sheets..."):
        latest = load_data_from_gsheets()
    if latest:
        st.session_state.notes_db = latest
        st.session_state.data_synced = True
        if show_message:
            st.success(f"Data refreshed successfully — {len(latest)} record(s) loaded.")
    else:
        # Keep existing data if the refresh failed/returned empty.
        st.warning("Refresh returned no records. Existing session data was retained.")
    st.session_state.accounts_db = {}
    if st.session_state.notes_db:
        live_df = pd.DataFrame(st.session_state.notes_db)
        if "Account" in live_df.columns and "Team" in live_df.columns:
            for acc in live_df["Account"].dropna().unique():
                st.session_state.accounts_db[acc] = live_df[live_df["Account"] == acc]["Team"].dropna().unique().tolist()

# =========================================================
# PRODUCTION REPORT ENGINE 2.0
# One report model -> HTML email + PDF + Excel
# Existing application/data-entry capabilities are preserved.
# =========================================================

def _report_period_label(dataframe):
    if dataframe is None or dataframe.empty:
        return "Current Selection"
    vals = []
    for col in ["Year", "Month", "Week"]:
        if col in dataframe.columns:
            unique = [str(v) for v in dataframe[col].dropna().unique() if not is_blank(v)]
            if len(unique) == 1:
                vals.append(unique[0])
    return " · ".join(vals) if vals else "Current Selection"


def _report_scope_text(dataframe):
    if dataframe is None or dataframe.empty:
        return "No records in the current selection"
    accounts = dataframe["Account"].nunique() if "Account" in dataframe.columns else 0
    teams = dataframe["Team"].nunique() if "Team" in dataframe.columns else 0
    return f"{len(dataframe)} update(s) · {accounts} account(s) · {teams} team(s)"


def _risk_dataframe(dataframe):
    if dataframe is None or dataframe.empty:
        return pd.DataFrame()
    status = dataframe.get("Status", pd.Series(index=dataframe.index, dtype=object)).fillna("").astype(str).str.strip()
    blocker = dataframe.get("Blocker", pd.Series(index=dataframe.index, dtype=object)).fillna("").astype(str).str.strip()
    return dataframe.loc[status.isin(["At Risk", "Blocked", "On Hold"]) | blocker.ne("")].copy()


def _delivery_outlook(dataframe, as_of=None):
    as_of = as_of or date.today()
    result = {"Overdue": 0, "Next 7 Days": 0, "8-14 Days": 0, "15+ Days": 0, "Completed": 0, "No Date": 0}
    if dataframe is None or dataframe.empty:
        return result
    for _, row in dataframe.iterrows():
        status = normalize_text(row.get("Status"))
        if status == "completed":
            result["Completed"] += 1
            continue
        due = date_from_row(row)
        if not due:
            result["No Date"] += 1
            continue
        days = (due - as_of).days
        if days < 0: result["Overdue"] += 1
        elif days <= 7: result["Next 7 Days"] += 1
        elif days <= 14: result["8-14 Days"] += 1
        else: result["15+ Days"] += 1
    return result


def _avg_completion(dataframe):
    if dataframe is None or dataframe.empty or "Completion %" not in dataframe.columns:
        return 0
    return int(round(dataframe["Completion %"].apply(pct_value).mean()))


def _avg_utilization(dataframe):
    if dataframe is None or dataframe.empty or "Utilization %" not in dataframe.columns:
        return 0
    work = dataframe.assign(_util=dataframe["Utilization %"].apply(pct_value))
    if "Resource" in work.columns:
        values = work.groupby("Resource")["_util"].mean()
        return int(round(values.mean())) if not values.empty else 0
    return int(round(work["_util"].mean()))


def _portfolio_health(dataframe):
    if dataframe is None or dataframe.empty:
        return 0
    total = max(len(dataframe), 1)
    risks = len(_risk_dataframe(dataframe))
    overdue = _delivery_outlook(dataframe)["Overdue"]
    score = 70 + (_avg_completion(dataframe) * 0.25) - (risks / total * 18) - (overdue / total * 22)
    return int(max(0, min(100, round(score))))


def _health_label(score):
    return "Healthy" if score >= 80 else "Attention" if score >= 60 else "Critical"


def _leadership_takeaway(dataframe, movement_df=None):
    if dataframe is None or dataframe.empty:
        return "No records are available for the selected reporting period."
    total = len(dataframe)
    resources = dataframe["Resource"].nunique() if "Resource" in dataframe.columns else 0
    risks = len(_risk_dataframe(dataframe))
    out = _delivery_outlook(dataframe)
    health = _portfolio_health(dataframe)
    text = f"Portfolio health is {health}/100 ({_health_label(health).lower()}) across {total} update(s) and {resources} resource(s), with average initiative completion at {_avg_completion(dataframe)}%."
    text += f" {risks} item(s) require risk or blocker attention." if risks else " No active risk/blocker items are currently flagged."
    if out["Overdue"] or out["Next 7 Days"]:
        text += f" Delivery outlook includes {out['Overdue']} overdue and {out['Next 7 Days']} item(s) due within 7 days."
    if movement_df is not None and not movement_df.empty:
        improved = int((movement_df["Movement"] == "Improved").sum())
        declined = int((movement_df["Movement"] == "Declined").sum())
        if improved or declined:
            text += f" WoW movement shows {improved} improved and {declined} declined item(s)."
    return text


def _build_report_model(dataframe, report_title):
    data = dataframe.copy() if dataframe is not None else pd.DataFrame()
    prior_df = pd.DataFrame()
    if not data.empty and {"Year", "Month", "Week"}.issubset(data.columns):
        ys, ms, ws = [data[c].dropna().unique().tolist() for c in ["Year", "Month", "Week"]]
        if len(ys) == len(ms) == len(ws) == 1:
            prior_df = get_latest_prior_week(df, ys[0], ms[0], ws[0])
    movement_df = compare_current_to_prior(data, prior_df)
    actions = build_action_items(data)
    risks_df = _risk_dataframe(data)
    updates = int(data["This Week Delivered"].fillna("").astype(str).str.strip().ne("").sum()) if "This Week Delivered" in data.columns else 0
    kpis = {
        "projects": len(data),
        "resources": int(data["Resource"].nunique()) if not data.empty and "Resource" in data.columns else 0,
        "updates": updates,
        "completion": _avg_completion(data),
        "utilization": _avg_utilization(data),
        "health": _portfolio_health(data),
        "critical": sum(x["Priority"] == "Critical" for x in actions),
        "attention": sum(x["Priority"] == "Attention" for x in actions),
        "risks": len(risks_df),
    }
    mc = {k: int((movement_df["Movement"] == k).sum()) for k in ["New", "Improved", "Declined", "Stable"]} if not movement_df.empty else {k: 0 for k in ["New", "Improved", "Declined", "Stable"]}
    highlights, priorities = [], []
    for _, row in data.iterrows():
        project = str(row.get("Project / Dashboard", "")).strip()
        delivered = str(row.get("This Week Delivered", "")).strip()
        nxt = str(row.get("Next Week Priority", "")).strip()
        if project and delivered and delivered.lower() not in {"none", "none reported", "n/a"}:
            highlights.append((project, delivered))
        if project and nxt and nxt.lower() not in {"none", "none reported", "n/a"}:
            priorities.append((project, nxt))
    return {
        "title": report_title, "period": _report_period_label(data), "scope": _report_scope_text(data),
        "generated": datetime.now(), "data": data, "prior": prior_df, "movement": movement_df,
        "actions": actions, "risks_df": risks_df, "outlook": _delivery_outlook(data), "kpis": kpis,
        "movement_counts": mc, "health_label": _health_label(kpis["health"]),
        "takeaway": _leadership_takeaway(data, movement_df), "highlights": highlights[:10], "priorities": priorities[:10]
    }


def _email_status_class(status):
    s = normalize_text(status)
    if s == "completed": return "success"
    if s == "blocked": return "critical"
    if s in {"at risk", "on hold"}: return "warning"
    if s == "on track": return "info"
    return "neutral"


def generate_html_report(report):
    """Production leadership email. Table-based HTML is intentionally used for Outlook compatibility."""
    k, m, o = report["kpis"], report["movement_counts"], report["outlook"]
    data = report["data"]
    esc = lambda v: html_escape(v)
    logo = '<div class="brand">FACTSPAN</div>'
    if os.path.exists("logo.png"):
        logo = '<img src="cid:factspan_logo" alt="Factspan" style="height:34px;display:block;">'
    rows = []
    for _, r in data.head(20).iterrows():
        status = str(r.get("Status", ""))
        rows.append(f'<tr><td><strong>{esc(r.get("Project / Dashboard", ""))}</strong><br><span class="muted">{esc(r.get("Resource", ""))}</span></td><td>{esc(r.get("Business Owner", ""))}</td><td><span class="pill {_email_status_class(status)}">{esc(status)}</span></td><td class="center"><b>{esc(r.get("Completion %", ""))}</b></td><td class="center">{esc(r.get("Utilization %", ""))}</td><td>{esc(r.get("Updated Expected Delivery date", ""))}</td></tr>')
    actions = []
    for item in report["actions"][:6]:
        critical = item["Priority"] == "Critical"
        action = "Escalate dependency and confirm recovery plan." if critical else "Confirm owner commitment and next milestone."
        actions.append(f'<div class="action {"critical" if critical else "warning"}"><strong>{esc(item["Project / Dashboard"])}</strong><div class="muted">{esc(item["Reason"])}</div><div class="action-text"><b>Recommended action:</b> {esc(action)}</div></div>')
    action_html = ''.join(actions) or '<div class="good">No active management attention items were identified for this selection.</div>'
    hi = ''.join(f'<li><strong>{esc(p)}</strong> — {esc(t)}</li>' for p,t in report["highlights"]) or '<li>No weekly accomplishment updates reported.</li>'
    nx = ''.join(f'<li><strong>{esc(p)}</strong> — {esc(t)}</li>' for p,t in report["priorities"]) or '<li>No next-week priorities reported.</li>'
    table_html = ''.join(rows) or '<tr><td colspan="6" class="muted center">No records in the current selection.</td></tr>'
    return f'''<!doctype html><html><head><meta charset="utf-8"><style>
body{{margin:0;background:#f4f7fb;font-family:Arial,Helvetica,sans-serif;color:#172033}}.wrap{{padding:24px 0}}.container{{max-width:940px;margin:auto;background:#fff;border:1px solid #e3e8ef}}.header{{padding:24px 30px 20px;border-bottom:4px solid #f58220}}.brand{{font-size:25px;font-weight:800;letter-spacing:1px;color:#16324f}}.eyebrow{{margin-top:16px;font-size:10px;font-weight:800;letter-spacing:1.5px;color:#f58220;text-transform:uppercase}}h1{{margin:5px 0;font-size:28px;color:#16324f}}.sub{{font-size:12px;color:#667085}}.section{{padding:20px 30px 3px}}.section-title{{font-size:13px;font-weight:800;letter-spacing:.7px;color:#16324f;text-transform:uppercase;margin-bottom:10px}}.kpis{{width:100%;border-collapse:separate;border-spacing:7px}}.kpi{{background:#f8fafc;border:1px solid #e3e8ef;padding:12px;text-align:center}}.kv{{font-size:24px;font-weight:800;color:#16324f}}.kl{{font-size:9px;font-weight:700;color:#667085;letter-spacing:.6px;margin-top:4px}}.takeaway{{background:#f7f9fc;border-left:4px solid #f58220;padding:14px 16px;font-size:13px;line-height:1.55}}.grid{{width:100%;border-collapse:collapse}}.grid td{{border:1px solid #e3e8ef;padding:9px;text-align:center}}.grid b{{font-size:16px;color:#16324f}}table.report{{width:100%;border-collapse:collapse;font-size:10px}}table.report th{{background:#edf2f7;color:#344054;padding:9px;text-align:left;font-size:8px;text-transform:uppercase}}table.report td{{padding:9px;border-bottom:1px solid #edf0f4;vertical-align:top}}.muted{{font-size:9px;color:#667085}}.center{{text-align:center}}.pill{{display:inline-block;padding:4px 7px;border-radius:12px;font-size:8px;font-weight:700}}.success{{background:#ecfdf3;color:#067647}}.warning{{background:#fff7e6;color:#a15c00}}.critical{{background:#fef2f2;color:#b42318}}.info{{background:#eef6ff;color:#175cd3}}.neutral{{background:#f2f4f7;color:#667085}}.action{{margin:7px 0;padding:10px 12px;border:1px solid #e3e8ef;border-left:4px solid #f79009;background:#fffdf7;font-size:10px}}.action.critical{{border-left-color:#d92d20;background:#fffafa}}.action-text{{margin-top:5px;color:#344054}}.good{{padding:11px;background:#ecfdf3;border:1px solid #abefc6;color:#067647;font-size:10px}}ul{{margin:4px 0 0 18px;padding:0;font-size:10px;line-height:1.65}}.footer{{margin-top:18px;padding:18px 30px 22px;border-top:1px solid #edf0f4;color:#98a2b3;font-size:9px}}
</style></head><body><div class="wrap"><div class="container"><div class="header">{logo}<div class="eyebrow">Weekly Project Management</div><h1>Weekly Delivery &amp; Portfolio Pulse</h1><div class="sub">{esc(report["period"])} · {esc(report["scope"])}</div></div>
<div class="section"><div class="section-title">Executive Pulse</div><table class="kpis"><tr><td class="kpi"><div class="kv">{k["projects"]}</div><div class="kl">PROJECT UPDATES</div></td><td class="kpi"><div class="kv">{k["resources"]}</div><div class="kl">RESOURCES</div></td><td class="kpi"><div class="kv">{k["completion"]}%</div><div class="kl">AVG. COMPLETION</div></td><td class="kpi"><div class="kv">{k["health"]}</div><div class="kl">PORTFOLIO HEALTH</div></td><td class="kpi"><div class="kv">{k["critical"]}</div><div class="kl">CRITICAL</div></td></tr></table></div>
<div class="section"><div class="section-title">Leadership Takeaway</div><div class="takeaway">{esc(report["takeaway"])}</div></div>
<div class="section"><div class="section-title">Week-over-Week Movement</div><table class="grid"><tr><td><b>{m["Improved"]}</b><br><span class="muted">Improved</span></td><td><b>{m["Declined"]}</b><br><span class="muted">Declined</span></td><td><b>{m["New"]}</b><br><span class="muted">New</span></td><td><b>{m["Stable"]}</b><br><span class="muted">Stable</span></td></tr></table></div>
<div class="section"><div class="section-title">Delivery Outlook</div><table class="grid"><tr><td><b>{o["Overdue"]}</b><br><span class="muted">Overdue</span></td><td><b>{o["Next 7 Days"]}</b><br><span class="muted">Next 7 Days</span></td><td><b>{o["8-14 Days"]}</b><br><span class="muted">8–14 Days</span></td><td><b>{o["15+ Days"]}</b><br><span class="muted">15+ Days</span></td><td><b>{o["Completed"]}</b><br><span class="muted">Completed</span></td></tr></table></div>
<div class="section"><div class="section-title">Management Attention</div>{action_html}</div><div class="section"><div class="section-title">Weekly Highlights</div><ul>{hi}</ul></div><div class="section"><div class="section-title">Next Week Priorities</div><ul>{nx}</ul></div>
<div class="section"><div class="section-title">Portfolio Detail</div><table class="report"><thead><tr><th>Initiative / Resource</th><th>Business Owner</th><th>Status</th><th>Completion</th><th>Utilization</th><th>Expected Delivery</th></tr></thead><tbody>{table_html}</tbody></table></div>
<div class="footer">Factspan · Weekly Project Report · Generated {report["generated"].strftime('%d %b %Y, %I:%M %p')}<br>Leadership summary shown above. Detailed PDF and Excel workbook are attached.</div></div></div></body></html>'''


def _pdf_safe(value):
    text = "" if value is None else str(value)
    for a,b in {"–":"-","—":"-","’":"'","“":'"',"”":'"',"•":"-","✓":"OK","⚠":"!","→":"->","₹":"Rs ","…":"..."}.items(): text=text.replace(a,b)
    return text.encode("latin-1","replace").decode("latin-1")


def _pdf_section(pdf, title):
    y = pdf.get_y() + 3
    pdf.set_fill_color(245,130,32); pdf.rect(15,y+1,2.2,7,"F")
    pdf.set_xy(21,y); pdf.set_text_color(22,50,79); pdf.set_font("Arial","B",12.5); pdf.cell(170,8,_pdf_safe(title),ln=True); pdf.ln(1)


def _pdf_lines(pdf, text, width, size=7.5, max_lines=4):
    text=_pdf_safe(text); words=text.split(); pdf.set_font("Arial","",size)
    if not words: return [""]
    lines=[]; cur=""
    for word in words:
        candidate=word if not cur else cur+" "+word
        if pdf.get_string_width(candidate)<=width: cur=candidate
        else:
            lines.append(cur or word); cur=word
            if len(lines)>=max_lines: break
    if cur and len(lines)<max_lines: lines.append(cur)
    if len(lines)==max_lines and len(" ".join(lines))<len(text):
        last=lines[-1]
        while last and pdf.get_string_width(last+"...")>width: last=last.rsplit(" ",1)[0] if " " in last else last[:-1]
        lines[-1]=(last.rstrip(" .")+"...") if last else "..."
    return lines


def _pdf_table(pdf, dataframe, headers, widths, getter):
    if dataframe is None or dataframe.empty: return
    pdf.set_fill_color(237,242,247); pdf.set_text_color(52,64,84); pdf.set_font("Arial","B",7)
    for h,w in zip(headers,widths): pdf.cell(w,8,_pdf_safe(h),border=1,fill=True)
    pdf.ln()
    for ridx,(_,row) in enumerate(dataframe.iterrows()):
        vals=getter(row); line_sets=[_pdf_lines(pdf,v,w-3) for v,w in zip(vals,widths)]
        h=max(7,max(len(x) for x in line_sets)*4.2+2)
        if pdf.get_y()+h>270:
            pdf.add_page(); pdf.set_y(18); pdf.set_fill_color(237,242,247); pdf.set_text_color(52,64,84); pdf.set_font("Arial","B",7)
            for hh,ww in zip(headers,widths): pdf.cell(ww,8,_pdf_safe(hh),border=1,fill=True)
            pdf.ln()
        x=15;y=pdf.get_y();pdf.set_fill_color(249,251,253) if ridx%2 else pdf.set_fill_color(255,255,255)
        for lines,w in zip(line_sets,widths):
            pdf.rect(x,y,w,h,"DF"); pdf.set_xy(x+1.5,y+1)
            pdf.set_text_color(52,64,84); pdf.set_font("Arial","",7.5)
            for line in lines:
                pdf.cell(w-3,4.2,_pdf_safe(line),ln=True);pdf.set_x(x+1.5)
            x+=w
        pdf.set_y(y+h)


def create_pdf_report(dataframe, header_title):
    if not FPDF_AVAILABLE: return None
    report=_build_report_model(dataframe,header_title);k=report["kpis"];o=report["outlook"]
    class ReportPDF(FPDF):
        def footer(self):
            self.set_y(-13);self.set_draw_color(226,232,240);self.line(15,285,195,285);self.set_text_color(100,116,139);self.set_font("Arial","",7);self.cell(125,5,"Factspan | Weekly Delivery & Portfolio Pulse");self.cell(55,5,f"Page {self.page_no()}",align="R")
    pdf=ReportPDF("P","mm","A4");pdf.set_margins(15,15,15);pdf.set_auto_page_break(True,18);pdf.add_page()
    pdf.set_fill_color(22,50,79);pdf.rect(0,0,210,45,"F");pdf.set_text_color(245,130,32);pdf.set_font("Arial","B",8);pdf.set_xy(15,8);pdf.cell(170,4,"FACTSPAN");pdf.set_text_color(255,255,255);pdf.set_font("Arial","B",20);pdf.set_xy(15,15);pdf.cell(175,9,"WEEKLY DELIVERY & PORTFOLIO PULSE");pdf.set_text_color(216,228,239);pdf.set_font("Arial","",9);pdf.set_xy(15,27);pdf.cell(175,5,_pdf_safe(report["period"]));pdf.set_font("Arial","",7.5);pdf.set_xy(15,35);pdf.cell(175,4,_pdf_safe(f"{report['scope']} | Generated {report['generated'].strftime('%d %b %Y, %I:%M %p')}"))
    pdf.set_y(51)
    cards=[("PROJECT UPDATES",k["projects"]),("RESOURCES",k["resources"]),("AVG. COMPLETION",f"{k['completion']}%"),("PORTFOLIO HEALTH",f"{k['health']}/100"),("CRITICAL",k["critical"])]
    for i,(lab,val) in enumerate(cards):
        x=15+i*37.2;pdf.set_fill_color(248,250,252);pdf.set_draw_color(226,232,240);pdf.rect(x,51,34.5,24,"DF");pdf.set_text_color(100,116,139);pdf.set_font("Arial","B",6.6);pdf.set_xy(x+2,55);pdf.cell(30.5,4,_pdf_safe(lab),align="C");pdf.set_text_color(22,50,79);pdf.set_font("Arial","B",14.5);pdf.set_xy(x+2,62);pdf.cell(30.5,8,_pdf_safe(val),align="C")
    _pdf_section(pdf,"Leadership Takeaway");pdf.set_text_color(52,64,84);pdf.set_font("Arial","",9.2);pdf.multi_cell(180,5.3,_pdf_safe(report["takeaway"]))
    _pdf_section(pdf,"Week-over-Week Movement");m=report["movement_counts"]
    for i,(lab,val) in enumerate([( "Improved",m["Improved"]),("Declined",m["Declined"]),("New",m["New"]),("Stable",m["Stable"])]):
        x=15+i*45.5;y=pdf.get_y();pdf.set_fill_color(248,250,252);pdf.rect(x,y,43.5,18,"DF");pdf.set_text_color(22,50,79);pdf.set_font("Arial","B",13);pdf.set_xy(x,y+2);pdf.cell(43.5,6,str(val),align="C");pdf.set_text_color(100,116,139);pdf.set_font("Arial","B",7);pdf.set_xy(x,y+10);pdf.cell(43.5,4,lab.upper(),align="C")
    pdf.set_y(pdf.get_y()+22);_pdf_section(pdf,"Delivery Outlook")
    for i,(lab,val) in enumerate([( "Overdue",o["Overdue"]),("Next 7 Days",o["Next 7 Days"]),("8-14 Days",o["8-14 Days"]),("15+ Days",o["15+ Days"]),("Completed",o["Completed"])]):
        x=15+i*37.2;y=pdf.get_y();pdf.set_fill_color(248,250,252);pdf.rect(x,y,34.5,17,"DF");pdf.set_text_color(22,50,79);pdf.set_font("Arial","B",12);pdf.set_xy(x,y+2);pdf.cell(34.5,5,str(val),align="C");pdf.set_text_color(100,116,139);pdf.set_font("Arial","B",6.2);pdf.set_xy(x,y+9);pdf.cell(34.5,4,lab.upper(),align="C")
    pdf.set_y(pdf.get_y()+21);_pdf_section(pdf,"Management Attention")
    if not report["actions"]:
        pdf.set_text_color(22,128,91);pdf.set_font("Arial","B",9);pdf.multi_cell(180,5.5,"No active management attention items were identified for this selection.")
    else:
        for item in report["actions"][:6]:
            priority=item["Priority"];pdf.set_text_color(185,28,28) if priority=="Critical" else pdf.set_text_color(161,92,0);pdf.set_font("Arial","B",8.8);pdf.multi_cell(180,4.8,_pdf_safe(f"{priority.upper()} | {item['Project / Dashboard']}"));pdf.set_text_color(71,84,103);pdf.set_font("Arial","",8.2);pdf.multi_cell(180,4.8,_pdf_safe(item["Reason"]));pdf.set_text_color(52,64,84);pdf.set_font("Arial","B",8);pdf.multi_cell(180,4.8,_pdf_safe("Recommended action: "+("Escalate dependency and confirm recovery plan." if priority=="Critical" else "Confirm owner commitment and next milestone.")));pdf.ln(1)
    pdf.add_page();pdf.set_text_color(22,50,79);pdf.set_font("Arial","B",18);pdf.cell(180,9,"Portfolio Delivery",ln=True);pdf.set_text_color(100,116,139);pdf.set_font("Arial","",8.5);pdf.cell(180,5,_pdf_safe(report["scope"]),ln=True);pdf.ln(3)
    _pdf_table(pdf,report["data"],["Initiative","Owner","Status","Completion","Utilization","Expected Delivery"],[58,27,27,20,20,28],lambda r:[r.get("Project / Dashboard",""),r.get("Business Owner",""),r.get("Status",""),r.get("Completion %",""),r.get("Utilization %",""),r.get("Updated Expected Delivery date","")])
    _pdf_section(pdf,"Weekly Highlights")
    for p,t in report["highlights"]:
        pdf.set_text_color(22,50,79);pdf.set_font("Arial","B",8.8);pdf.multi_cell(180,4.7,_pdf_safe(p));pdf.set_text_color(71,84,103);pdf.set_font("Arial","",8.3);pdf.multi_cell(180,4.7,_pdf_safe(t));pdf.ln(.5)
    if not report["highlights"]: pdf.set_text_color(100,116,139);pdf.set_font("Arial","",8.5);pdf.multi_cell(180,5,"No weekly accomplishment updates reported.")
    _pdf_section(pdf,"Next Week Priorities")
    for p,t in report["priorities"]:
        pdf.set_text_color(22,50,79);pdf.set_font("Arial","B",8.8);pdf.multi_cell(180,4.7,_pdf_safe(p));pdf.set_text_color(71,84,103);pdf.set_font("Arial","",8.3);pdf.multi_cell(180,4.7,_pdf_safe(t));pdf.ln(.5)
    if not report["priorities"]: pdf.set_text_color(100,116,139);pdf.set_font("Arial","",8.5);pdf.multi_cell(180,5,"No next-week priorities reported.")
    pdf.add_page();pdf.set_text_color(22,50,79);pdf.set_font("Arial","B",18);pdf.cell(180,9,"Detailed Weekly Updates",ln=True);pdf.ln(3)
    _pdf_table(pdf,report["data"],["Initiative","Resource","Work Type","This Week Delivered","Next Week Priority"],[40,27,25,44,44],lambda r:[r.get("Project / Dashboard",""),r.get("Resource",""),r.get("Work Type",""),r.get("This Week Delivered",""),r.get("Next Week Priority","")])
    if not report["risks_df"].empty:
        _pdf_section(pdf,"Risks & Blockers");_pdf_table(pdf,report["risks_df"],["Initiative","Status","Blocker / Risk","Delivery"],[50,28,72,30],lambda r:[r.get("Project / Dashboard",""),r.get("Status",""),r.get("Blocker","") or f"Status marked as: {r.get('Status','')}",r.get("Updated Expected Delivery date","")])
    out=pdf.output(dest="S");return out.encode("latin-1") if isinstance(out,str) else bytes(out)


def create_excel_report(dataframe, report_title):
    if not OPENPYXL_AVAILABLE:
        buffer=io.BytesIO();dataframe.to_excel(buffer,index=False);return buffer.getvalue()
    report=_build_report_model(dataframe,report_title);k=report["kpis"];wb=Workbook();buffer=io.BytesIO()
    navy="16324F";orange="F58220";light="F5F7FA";border="D9E2EC";muted="667085";header="EAF0F6"
    def title(ws,title,subtitle,cols):
        ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=cols);ws.cell(1,1,title);ws.cell(1,1).font=Font(name="Aptos Display",size=20,bold=True,color="FFFFFF");ws.cell(1,1).fill=PatternFill("solid",fgColor=navy);ws.cell(1,1).alignment=Alignment(vertical="center");ws.row_dimensions[1].height=32
        ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=cols);ws.cell(2,1,subtitle);ws.cell(2,1).font=Font(name="Aptos",size=10,color="DCE7F2");ws.cell(2,1).fill=PatternFill("solid",fgColor=navy);ws.row_dimensions[2].height=20
    def hdr(ws,row,cols):
        for c in range(1,cols+1): ws.cell(row,c).fill=PatternFill("solid",fgColor=header);ws.cell(row,c).font=Font(name="Aptos",bold=True,color=navy,size=9);ws.cell(row,c).alignment=Alignment(horizontal="center",vertical="center",wrap_text=True);ws.cell(row,c).border=Border(bottom=Side(style="thin",color=border))
        ws.row_dimensions[row].height=28
    # Dashboard
    ws=wb.active;ws.title="Executive Dashboard";ws.sheet_view.showGridLines=False;title(ws,"WEEKLY DELIVERY & PORTFOLIO PULSE",f"{report['period']} | {report['scope']} | Generated {report['generated'].strftime('%d %b %Y, %I:%M %p')}",10)
    cards=[("PROJECT UPDATES",k["projects"]),("RESOURCES",k["resources"]),("AVG. COMPLETION",f"{k['completion']}%"),("PORTFOLIO HEALTH",f"{k['health']}/100"),("CRITICAL",k["critical"])]
    for i,(lab,val) in enumerate(cards):
        c=1+i*2;ws.merge_cells(start_row=4,start_column=c,end_row=4,end_column=c+1);ws.merge_cells(start_row=5,start_column=c,end_row=6,end_column=c+1);ws.cell(4,c,lab);ws.cell(5,c,val)
        for rr in [4,5,6]:
            for cc in [c,c+1]: ws.cell(rr,cc).fill=PatternFill("solid",fgColor=light);ws.cell(rr,cc).border=Border(left=Side(style="thin",color=border),right=Side(style="thin",color=border),top=Side(style="thin",color=border),bottom=Side(style="thin",color=border))
        ws.cell(4,c).font=Font(bold=True,size=8,color=muted);ws.cell(4,c).alignment=Alignment(horizontal="center");ws.cell(5,c).font=Font(bold=True,size=18,color=navy);ws.cell(5,c).alignment=Alignment(horizontal="center",vertical="center")
    ws["A8"]="Leadership Takeaway";ws["A8"].font=Font(bold=True,size=12,color=navy);ws.merge_cells("A9:J11");ws["A9"]=report["takeaway"];ws["A9"].alignment=Alignment(wrap_text=True,vertical="top");ws["A9"].fill=PatternFill("solid",fgColor="F7F9FC")
    ws["A13"]="Week-over-Week Movement";ws["A13"].font=Font(bold=True,size=12,color=navy);ws["F13"]="Delivery Outlook";ws["F13"].font=Font(bold=True,size=12,color=navy)
    for i,(lab,val) in enumerate([(x,report["movement_counts"][x]) for x in ["Improved","Declined","New","Stable"]],1):ws.cell(14,i,lab);ws.cell(15,i,val);ws.cell(14,i).font=Font(bold=True,size=8,color=muted);ws.cell(15,i).font=Font(bold=True,size=15,color=navy);ws.cell(14,i).alignment=ws.cell(15,i).alignment=Alignment(horizontal="center")
    for i,(lab,val) in enumerate([(x,report["outlook"][x]) for x in ["Overdue","Next 7 Days","8-14 Days","15+ Days","Completed"]],6):ws.cell(14,i,lab);ws.cell(15,i,val);ws.cell(14,i).font=Font(bold=True,size=8,color=muted);ws.cell(15,i).font=Font(bold=True,size=15,color=navy);ws.cell(14,i).alignment=ws.cell(15,i).alignment=Alignment(horizontal="center")
    ws["A18"]="Management Attention";ws["A18"].font=Font(bold=True,size=12,color=navy);heads=["Priority","Initiative","Account","Team","Resource","Status","Delivery","Reason","Recommended Action"]
    for c,h in enumerate(heads,1):ws.cell(19,c,h)
    hdr(ws,19,len(heads))
    for r,item in enumerate(report["actions"][:15],20):
        action="Escalate dependency / confirm recovery plan" if item["Priority"]=="Critical" else "Confirm owner commitment and next milestone";vals=[item["Priority"],item["Project / Dashboard"],item["Account"],item["Team"],item["Resource"],item["Status"],item["Delivery"],item["Reason"],action]
        for c,v in enumerate(vals,1):ws.cell(r,c,v);ws.cell(r,c).alignment=Alignment(vertical="top",wrap_text=True);ws.cell(r,c).fill=PatternFill("solid",fgColor="FDECEC" if item["Priority"]=="Critical" else "FFF8E7")
    for i,w in enumerate([13,30,18,18,22,14,18,46,42],1):ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes="A19"
    # Tracker
    tr=wb.create_sheet("Project Tracker");tr.sheet_view.showGridLines=False;cols=["Account","Team","Year","Month","Week","Resource","Project / Dashboard","Business Owner","Start date","Initial Delivery date","Updated Expected Delivery date","Work Type","Completion %","Utilization %","Status","This Week Delivered","Next Week Priority","Blocker","Health","Attention"];title(tr,"PROJECT TRACKER",f"Detailed source-aligned project and weekly update view | {report['period']}",len(cols))
    for c,h in enumerate(cols,1):tr.cell(4,c,h)
    hdr(tr,4,len(cols))
    for r,row in enumerate(report["data"].to_dict("records"),5):
        pr,_=classify_attention(row);vals=[row.get(c,"") for c in cols[:-2]]+[f"{_portfolio_health(pd.DataFrame([row]))}/100",pr]
        for c,v in enumerate(vals,1):
            cell=tr.cell(r,c,v);cell.alignment=Alignment(vertical="top",wrap_text=True)
            if cols[c-1] in {"Completion %","Utilization %"}:cell.value=pct_value(v)/100;cell.number_format="0%"
            if cols[c-1] in {"Start date","Initial Delivery date","Updated Expected Delivery date"}:cell.number_format="dd-mmm-yyyy"
            if cols[c-1]=="Status":cell.font=Font(bold=True);cell.fill=PatternFill("solid",fgColor={"Completed":"EAF7EF","On Track":"EAF6FF","At Risk":"FFF4D6","Blocked":"FDECEC","On Hold":"FFF4D6"}.get(str(row.get("Status")),"F2F4F7"))
    for i,w in enumerate([15,16,10,13,10,22,32,22,14,18,22,18,13,13,14,44,44,38,14,13],1):tr.column_dimensions[get_column_letter(i)].width=w
    tr.freeze_panes="A5";tr.auto_filter.ref=f"A4:{get_column_letter(len(cols))}{max(4,len(report['data'])+4)}"
    # Other sheets
    ra=wb.create_sheet("Risks & Actions");ra.sheet_view.showGridLines=False;rc=["Priority","Project / Dashboard","Account","Team","Resource","Status","Delivery","Reason","Recommended Action"];title(ra,"RISKS & MANAGEMENT ACTIONS","Prioritized issues requiring attention",len(rc));[ra.cell(4,c,h) for c,h in enumerate(rc,1)];hdr(ra,4,len(rc))
    for r,item in enumerate(report["actions"],5):
        vals=[item[x] for x in ["Priority","Project / Dashboard","Account","Team","Resource","Status","Delivery","Reason"]]+["Escalate dependency / confirm recovery plan" if item["Priority"]=="Critical" else "Confirm owner commitment and next milestone"]
        for c,v in enumerate(vals,1):ra.cell(r,c,v);ra.cell(r,c).alignment=Alignment(vertical="top",wrap_text=True);ra.cell(r,c).fill=PatternFill("solid",fgColor="FDECEC" if item["Priority"]=="Critical" else "FFF8E7")
    for i,w in enumerate([13,32,18,18,22,14,18,50,44],1):ra.column_dimensions[get_column_letter(i)].width=w
    ra.freeze_panes="A5";ra.auto_filter.ref=f"A4:I{max(4,len(report['actions'])+4)}"
    wow=wb.create_sheet("Week-over-Week");wow.sheet_view.showGridLines=False;wc=["Project / Dashboard","Resource","Movement","Completion Δ","Utilization Δ","Status Change"];title(wow,"WEEK-OVER-WEEK MOVEMENT","Project-level movement against the immediately preceding reported week",len(wc));[wow.cell(4,c,h) for c,h in enumerate(wc,1)];hdr(wow,4,len(wc))
    for r,row in enumerate(report["movement"].to_dict("records"),5):
        for c,v in enumerate([row.get(x,"") for x in wc],1):wow.cell(r,c,v);wow.cell(r,c).alignment=Alignment(vertical="top",wrap_text=True)
    for i,w in enumerate([34,22,14,16,18,28],1):wow.column_dimensions[get_column_letter(i)].width=w
    wow.freeze_panes="A5";wow.auto_filter.ref=f"A4:F{max(4,len(report['movement'])+4)}"
    dq=wb.create_sheet("Data Quality");dq.sheet_view.showGridLines=False;dc=["Project / Dashboard","Resource","Status","Issue"];title(dq,"DATA QUALITY","Exceptions detected using the application's existing validation rules",len(dc));[dq.cell(4,c,h) for c,h in enumerate(dc,1)];hdr(dq,4,len(dc));rr=5
    for _,row in report["data"].iterrows():
        for issue in data_quality_flags(row):
            for c,v in enumerate([row.get("Project / Dashboard",""),row.get("Resource",""),row.get("Status",""),issue],1):dq.cell(rr,c,v);dq.cell(rr,c).alignment=Alignment(vertical="top",wrap_text=True)
            rr+=1
    for i,w in enumerate([34,22,16,65],1):dq.column_dimensions[get_column_letter(i)].width=w
    dq.freeze_panes="A5";dq.auto_filter.ref=f"A4:D{max(4,rr-1)}"
    for sh in wb.worksheets:
        sh.sheet_properties.pageSetUpPr.fitToPage=True;sh.page_setup.fitToWidth=1;sh.page_setup.fitToHeight=0;sh.page_margins.left=.25;sh.page_margins.right=.25;sh.page_margins.top=.5;sh.page_margins.bottom=.5;sh.oddFooter.center.text="Factspan | Weekly Project Report";sh.oddFooter.right.text="Page &P of &N"
    wb.save(buffer);return buffer.getvalue()


def _validated_emails(values):
    valid=[];invalid=[]
    for raw in values:
        addr=getaddresses([raw])[0][1].strip().lower()
        if addr and re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+",addr):valid.append(addr)
        elif raw.strip():invalid.append(raw.strip())
    return list(dict.fromkeys(valid)),list(dict.fromkeys(invalid))


def send_report_email(pdf_bytes, report_title, recipients, cc_recipients=None, excel_bytes=None, html_body=None):
    recipients,invalid_to=_validated_emails([x.strip() for x in recipients if x.strip()]);cc_recipients,invalid_cc=_validated_emails([x.strip() for x in (cc_recipients or []) if x.strip()])
    if invalid_to or invalid_cc:return False,"Invalid email address(es): "+", ".join(invalid_to+invalid_cc)
    if not recipients:return False,"Please enter at least one recipient email address."
    cfg=st.secrets if hasattr(st,"secrets") else {};host=cfg.get("SMTP_HOST",os.getenv("SMTP_HOST",""));port=int(cfg.get("SMTP_PORT",os.getenv("SMTP_PORT","587")));username=cfg.get("SMTP_USERNAME",os.getenv("SMTP_USERNAME",""));password=cfg.get("SMTP_PASSWORD",os.getenv("SMTP_PASSWORD",""));sender=cfg.get("SMTP_FROM",os.getenv("SMTP_FROM",username));tls=str(cfg.get("SMTP_USE_TLS",os.getenv("SMTP_USE_TLS","true"))).lower()=="true"
    if not host or not sender:return False,"Email sending is not configured (SMTP secrets missing)."
    msg=EmailMessage();msg["Subject"]=report_title;msg["From"]=sender;msg["To"]=", ".join(recipients)
    if cc_recipients:msg["Cc"]=", ".join(cc_recipients)
    msg.set_content(f"Please find attached the {report_title}. The email contains the leadership summary; the PDF and Excel workbook contain the detailed report.")
    msg.add_alternative(html_body or "<html><body><p>Please find the attached weekly report.</p></body></html>",subtype="html")
    if os.path.exists("logo.png") and html_body and "cid:factspan_logo" in html_body:
        with open("logo.png","rb") as f:msg.get_payload()[1].add_related(f.read(),maintype="image",subtype="png",cid="<factspan_logo>",filename="logo.png")
    if pdf_bytes:msg.add_attachment(pdf_bytes,maintype="application",subtype="pdf",filename="Weekly_Project_Report.pdf")
    if excel_bytes:msg.add_attachment(excel_bytes,maintype="application",subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",filename="Weekly_Project_Report.xlsx")
    try:
        with smtplib.SMTP(host,port,timeout=25) as server:
            server.ehlo()
            if tls:server.starttls();server.ehlo()
            if username and password:server.login(username,password)
            server.send_message(msg)
        return True,f"Report sent successfully to {', '.join(recipients)}."
    except Exception as exc:return False,f"Unable to send report: {exc}"


# =========================================================
# DATA + GLOBAL TOP AREA
# =========================================================
df = pd.DataFrame(st.session_state.notes_db)
filtered_df = pd.DataFrame()
reporting_week = "All Weeks"
reporting_month = "All Months"

st.markdown(f"""
<div class="app-hero">
    <div style="position: absolute; right: 25px; top: 25px; font-size: 13px; font-weight: bold; background: rgba(255,255,255,0.15); padding: 5px 12px; border-radius: 20px;">👤 {st.session_state.current_email} ({st.session_state.current_role})</div>
    <div class="app-hero-eyebrow">Weekly Project Management</div>
    <div class="app-hero-title">Weekly Project Report</div>
    <div class="app-hero-subtitle">Track progress, priorities, utilization and delivery across your portfolio.</div>
</div>
""", unsafe_allow_html=True)

# Lightweight utility bar — additive only.
util_c1, util_c2, util_c3, util_c4 = st.columns([1.2, 1.4, 5.2, 1.8])
with util_c1:
    if st.button("🔄 Refresh Data", key="refresh_data_top", use_container_width=True):
        refresh_cloud_data()
        st.session_state.last_refresh_at = datetime.now()
        st.rerun()
with util_c2:
    st.caption(f"Data loaded: {st.session_state.last_refresh_at.strftime('%d %b %Y, %I:%M %p')}")
with util_c4:
    if st.button("🔒 Logout", key="logout_top", use_container_width=True):
        st.session_state.authenticated = False
        st.session_state.current_email = ""
        st.session_state.current_role = ""
        st.session_state.current_scope = ""
        st.session_state.current_acc_scope = ""
        st.session_state.login_attempts = 0
        st.rerun()

# Navigation
labels = ["✍️ Enter Notes", "📋 Weekly Summary", "📈 Executive Report", "🚦 Action Center"]
cols = st.columns([1,1,1,1,1.5]) 
for col, label, m_idx in zip(cols[:4], labels, [1, 2, 3, 4]):
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
elif nav_selection == "🚦 Action Center":
    st.markdown('<div class="page-head"><div class="page-title">🚦 Action Center</div><div class="page-subtitle">Prioritized delivery risks, overdue work, data-quality issues and items needing management attention.</div></div>', unsafe_allow_html=True)

if nav_selection != "✍️ Enter Notes" and not df.empty:
    
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
        filter_year = st.multiselect("Year", options=year_options, key="filter_year")
        if filter_year: filtered_df = filtered_df[filtered_df["Year"].isin(filter_year)]
    with f_col4:
        month_options = list(filtered_df["Month"].unique())
        filter_month = st.multiselect("Month", options=month_options, key="filter_month")
        if filter_month: filtered_df = filtered_df[filtered_df["Month"].isin(filter_month)]
    with f_col5:
        week_options = list(filtered_df["Week"].unique())
        filter_week = st.multiselect("Week", options=week_options, key="filter_week")
        if filter_week: filtered_df = filtered_df[filtered_df["Week"].isin(filter_week)]
    with f_col6:
        res_options = list(filtered_df["Resource"].unique())
        selected_resources = st.multiselect("Resource Email(s)", options=res_options, key="filter_resources")
        if selected_resources: filtered_df = filtered_df[filtered_df["Resource"].isin(selected_resources)]

    reporting_month = format_badge_text(filter_month, month_options)
    reporting_week = format_badge_text(filter_week, week_options)

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
                with st.spinner("Generating leadership email, PDF and Excel report..."):
                    report_model = _build_report_model(filtered_df, subject_value)
                    current_pdf = create_pdf_report(filtered_df, subject_value) if FPDF_AVAILABLE else None
                    current_excel = create_excel_report(filtered_df, subject_value)
                    current_html = generate_html_report(report_model)
                if current_pdf is None: st.error("PDF generation unavailable.")
                else:
                    ok, message = send_report_email(current_pdf, subject_value, to_list, cc_list, current_excel, current_html)
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
            with t_col4: resource = st.text_input("Resource Email", value=st.session_state.current_email)

        # Copy Previous Week — additive shortcut. It never saves automatically.
        copy_col1, copy_col2, copy_col3 = st.columns([2.2, 2.2, 5.6])
        with copy_col1:
            if st.button("📋 Load Previous Week", key="load_previous_week", use_container_width=True):
                prev_source_df = pd.DataFrame(st.session_state.notes_db) if st.session_state.notes_db else pd.DataFrame()
                prev_for_copy = get_latest_prior_week(prev_source_df, year, month, week)
                # The expression above can be ambiguous for DataFrames; handle explicitly below.
                if prev_for_copy.empty:
                    st.info("No previous reported week was found for this timeframe.")
                else:
                    same_resource = prev_for_copy[prev_for_copy["Resource"].astype(str).str.lower() == str(resource).strip().lower()]
                    if same_resource.empty:
                        st.info("No previous-week entries found for this resource.")
                    else:
                        st.session_state.copy_previous_records = same_resource.to_dict("records")
                        st.session_state.copy_source_loaded = True
                        st.session_state.project_count = max(1, len(same_resource))
                        st.rerun()
        with copy_col2:
            if st.session_state.get("copy_source_loaded"):
                if st.button("✖ Clear Copied Values", key="clear_copied_values", use_container_width=True):
                    st.session_state.copy_previous_records = []
                    st.session_state.copy_source_loaded = False
                    st.rerun()
        if st.session_state.get("copy_source_loaded"):
            st.caption("Previous-week values are loaded as a starting point only. Review dates, completion, delivery notes and blockers before saving.")

        st.markdown('<div class="section-header">03 — Weekly Accomplishments</div>', unsafe_allow_html=True)
        projects_to_save = []

        for i in range(st.session_state.project_count):
            with st.container(border=True):
                st.markdown(f'<div class="card-title">Project Entry #{i + 1}</div>', unsafe_allow_html=True)
                st.markdown('<hr>', unsafe_allow_html=True)
                r1c1, r1c2 = st.columns(2)
                copy_rec = st.session_state.get("copy_previous_records", [])[i] if i < len(st.session_state.get("copy_previous_records", [])) else {}
                project = r1c1.text_input("Project / Dashboard Name", value=str(copy_rec.get("Project / Dashboard", "")), placeholder="e.g., Q3 Sales Analysis", key=f"proj_{i}")
                business_owner = r1c2.text_input("Business Owner", value=str(copy_rec.get("Business Owner", "")), placeholder="e.g., Marketing Team", key=f"owner_{i}")
                
                r2c1, r2c2, r2c3 = st.columns(3)
                start_date = r2c1.date_input("Start Date", value=parse_date(copy_rec.get("Start date")) if copy_rec else None, key=f"start_{i}")
                initial_delivery = r2c2.date_input("Initial Delivery Date", value=parse_date(copy_rec.get("Initial Delivery date")) if copy_rec else None, key=f"init_{i}")
                updated_delivery = r2c3.date_input("Updated Expected Delivery", value=parse_date(copy_rec.get("Updated Expected Delivery date")) if copy_rec else None, key=f"upd_{i}")
                
                r3c1, r3c2, r3c3, r3c4 = st.columns(4)
                wt_options = ["Development", "Enhancement", "Adhoc", "Migration", "Other"]
                existing_copy_wt = str(copy_rec.get("Work Type", ""))
                wt_selection = r3c1.selectbox("Work Type", wt_options, index=wt_options.index(existing_copy_wt) if existing_copy_wt in wt_options else 0, key=f"wt_sel_{i}")
                if wt_selection == "Other": work_type = r3c1.text_input("Specify Work Type", placeholder="Enter work type...", key=f"work_{i}")
                else: work_type = wt_selection
                
                status_opts = ["On Track", "At Risk", "Blocked", "Completed", "Not Started", "On Hold", "Cancelled"]
                copy_status = str(copy_rec.get("Status", "On Track"))
                status = r3c2.selectbox("Status", status_opts, index=status_opts.index(copy_status) if copy_status in status_opts else 0, key=f"status_{i}")
                completion_pct = r3c3.number_input("Completion %", min_value=0, max_value=100, step=5, value=int(pct_value(copy_rec.get("Completion %", 0))), key=f"comp_{i}")
                utilization_pct = r3c4.number_input("Utilization %", min_value=0, max_value=100, step=5, value=int(pct_value(copy_rec.get("Utilization %", 0))), key=f"util_{i}")
                
                c_text1, c_text2 = st.columns(2)
                this_week = c_text1.text_area("This Week Delivered", value="" if copy_rec else "", placeholder="What was accomplished this week?", height=85, key=f"this_{i}")
                next_week = c_text2.text_area("Next Week Priority", value=str(copy_rec.get("Next Week Priority", "")), placeholder="What is the focus for next week?", height=85, key=f"next_{i}")
                blocker = st.text_area("Blocker / Risks", value="", placeholder="Detail any risks or blockers here. (Leave blank if none)", height=68, key=f"block_{i}")
                
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
                if not resource: st.error("Please enter your Resource Email in Step 2 before saving.")
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
                            st.session_state.notes_db = load_data_from_gsheets()
                            st.session_state.data_synced = True
                            st.session_state.last_refresh_at = datetime.now()
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
        
        avg_util = int(filtered_df.groupby("Resource")["Utilization %"].apply(lambda x: sum(get_pct_decimal(v) * 100 for v in x)).mean()) if not filtered_df.empty else 0
        
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
            '<thead><tr><th>Resource Email</th><th>Business Owner</th><th>Project / Dashboard</th><th>Expected Delivery</th><th>Completion</th><th>Utilization</th><th>Status</th></tr></thead>',
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
            row_html = f"<tr><td>{html_escape(row['Resource'])}</td><td>{html_escape(row['Business Owner'])}</td><td><span class='proj-name'>{html_escape(row['Project / Dashboard'])}</span><span class='work-type'>{html_escape(row['Work Type'])}</span></td><td>{html_escape(row['Updated Expected Delivery date'])}</td><td>{html_escape(row['Completion %'])}</td><td>{html_escape(row['Utilization %'])}</td><td><span class='status-pill {pill}'>{html_escape(display_status)}</span></td></tr>"
            html_lines.append(row_html)
        html_lines.append("</tbody></table>")
        st.markdown("".join(html_lines), unsafe_allow_html=True)

        st.markdown('<div class="section-header">What Changed?</div>', unsafe_allow_html=True)
        single_year = filter_year[0] if len(filter_year) == 1 else None
        single_month = filter_month[0] if len(filter_month) == 1 else None
        single_week = filter_week[0] if len(filter_week) == 1 else None
        prior_df = get_latest_prior_week(df, single_year, single_month, single_week) if single_year and single_month and single_week else pd.DataFrame()
        movement_df = compare_current_to_prior(filtered_df, prior_df)
        if movement_df.empty:
            st.info("Select exactly one Year, Month and Week to see project-level week-over-week movement.")
        else:
            mv1, mv2, mv3, mv4 = st.columns(4)
            with mv1: st.markdown(render_metric_card("NEW", int((movement_df["Movement"] == "New").sum())), unsafe_allow_html=True)
            with mv2: st.markdown(render_metric_card("IMPROVED", int((movement_df["Movement"] == "Improved").sum())), unsafe_allow_html=True)
            with mv3: st.markdown(render_metric_card("DECLINED", int((movement_df["Movement"] == "Declined").sum())), unsafe_allow_html=True)
            with mv4: st.markdown(render_metric_card("STABLE", int((movement_df["Movement"] == "Stable").sum())), unsafe_allow_html=True)
            declined = movement_df[movement_df["Movement"] == "Declined"]
            if not declined.empty:
                st.warning("Completion declines of 10+ points were detected in the current selection.")
                st.dataframe(declined[["Project / Dashboard","Resource","Completion Δ","Status Change"]], use_container_width=True, hide_index=True)

        st.markdown('<div class="section-header">Blockers & Risks</div>', unsafe_allow_html=True)
        if risk_count == 0: st.success("No active blockers or risks reported this week! 🎉")
        else:
            for _, row in risk_items.iterrows():
                blocker_text = row['Blocker'] if row['Blocker'] else f"Status marked as: {row['Status']}"
                st.markdown(f'<div class="risk-box"><div class="risk-title">⚠️ {row["Project / Dashboard"]}</div><div class="risk-desc">{blocker_text}</div></div>', unsafe_allow_html=True)

# =========================================================
# VIEW 3: EXECUTIVE REPORT (WITH SMART EDIT BUTTONS)
# =========================================================
elif nav_selection == "📈 Executive Report":
    if df.empty or filtered_df.empty:
        st.info("No data available. Go to the 'Enter Notes' tab to log your first update.")
    else:
        st.markdown('<div class="section-header">Portfolio Management View</div>', unsafe_allow_html=True)
        ec1, ec2, ec3 = st.columns([2, 2, 6])
        with ec1:
            st.caption(f"{len(filtered_df)} project update(s) in current selection")
        with ec2:
            st.caption(f"{sum(1 for _, r in filtered_df.iterrows() if classify_attention(r)[0] == 'Critical')} critical item(s)")
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
                                    status_opts = ["On Track", "At Risk", "Blocked", "Completed", "Not Started", "On Hold", "Cancelled"]
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
                                                st.session_state.notes_db = load_data_from_gsheets()
                                                st.session_state.data_synced = True
                                                st.session_state.last_refresh_at = datetime.now()
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
                                        # --- SMART SECURITY LOGIC FOR EDIT BUTTON (WITH STRICT ACCOUNT MATCHING) ---
                                        can_edit = False
                                        curr_role = st.session_state.get("current_role", "")
                                        curr_email = st.session_state.get("current_email", "").strip().lower()
                                        curr_scope = [s.strip().lower() for s in st.session_state.get("current_scope", "").split(",") if s.strip()]
                                        curr_acc_scope = [s.strip().lower() for s in st.session_state.get("current_acc_scope", "").split(",") if s.strip()]
                                        
                                        row_team = str(row.get("Team", "")).strip().lower()
                                        row_acc = str(row.get("Account", "")).strip().lower()
                                        row_email = str(row.get("Resource", "")).strip().lower()
                                        
                                        if curr_role == "Admin":
                                            can_edit = True
                                        elif curr_role == "Account Manager" and row_acc in curr_scope:
                                            can_edit = True
                                        elif curr_role == "Team Lead" and row_team in curr_scope and row_acc in curr_acc_scope:
                                            can_edit = True
                                        elif row_email == curr_email:
                                            can_edit = True
                                            
                                        if can_edit:
                                            if st.button("✏️ Edit", key=f"edit_btn_{rec_id}", use_container_width=True):
                                                st.session_state.edit_id = rec_id; st.rerun()
                                        else:
                                            st.markdown("<div style='color:#94a3b8; font-size:12px; text-align:center; padding-top:10px;'>🔒 View Only</div>", unsafe_allow_html=True)

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
                                    # Delivery intelligence and data-quality cues.
                                    d_state = delivery_state(row)
                                    d_flags = data_quality_flags(row)
                                    if d_state.startswith("Overdue"):
                                        st.markdown(f"<div class='risk-box' style='margin-top:10px'><div class='risk-title'>⏰ {html_escape(d_state)}</div><div class='risk-desc'>Expected delivery: {html_escape(row.get('Updated Expected Delivery date','N/A'))}</div></div>", unsafe_allow_html=True)
                                    elif d_state.startswith("Due in"):
                                        st.markdown(f"<div class='edit-banner' style='margin-top:10px'>📅 <b>{html_escape(d_state)}</b> · Expected delivery: {html_escape(row.get('Updated Expected Delivery date','N/A'))}</div>", unsafe_allow_html=True)
                                    if d_flags:
                                        st.markdown(f"<div class='edit-banner' style='margin-top:10px'>🧹 <b>Data quality:</b> {html_escape('; '.join(d_flags))}</div>", unsafe_allow_html=True)

                                    if row["Blocker"] and str(row["Blocker"]).strip().lower() not in ["none", "na", "n/a", ""]:
                                        st.markdown(f"<div class='alert-card alert-danger' style='margin-top:14px; margin-bottom:0;'><div class='alert-title'>⚠ Blocker</div><div class='alert-desc'>{row['Blocker']}</div></div>", unsafe_allow_html=True)

# =========================================================
# VIEW 4: ACTION CENTER
# =========================================================
elif nav_selection == "🚦 Action Center":
    if df.empty:
        st.info("No data available. Go to the 'Enter Notes' tab to log your first update.")
    else:
        action_items = build_action_items(filtered_df if not filtered_df.empty else df)
        critical = [x for x in action_items if x["Priority"] == "Critical"]
        attention = [x for x in action_items if x["Priority"] == "Attention"]
        a1, a2, a3, a4 = st.columns(4)
        with a1: st.markdown(render_metric_card("CRITICAL", len(critical)), unsafe_allow_html=True)
        with a2: st.markdown(render_metric_card("ATTENTION", len(attention)), unsafe_allow_html=True)
        with a3: st.markdown(render_metric_card("OVERDUE", sum(1 for x in action_items if str(x["Delivery"]).startswith("Overdue"))), unsafe_allow_html=True)
        with a4: st.markdown(render_metric_card("TOTAL FLAGGED", len(action_items)), unsafe_allow_html=True)
        st.markdown('<div class="section-header">Prioritized Work Items</div>', unsafe_allow_html=True)
        if not action_items:
            st.success("Everything is currently within the defined attention rules. 🎉")
        else:
            action_df = pd.DataFrame(action_items)
            st.dataframe(action_df[["Priority","Project / Dashboard","Account","Team","Resource","Status","Delivery","Reason"]], use_container_width=True, hide_index=True)
        st.markdown('<div class="section-header">Data Quality Checks</div>', unsafe_allow_html=True)
        dq_records = []
        for _, row in (filtered_df if not filtered_df.empty else df).iterrows():
            flags = data_quality_flags(row)
            if flags:
                dq_records.append({"Project / Dashboard": row.get("Project / Dashboard",""), "Resource": row.get("Resource",""), "Status": row.get("Status",""), "Issues": "; ".join(flags)})
        if dq_records:
            st.dataframe(pd.DataFrame(dq_records), use_container_width=True, hide_index=True)
        else:
            st.success("No data-quality exceptions detected.")
