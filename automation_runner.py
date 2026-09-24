import os
import io
import json
import smtplib
import html
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from email.message import EmailMessage
from email.utils import getaddresses
import re
from collections import Counter

import gspread
from google.oauth2.service_account import Credentials
import pandas as pd

try:
    from fpdf import FPDF
    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False

IST = ZoneInfo("Asia/Kolkata")
SHEET_NAME = os.getenv("SHEET_NAME", "Weekly Notes Database")
EMAIL_LOGO_URL = "https://www.factspan.com/wp-content/uploads/2021/10/Factspan-Logo.png"

ALERT_WINDOW_HOURS = 3          
ALERT_REPEAT_MINUTES = 60       
RUN_TOLERANCE_MINUTES = 20      

# =========================================================
# CONNECTION & HELPERS
# =========================================================
def get_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    raw = os.environ["GCP_SERVICE_ACCOUNT_JSON"]
    info = json.loads(raw)
    creds = Credentials.from_service_account_info(info, scopes=scopes)
    return gspread.authorize(creds)

def get_spreadsheet(client):
    return client.open(SHEET_NAME)

def normalize_text(v): return str(v or "").strip().lower()

def pct_value(value):
    try: return max(0.0, min(100.0, float(str(value).replace("%", "").strip())))
    except Exception: return 0.0

def get_pct_decimal(pct_str):
    try: return float(str(pct_str).replace('%', '').strip()) / 100.0
    except Exception: return 0.0

def html_escape(value): return html.escape("" if value is None else str(value))

def safe_pdf_text(value):
    text = "" if value is None else str(value)
    replacements = {"–":"-", "—":"-", "'":"'", "“":"\"", "”":"\"", "•":"-", "✓":"OK", "⚠":"!", "×":"x"}
    for old, new in replacements.items(): text = text.replace(old, new)
    return text.encode("latin-1", "replace").decode("latin-1")

def compute_current_period(now=None):
    now = now or datetime.now(IST)
    year = now.year
    month = now.strftime("%B")
    week_num = min(5, ((now.day - 1) // 7) + 1)
    week = f"Week {week_num}"
    return year, month, week

def period_label(year, month, week): return f"{month} {year} · {week}"

def parse_hhmm(value, default="17:00"):
    try: return datetime.strptime(str(value).strip() or default, "%H:%M").time()
    except Exception: return datetime.strptime(default, "%H:%M").time()

def _validated_emails(values):
    valid, invalid = [], []
    for raw in values:
        addr = getaddresses([raw])[0][1].strip().lower()
        if addr and re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", addr): valid.append(addr)
        elif raw.strip(): invalid.append(raw.strip())
    return list(dict.fromkeys(valid)), list(dict.fromkeys(invalid))

def send_email(subject, html_body, recipients, cc=None, pdf_bytes=None, pdf_name="Weekly_Project_Report.pdf"):
    recipients, _ = _validated_emails([x for x in recipients if x])
    cc, _ = _validated_emails([x for x in (cc or []) if x])
    if not recipients: return False, "No valid recipients."

    host = os.getenv("SMTP_HOST", "")
    port = int(os.getenv("SMTP_PORT", "587"))
    username = os.getenv("SMTP_USERNAME", "")
    password = os.getenv("SMTP_PASSWORD", "")
    sender = os.getenv("SMTP_FROM", username)
    use_tls = str(os.getenv("SMTP_USE_TLS", "true")).lower() == "true"

    if not host or not sender: return False, "SMTP not configured."

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    if cc: msg["Cc"] = ", ".join(cc)
    
    msg.set_content("This message requires an HTML-capable email client to view.")
    msg.add_alternative(html_body, subtype="html")
    
    if os.path.exists("logo.png") and "cid:factspan_logo" in html_body:
        try:
            with open("logo.png", "rb") as img:
                msg.get_payload()[1].add_related(img.read(), maintype='image', subtype='png', cid='<factspan_logo>')
        except Exception: pass
        
    if pdf_bytes:
        msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)

    try:
        with smtplib.SMTP(host, port, timeout=20) as server:
            if use_tls: server.starttls()
            if username and password: server.login(username, password)
            server.send_message(msg)
        return True, f"Sent to {', '.join(recipients)}"
    except Exception as exc: return False, f"SMTP error: {exc}"

def resource_utilization(df):
    if df.empty or "Resource" not in df.columns: return 0
    vals = []
    for _, group in df.groupby("Resource"):
        nums = [pct_value(v) for v in group["Utilization %"]]
        if nums: vals.append(sum(nums)/len(nums))
    return int(round(sum(vals)/len(vals))) if vals else 0

def missing_resources_for_schedule(schedule, notes_df, year, month, week):
    expected = set(e.strip().lower() for e in str(schedule.get("Expected_Resources", "")).split(",") if e.strip())
    if notes_df.empty or not expected: return expected, expected, set()

    subset = notes_df[
        (notes_df.get("Account") == schedule.get("Account")) &
        (notes_df.get("Team") == schedule.get("Team")) &
        (notes_df.get("Year").astype(str) == str(year)) &
        (notes_df.get("Month") == month) &
        (notes_df.get("Week") == week)
    ]
    submitted = set(str(x).strip().lower() for x in subset["Resource"]) if not subset.empty and "Resource" in subset.columns else set()
    return expected - submitted, expected, submitted

# =========================================================
# PERFECTED HTML & PDF GENERATORS (RESTORED)
# =========================================================
def build_html_report(df, report_title, account, team):
    total_projects = len(df)
    delivered = int((df["This Week Delivered"].fillna("").astype(str).str.strip() != "").sum()) if not df.empty else 0
    avg_comp = int(df["Completion %"].apply(lambda x: get_pct_decimal(x) * 100).mean()) if not df.empty else 0
    risks = int(((df["Status"].isin(["At Risk", "Blocked"])) | (df["Blocker"].fillna("").astype(str).str.strip() != "")).sum()) if not df.empty else 0
    date_str = datetime.now(IST).strftime("%d %b %Y")
    account_team_str = f"{account} - {team}"

    table_rows = ""
    for _, row in df.iterrows():
        proj = html_escape(row.get("Project / Dashboard", ""))
        owner = html_escape(row.get("Business Owner", ""))
        raw_name = str(row.get("Resource Name", "")).strip()
        if raw_name and raw_name.lower() not in ["none", "na", "n/a", "nan", ""]: res = html_escape(raw_name)
        else: res = html_escape(str(row.get("Resource", "")).strip().split('@')[0].replace('.', ' ').title())
            
        status = html_escape(row.get("Status", "On Track"))
        comp = int(pct_value(row.get("Completion %")))
        due = html_escape(row.get("Updated Expected Delivery date", "N/A"))
        
        if status == "Completed": status_color = "#16a34a"
        elif status in ["At Risk", "In Progress", "On Hold"]: status_color = "#ea580c"
        elif status == "Blocked": status_color = "#dc2626"
        else: status_color = "#64748b"

        table_rows += f"""
        <tr>
            <td style="padding: 14px 12px; border-bottom: 1px solid #e2e8f0; font-size: 13px; color: #1e293b; word-break: break-word;">
                <strong>{proj}</strong><br><span style="color: #64748b; font-size: 11px;">Owner: {owner}</span>
            </td>
            <td style="padding: 14px 12px; border-bottom: 1px solid #e2e8f0; font-size: 13px; color: #475569; word-break: break-word;">{res}</td>
            <td style="padding: 14px 12px; border-bottom: 1px solid #e2e8f0; vertical-align: middle;">
                <div style="font-size: 11px; font-weight: bold; color: #1e293b; margin-bottom: 4px;">{comp}%</div>
                <div style="background-color: #e2e8f0; width: 100%; height: 6px; border-radius: 3px; overflow: hidden;">
                    <div style="background-color: {status_color}; width: {comp}%; height: 100%;"></div>
                </div>
            </td>
            <td style="padding: 14px 12px; border-bottom: 1px solid #e2e8f0; font-size: 13px; color: #475569;">{due}</td>
            <td style="padding: 14px 12px; border-bottom: 1px solid #e2e8f0; font-size: 12px; font-weight: bold; color: {status_color};">{status}</td>
        </tr>
        """

    highlights_html = ""
    added_projects = set()
    count = 0
    for _, row in df.iterrows():
        if count >= 4: break
        proj = row.get("Project / Dashboard", "")
        if proj in added_projects: continue
        deliv = str(row.get("This Week Delivered", "")).strip()
        if deliv and deliv.lower() not in ["none", "na", "n/a", "nan", "none reported.", ""]:
            highlights_html += f'<li style="margin-bottom: 8px;"><span style="color: #ea580c;">■</span> <strong>{html_escape(proj)}:</strong> {html_escape(deliv)}</li>'
            added_projects.add(proj)
            count += 1
                
    if count == 0: highlights_html = '<li><span style="color: #ea580c;">■</span> Routine progress tracking across all initiatives.</li>'
            
    logo_src = "cid:factspan_logo" if os.path.exists("logo.png") else EMAIL_LOGO_URL

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            @media only screen and (max-width: 600px) {{
                .email-container {{ width: 100% !important; padding: 0 !important; }}
                .mobile-stack {{ display: block !important; width: 100% !important; padding: 10px 0 !important; text-align: center !important; }}
                .mobile-hide {{ display: none !important; }}
                .kpi-box {{ display: block !important; width: 100% !important; margin-bottom: 10px !important; box-sizing: border-box; }}
                .logo-img {{ margin: 0 auto !important; }}
                .table-scroll {{ overflow-x: auto !important; display: block !important; width: 100% !important; }}
                .responsive-table {{ min-width: 600px !important; }}
                .footer-text {{ text-align: center !important; padding: 5px 0 !important; display: block !important; width: 100% !important; }}
            }}
        </style>
    </head>
    <body style="margin: 0; padding: 20px; background-color: #f4f7fb;">
        <div class="email-container" style="font-family: Arial, sans-serif; max-width: 850px; margin: 0 auto; background: #ffffff; border: 1px solid #dfe5ec;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="padding: 25px;">
                <tr>
                    <td class="mobile-stack" width="40%" valign="middle" align="left">
                        <img src="{logo_src}" alt="FACTSPAN" width="200" class="logo-img" style="display: block; border: 0; color: #1e293b; font-size: 26px; font-weight: bold;" />
                    </td>
                    <td class="mobile-stack" width="60%" align="right" valign="middle">
                        <div style="color: #ea580c; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px;">WEEKLY PROJECT REPORT</div>
                        <div style="font-size: 22px; color: #1e293b; margin: 4px 0 2px 0; font-weight: bold;">{account_team_str}</div>
                        <div style="color: #64748b; font-size: 12px;">Weekly Progress Snapshot &bull; {date_str}</div>
                    </td>
                </tr>
            </table>
            <div style="border-top: 4px solid #ea580c;"></div>
            <div style="padding: 25px;">
                <p style="color: #334155; font-size: 14px; margin-top: 0;">Hi Team,</p>
                <p style="color: #475569; font-size: 14px; line-height: 1.6; margin-bottom: 25px;">
                    Please find below the latest weekly progress snapshot for <strong>{account_team_str}</strong>.
                </p>
                <h3 style="color: #0f172a; font-size: 16px; margin-bottom: 12px; border-bottom: 2px solid #ea580c; padding-bottom: 5px; display: inline-block;">Portfolio Snapshot</h3>
                <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 25px;">
                    <tr>
                        <td class="kpi-box" width="23%" align="center" style="border: 1px solid #e2e8f0; padding: 20px 0; background: #f8fafc;">
                            <div style="font-size: 28px; font-weight: bold; color: #1e293b;">{total_projects}</div>
                            <div style="font-size: 10px; color: #64748b; text-transform: uppercase; margin-top: 4px;">PROJECTS</div>
                        </td>
                        <td class="mobile-hide" width="2%"></td>
                        <td class="kpi-box" width="23%" align="center" style="border: 1px solid #e2e8f0; padding: 20px 0; background: #f8fafc;">
                            <div style="font-size: 28px; font-weight: bold; color: #16a34a;">{delivered}</div>
                            <div style="font-size: 10px; color: #64748b; text-transform: uppercase; margin-top: 4px;">UPDATES</div>
                        </td>
                        <td class="mobile-hide" width="2%"></td>
                        <td class="kpi-box" width="23%" align="center" style="border: 1px solid #e2e8f0; padding: 20px 0; background: #fffbeb;">
                            <div style="font-size: 28px; font-weight: bold; color: #ea580c;">{avg_comp}%</div>
                            <div style="font-size: 10px; color: #64748b; text-transform: uppercase; margin-top: 4px;">COMPLETION</div>
                        </td>
                        <td class="mobile-hide" width="2%"></td>
                        <td class="kpi-box" width="23%" align="center" style="border: 1px solid #e2e8f0; padding: 20px 0; background: #fef2f2;">
                            <div style="font-size: 28px; font-weight: bold; color: #dc2626;">{risks}</div>
                            <div style="font-size: 10px; color: #64748b; text-transform: uppercase; margin-top: 4px;">RISKS</div>
                        </td>
                    </tr>
                </table>
                <div class="table-scroll">
                    <table class="responsive-table" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse: collapse; border: 1px solid #e2e8f0; margin-bottom: 30px;">
                        <thead>
                            <tr style="background-color: #1e293b;">
                                <th align="left" width="30%" style="padding: 14px 12px; color: #ffffff; font-size: 11px; font-weight: bold; text-transform: uppercase;">PROJECT / OWNER</th>
                                <th align="left" width="20%" style="padding: 14px 12px; color: #ffffff; font-size: 11px; font-weight: bold; text-transform: uppercase;">RESOURCE</th>
                                <th align="left" width="25%" style="padding: 14px 12px; color: #ffffff; font-size: 11px; font-weight: bold; text-transform: uppercase;">DELIVERY PROGRESS</th>
                                <th align="left" width="15%" style="padding: 14px 12px; color: #ffffff; font-size: 11px; font-weight: bold; text-transform: uppercase;">TARGET DATE</th>
                                <th align="left" width="10%" style="padding: 14px 12px; color: #ffffff; font-size: 11px; font-weight: bold; text-transform: uppercase;">STATUS</th>
                            </tr>
                        </thead>
                        <tbody>{table_rows}</tbody>
                    </table>
                </div>
                <div style="background-color: #fff7ed; border-left: 4px solid #ea580c; padding: 20px; margin-bottom: 30px;">
                    <h4 style="color: #0f172a; margin: 0 0 12px 0; font-size: 16px;">Key Updates & Actions</h4>
                    <ul style="margin: 0; padding-left: 20px; color: #475569; font-size: 13px; line-height: 1.6;">
                        {highlights_html}
                    </ul>
                </div>
            </div>
            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #1e293b; padding: 15px 25px;">
                <tr>
                    <td class="footer-text" align="left" style="color: #f1f5f9; font-size: 11px;">Factspan &bull; {account_team_str}</td>
                    <td class="footer-text" align="right" style="color: #ea580c; font-size: 11px; font-weight: bold;">Staying Relevant and Ahead through Fluid Intelligence</td>
                </tr>
            </table>
        </div>
    </body>
    </html>
    """

def build_pdf_report(dataframe, report_title, account, team):
    if not FPDF_AVAILABLE: return None
    account_team_str = f"{account} - {team}"
    date_str = datetime.now(IST).strftime("%d %b %Y, %I:%M %p")

    class ReportPDF(FPDF):
        def header(self):
            if os.path.exists("logo.png"):
                try: self.image("logo.png", 15, 10, 40)
                except Exception:
                    self.set_text_color(30, 41, 59); self.set_font("Arial", "B", 16)
                    self.set_xy(15, 15); self.cell(40, 10, "FACTSPAN", ln=False)
            else:
                self.set_text_color(30, 41, 59); self.set_font("Arial", "B", 16)
                self.set_xy(15, 15); self.cell(40, 10, "FACTSPAN", ln=False)

            self.set_text_color(234, 88, 12); self.set_font("Arial", "B", 8); self.set_xy(15, 10)
            self.cell(180, 4, "WEEKLY PROJECT REPORT", align="R", ln=True)
            self.set_text_color(30, 41, 59); self.set_font("Arial", "B", 18); self.set_x(15)
            self.cell(180, 8, safe_pdf_text(account_team_str), align="R", ln=True)
            self.set_text_color(100, 116, 139); self.set_font("Arial", "", 8); self.set_x(15)
            self.cell(180, 4, safe_pdf_text(report_title), align="R", ln=True)
            self.set_x(15); self.cell(180, 4, safe_pdf_text(f"Generated: {date_str}"), align="R", ln=True)
            self.set_y(32); self.set_draw_color(234, 88, 12); self.set_line_width(1)
            self.line(15, 32, 195, 32); self.set_line_width(0.2); self.ln(5) 

        def footer(self):
            self.set_y(-15); self.set_draw_color(234, 88, 12); self.set_line_width(0.5)
            self.line(15, 282, 195, 282); self.set_text_color(30, 41, 59); self.set_font("Arial", "B", 8)
            self.set_xy(15, 285); self.cell(90, 5, safe_pdf_text(f"Factspan \x95 {account_team_str}"), align="L")
            self.set_text_color(100, 116, 139); self.set_font("Arial", "", 8); self.set_xy(100, 285)
            self.cell(95, 5, f"Page {self.page_no()}", align="R")

    pdf = ReportPDF(orientation="P", unit="mm", format="A4")
    pdf.set_margins(15, 15, 15)
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

    total_projects = len(dataframe)
    total_resources = dataframe["Resource"].nunique() if not dataframe.empty and "Resource" in dataframe.columns else 0
    delivered = int((dataframe["This Week Delivered"].fillna("").astype(str).str.strip() != "").sum()) if not dataframe.empty else 0
    avg_comp = int(dataframe["Completion %"].apply(lambda x: get_pct_decimal(x) * 100).mean()) if not dataframe.empty else 0
    risks = int(((dataframe["Status"].isin(["At Risk", "Blocked"])) | (dataframe["Blocker"].fillna("").astype(str).str.strip() != "")).sum()) if not dataframe.empty else 0

    kpis = [("PROJECTS", total_projects), ("RESOURCES", total_resources), ("UPDATES", delivered), ("COMPLETION", f"{avg_comp}%"), ("RISKS", risks)]
    box_w = 33; gap = 3.75; y = 42
    for i, (label, value) in enumerate(kpis):
        x = 15 + i * (box_w + gap)
        pdf.set_fill_color(248, 250, 252); pdf.set_draw_color(226, 232, 240); pdf.rect(x, y, box_w, 24, "DF")
        pdf.set_text_color(100, 116, 139); pdf.set_font("Arial", "B", 6.8); pdf.set_xy(x, y + 4)
        pdf.cell(box_w, 4, safe_pdf_text(label), align="C")
        pdf.set_text_color(30, 41, 59); pdf.set_font("Arial", "B", 15); pdf.set_xy(x, y + 10)
        pdf.cell(box_w, 8, safe_pdf_text(str(value)), align="C")
    pdf.set_y(y + 32)

    def section(title):
        pdf.ln(5); y0 = pdf.get_y()
        pdf.set_fill_color(37, 99, 235)
        pdf.rect(13.5, y0 + 1, 1.0, 6, "F")  # Perfect 1mm outdent alignment
        pdf.set_xy(15, y0); pdf.set_text_color(22, 50, 79); pdf.set_font("Arial", "B", 13)
        pdf.cell(170, 8, safe_pdf_text(title), ln=True); pdf.ln(1)

    section("Executive Summary")
    pdf.set_x(15); pdf.set_text_color(52, 64, 84); pdf.set_font("Arial", "", 9.5)
    summary = (f"The selected report contains {total_projects} project update(s) across {total_resources} resource(s). "
               f"{delivered} project(s) include a weekly delivery update. Average completion is {avg_comp}% and average utilization is {resource_utilization(dataframe)}%. "
               f"There are {risks} item(s) requiring risk or blocker attention.")
    pdf.multi_cell(180, 5.8, safe_pdf_text(summary))

    section("Project Status")
    headers = [("Project", 60, "L"), ("Resource", 28, "L"), ("Status", 28, "L"), ("Comp.", 17, "C"), ("Util.", 17, "C"), ("Expected", 30, "L")]
    pdf.set_fill_color(231, 238, 246); pdf.set_text_color(37, 54, 74); pdf.set_font("Arial", "B", 7.8); pdf.set_x(15)
    for label, width, al in headers: pdf.cell(width, 8, safe_pdf_text(label), border=1, fill=True, align=al)
    pdf.ln(); pdf.set_font("Arial", "", 8.2)
    for ridx, (_, row) in enumerate(dataframe.iterrows()):
        pdf.set_x(15)
        if ridx % 2 == 1: pdf.set_fill_color(249, 251, 253)
        else: pdf.set_fill_color(255, 255, 255)
        project = safe_pdf_text(row.get("Project / Dashboard", "N/A"))[:40]
        raw_name = str(row.get("Resource Name", "")).strip()
        if raw_name and raw_name.lower() not in ["none", "na", "n/a", "nan", ""]: resource = safe_pdf_text(raw_name)[:20]
        else: resource = safe_pdf_text(str(row.get("Resource", "N/A")).strip().split('@')[0].replace('.', ' ').title())[:20]
        status = safe_pdf_text(row.get("Status", "N/A"))[:15]
        comp = safe_pdf_text(row.get("Completion %", "N/A"))
        util = safe_pdf_text(row.get("Utilization %", "N/A"))
        expected = safe_pdf_text(row.get("Updated Expected Delivery date", "N/A"))
        
        pdf.set_text_color(52, 64, 84)
        pdf.cell(60, 8, project, border=1, fill=True, align="L")
        pdf.cell(28, 8, resource, border=1, fill=True, align="L")
        if status in ("At Risk", "Blocked"): pdf.set_text_color(160, 55, 55)
        elif status == "Completed": pdf.set_text_color(22, 128, 91)
        pdf.cell(28, 8, status, border=1, fill=True, align="L"); pdf.set_text_color(52, 64, 84)
        pdf.cell(17, 8, comp, border=1, fill=True, align="C")
        pdf.cell(17, 8, util, border=1, fill=True, align="C")
        pdf.cell(30, 8, expected, border=1, fill=True, align="L"); pdf.ln()

    section("Weekly Highlights")
    for _, row in dataframe.iterrows():
        pdf.set_x(15)
        project = safe_pdf_text(row.get("Project / Dashboard", "N/A"))
        delivered_text = safe_pdf_text(row.get("This Week Delivered", "None reported."))
        pdf.set_text_color(22, 50, 79); pdf.set_font("Arial", "B", 9); pdf.multi_cell(180, 5, project)
        pdf.set_x(15); pdf.set_text_color(71, 84, 103); pdf.set_font("Arial", "", 9); pdf.multi_cell(180, 5, f"Delivered: {delivered_text or 'None reported.'}"); pdf.ln(1.5)

    risk_df = dataframe[(dataframe["Status"].isin(["At Risk", "Blocked"])) | (dataframe["Blocker"].fillna("").astype(str).str.strip() != "")]
    section("Risks & Blockers")
    if risk_df.empty:
        pdf.set_x(15); pdf.set_text_color(22, 128, 91); pdf.set_font("Arial", "B", 9); pdf.multi_cell(180, 6, safe_pdf_text("No active blockers or risks reported for this period."))
    else:
        for _, row in risk_df.iterrows():
            pdf.set_x(15); pdf.set_text_color(194, 65, 61); pdf.set_font("Arial", "B", 9); pdf.multi_cell(180, 5, safe_pdf_text(row.get("Project / Dashboard", "N/A")))
            pdf.set_x(15); pdf.set_text_color(71, 84, 103); pdf.set_font("Arial", "", 9)
            blocker = row.get("Blocker", "") or f"Status marked as: {row.get('Status', 'N/A')}"
            pdf.multi_cell(180, 5, safe_pdf_text(blocker)); pdf.ln(1.5)

    try: 
        out = pdf.output(dest="S")
        return out.encode("latin-1") if isinstance(out, str) else bytes(out)
    except Exception: return None

# =========================================================
# MAIN EXECUTION (Background Action)
# =========================================================
def main():
    client = get_client()
    spreadsheet = get_spreadsheet(client)

    try: schedules_sheet = spreadsheet.worksheet("Schedules")
    except Exception: return

    schedule_headers = schedules_sheet.row_values(1)
    schedules = schedules_sheet.get_all_records()
    notes = spreadsheet.sheet1.get_all_records()
    notes_df = pd.DataFrame(notes)
    
    now = datetime.now(IST)
    today_name = now.strftime("%A")
    today_date_str = now.strftime("%Y-%m-%d")
    year, month, week = compute_current_period(now)
    curr_period = period_label(year, month, week)

    for idx, s in enumerate(schedules):
        row_idx = idx + 2
        if str(s.get("Active", "Y")).strip().upper() == "N": continue

        account, team = s.get("Account", ""), s.get("Team", "")
        missing, expected, submitted = missing_resources_for_schedule(s, notes_df, year, month, week)
        recipients = [x.strip() for x in str(s.get("Recipients", "")).split(",") if x.strip()]
        cc = [x.strip() for x in str(s.get("CC", "")).split(",") if x.strip()]
        owner_email = s.get("Created_By_Email", "")

        send_time = parse_hhmm(s.get("Send_Time"))
        reminder_time = parse_hhmm(s.get("Reminder_Time"), default="09:00")
        send_dt_today = datetime.combine(now.date(), send_time, tzinfo=IST)
        reminder_dt_today = datetime.combine(now.date(), reminder_time, tzinfo=IST)
        window_start = send_dt_today - timedelta(hours=ALERT_WINDOW_HOURS)

        updates = {}

        # 1. Morning reminder
        if expected and abs((now - reminder_dt_today).total_seconds()) <= RUN_TOLERANCE_MINUTES * 60 and str(s.get("Last_Reminder_Date", "")) != today_date_str:
            body = f"<p>Hi team,</p><p>Friendly reminder to submit your weekly notes for <b>{html_escape(account)} / {html_escape(team)}</b> ({month} {year}, {week}) ahead of this {today_name}'s report.</p>"
            ok, msg = send_email(f"Reminder: Submit weekly notes — {account} / {team}", body, sorted(expected), [owner_email])
            updates["Last_Reminder_Date"] = today_date_str

        # 2. Escalating pre-send alerts
        if missing and today_name == s.get("Send_Day") and window_start <= now < send_dt_today:
            last_alert = s.get("Last_Alert_At", "")
            should_alert = True
            if last_alert:
                try:
                    last_alert_dt = datetime.strptime(last_alert, "%Y-%m-%d %H:%M:%S").replace(tzinfo=IST)
                    should_alert = (now - last_alert_dt).total_seconds() >= ALERT_REPEAT_MINUTES * 60
                except Exception: should_alert = True
                
            if should_alert:
                body = (f"<p>Hi,</p><p><b>{len(missing)} of {len(expected)}</b> team member(s) have not yet submitted weekly notes for "
                        f"<b>{html_escape(account)} / {html_escape(team)}</b> ({month} {year}, {week}). "
                        f"The report is scheduled to send at {s.get('Send_Time')} IST today.</p>"
                        f"<ul>{''.join(f'<li>{html_escape(m)}</li>' for m in sorted(missing))}</ul><p>Please submit as soon as possible.</p>")
                ok, msg = send_email(f"URGENT: Weekly notes pending — {account} / {team}", body, sorted(missing) + [owner_email])
                updates["Last_Alert_At"] = now.strftime("%Y-%m-%d %H:%M:%S")

        # 3. Scheduled Send
        if today_name == s.get("Send_Day") and now >= send_dt_today and str(s.get("Last_Sent_Period", "")) != curr_period:
            subset = notes_df[(notes_df.get("Account") == account) & (notes_df.get("Team") == team) & (notes_df.get("Year").astype(str) == str(year)) & (notes_df.get("Month") == month) & (notes_df.get("Week") == week)] if not notes_df.empty else pd.DataFrame()
            report_title = f"{s.get('Report_Type','Weekly Summary')}: {month} {year} - {week}"
            html_body = build_html_report(subset, report_title, account, team)
            pdf_bytes = build_pdf_report(subset, report_title, account, team) if not subset.empty else None
            ok, msg = send_email(report_title, html_body, recipients, cc, pdf_bytes=pdf_bytes)
            updates["Last_Sent_Period"] = curr_period

        if updates:
            for col, val in updates.items():
                if col in schedule_headers:
                    schedules_sheet.update_cell(row_idx, schedule_headers.index(col) + 1, val)

if __name__ == "__main__":
    main()
