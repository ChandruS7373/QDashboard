import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import anthropic
import os
import html
import re
from datetime import datetime, date, timedelta
from jinja2 import Template
import auth

# ── PAGE CONFIG ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Qualesce AI Project Manager",
    page_icon="Q",
    layout="wide",
    initial_sidebar_state="collapsed",
)
if "db_initialized" not in st.session_state:
    auth.init_db()
    st.session_state.db_initialized = True

# ── EXCEL PATH ────────────────────────────────────────────────────────────────
EXCEL_PATH       = os.path.join(os.path.dirname(os.path.abspath(__file__)), "projects.xlsx")
USERS_EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "users.xlsx")
EXCEL_COLS = ["id","name","client","lead","employee","status","proj_type","start","end","due_date","po","desc",
              "manual_hrs","auto_hrs","cost_per_hr","hours_saved","cost_saved","roi_pct","is_new","is_active"]

# ── BASE DATA ─────────────────────────────────────────────────────────────────
BASE_PROJECTS = [
    {"id":1,  "name":"Raychem GATE Entry and GRN Process - Part A",     "client":"Raychem",                 "employee":"Nandukanth & Radhika","start":"20/07/2025","end":"",           "status":"R&M",          "po":"456788","desc":"GATE Entry and GRN Creation",                               "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":2,  "name":"PO Creation Process",                              "client":"Swagekklok - California","employee":"Akhila Kovuri",        "start":"22/09/2025","end":"17/12/2025","status":"R&M",          "po":"789747","desc":"Downloading PO Creation Report",                                "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":3,  "name":"Payment Application Process for all Company",      "client":"Swagekklok - California","employee":"Chethan B N",          "start":"22/09/2025","end":"19/01/2026","status":"R&M",          "po":"984534","desc":"Posting Cr amount to required Company code",                     "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":4,  "name":"Payments Application Process for LAMPAY",          "client":"Swagekklok - California","employee":"Chethan B N",          "start":"22/09/2025","end":"24/02/2026","status":"Discontinued", "po":"786540","desc":"",                                                             "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":5,  "name":"LAMCON Invoice Consignment",                       "client":"Swagekklok - California","employee":"Akhila Kovuri",        "start":"16/10/2025","end":"",           "status":"UAT",          "po":"983240","desc":"Creating Invoice Number by posting Material Number in SAP B1", "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":6,  "name":"CyberSource Application Process",                  "client":"Swagekklok - California","employee":"Akhila Kovuri",        "start":"02/12/2026","end":"",           "status":"PDD",          "po":"451238","desc":"",                                                             "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":7,  "name":"HR Master Data",                                   "client":"TEPL",                   "employee":"Mathan",               "start":"18/06/2022","end":"",           "status":"R&M",          "po":"933248","desc":"",                                                             "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":8,  "name":"MIS Report Process",                               "client":"TEPL",                   "employee":"Narendra",             "start":"20/04/2024","end":"",           "status":"UAT",          "po":"84973", "desc":"",                                                             "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":9,  "name":"PO Parking",                                       "client":"TEPL",                   "employee":"Nandukanth",           "start":"20/11/2024","end":"18/01/2025","status":"Completed",     "po":"213480","desc":"",                                                             "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":10, "name":"Greeting Process",                                 "client":"TEPL",                   "employee":"Sushma",               "start":"01/03/2025","end":"25/03/2025","status":"R&M",          "po":"345576","desc":"",                                                             "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":11, "name":"Data Migration from SAP to Salesforce",            "client":"Swagelok - Alabama",     "employee":"Sushma",               "start":"04/03/2025","end":"14/05/2025","status":"R&M",          "po":"543778","desc":"Moving data from SAP to Salesforce",                            "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":12, "name":"Tubing SCS Inspection Instructions",               "client":"Swagelok - Alabama",     "employee":"Sushma",               "start":"04/04/2025","end":"",           "status":"Discontinued", "po":"432670","desc":"Extraction of Specific data from PDFs",                         "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":13, "name":"Generate Batch Invoices",                          "client":"Swagelok - Alabama",     "employee":"Sushma",               "start":"06/05/2025","end":"28/02/2026","status":"R&M",          "po":"355377","desc":"Identifying the correct batches in SAP B1",                     "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":14, "name":"Sorting CTM Requests in ZenDesk",                 "client":"Swagelok - Alabama",     "employee":"Chethan B N",          "start":"06/10/2025","end":"",           "status":"Discontinued", "po":"872351","desc":"Organizing and prioritizing CTM tickets in Zendesk",            "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":15, "name":"Emailed PDF Cert Instructions",                    "client":"Swagelok - Alabama",     "employee":"Vikas",                "start":"09/03/2025","end":"17/11/2025","status":"R&M",          "po":"762345","desc":"Extracting the PO numbers from the PDFs",                       "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":16, "name":"Generate Quotes from Solidworks BOM",              "client":"Swagelok - Alabama",     "employee":"Vikas",                "start":"30/09/2025","end":"",           "status":"UAT",          "po":"672552","desc":"Creating quotations using data from SolidWorks BOM",            "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":17, "name":"Renaming and Saving Quality Records",              "client":"Swagelok - Alabama",     "employee":"Sushma",               "start":"01/08/2026","end":"06/02/2026","status":"R&M",          "po":"765428","desc":"Renaming and storing quality documents",                         "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":18, "name":"TDS Process",                                      "client":"TEPL",                   "employee":"Avinash",              "start":"07/10/2025","end":"",           "status":"UAT",          "po":"267357","desc":"Create TDS report and share to User",                           "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":19, "name":"GST Process",                                      "client":"TEPL",                   "employee":"Sharan",               "start":"07/10/2025","end":"",           "status":"UAT",          "po":"872610","desc":"Create GST report and share to User",                           "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":20, "name":"Job Scheduling - Cancellation Process",            "client":"TEPL",                   "employee":"Shiv Shankar",         "start":"22/08/2025","end":"01/06/2026","status":"Completed",     "po":"465738","desc":"Check for cancellation jobs in SAP and send alert mail",         "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":21, "name":"Job Scheduling - Active Process",                  "client":"TEPL",                   "employee":"Shiv Shankar",         "start":"24/10/2025","end":"01/06/2026","status":"Completed",     "po":"749474","desc":"Check for Active jobs in SAP and send alert mail",              "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":22, "name":"Fixture Automation",                               "client":"TEPL",                   "employee":"Nischal",              "start":"14/11/2025","end":"",           "status":"UAT",          "po":"248490","desc":"Download the fixture dump data and append every 2 mins",        "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":23, "name":"Invoice Posting",                                  "client":"TEPL",                   "employee":"Mathan",               "start":"",          "end":"",           "status":"Completed",     "po":"353628","desc":"Need to post the invoices",                                      "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":24, "name":"Vendor Confirmation",                              "client":"TEPL",                   "employee":"Mathan",               "start":"06/08/2025","end":"06/11/2025","status":"Completed",     "po":"235367","desc":"Need to confirm the vendor codes",                               "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":25, "name":"Block Stock Auto Mail Trigger Process",            "client":"TEPL",                   "employee":"Chethan B N",          "start":"01/07/2025","end":"21/01/2026","status":"R&M",          "po":"484640","desc":"Downloading Block Stock Report and send to user",               "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":26, "name":"Tool Crib Auto Reservation - FIFO Process",        "client":"TEPL",                   "employee":"Shiv Shankar",         "start":"12/02/2025","end":"",           "status":"In Progress",  "po":"674537","desc":"Reserve the tool crib data in FIFO Order",                      "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":27, "name":"Tool Crib Posting Goods Issue (GI) - MT 201",      "client":"TEPL",                   "employee":"Nischal",              "start":"12/02/2025","end":"",           "status":"In Progress",  "po":"380273","desc":"Post Goods Issue to movement type 201",                          "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":28, "name":"QA32 Dump Process",                                "client":"TEPL",                   "employee":"Avinash",              "start":"02/10/2026","end":"",           "status":"In Progress",  "po":"345468","desc":"Download the QA32 Dump file and append every 2 mins",          "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":29, "name":"AP Scheduling",                                    "client":"TEPL",                   "employee":"Mathan & Fiaz",        "start":"13/02/2026","end":"",           "status":"PDD",          "po":"189375","desc":"Create Finance Report",                                          "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":30, "name":"Attendance DB Update - TEPL Assembly & Raw Clock", "client":"TEPL",                   "employee":"Sivin",                "start":"26/12/2026","end":"05/02/2026","status":"Completed",     "po":"345465","desc":"Downloading and Updating Employee Assembly and Raw Clock data",  "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":32, "name":"BOM",                                              "client":"TEPL",                   "employee":"Sharan",               "start":"",          "end":"",           "status":"Completed",     "po":"",      "desc":"",                                                             "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":33, "name":"Production Order",                                 "client":"TEPL",                   "employee":"Sharan",               "start":"",          "end":"",           "status":"Completed",     "po":"",      "desc":"",                                                             "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":34, "name":"Google Cloud Platform",                            "client":"Internal POC",           "employee":"Faiyaz",               "start":"15/12/2025","end":"10/01/2026","status":"Internal POC",  "po":"",      "desc":"Agentic Platform using Google Cloud and NotebookLM",             "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":35, "name":"Microsoft Agent Frontier",                         "client":"Internal POC",           "employee":"Faiyaz",               "start":"05/11/2025","end":"14/11/2025","status":"Internal POC",  "po":"",      "desc":"Agentic Platform using Microsoft Copilot Frontier",              "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":36, "name":"UiPath Test Automation",                           "client":"Internal POC",           "employee":"Faiyaz",               "start":"07/01/2026","end":"",           "status":"Internal POC",  "po":"",      "desc":"Explored TestManager, Test Cloud and Test Automation",           "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":37, "name":"Qpdf",                                             "client":"Internal POC",           "employee":"Chandru S",            "start":"10/03/2025","end":"14/05/2025","status":"Internal POC",  "po":"",      "desc":"Chat with PDF AI",                                               "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":38, "name":"EDI",                                              "client":"Internal POC",           "employee":"Chandru S & Sivin",    "start":"22/05/2025","end":"20/06/2025","status":"Internal POC",  "po":"",      "desc":"Conversion of Medicare Insurance PDF to EDI",                    "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":39, "name":"krista.ai",                                        "client":"Internal POC",           "employee":"Chandru S",            "start":"02/06/2025","end":"",           "status":"Internal POC",  "po":"",      "desc":"Agentic Platform using Krista.ai",                               "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":40, "name":"UiPath with GitHub",                               "client":"Internal POC",           "employee":"Chandru S",            "start":"20/06/2025","end":"24/06/2025","status":"Internal POC",  "po":"",      "desc":"Integrated UiPath usecases with GitHub",                         "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":41, "name":"Microsoft Copilot Agent with Studio",              "client":"Internal POC",           "employee":"Chandru S",            "start":"10/11/2025","end":"",           "status":"Internal POC",  "po":"",      "desc":"Agentic Platform using Microsoft Copilot Studio",                "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":42, "name":"N8N",                                              "client":"Internal POC",           "employee":"Chandru S",            "start":"07/07/2025","end":"09/07/2025","status":"Internal POC",  "po":"",      "desc":"Agentic Platform using N8N",                                     "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":43, "name":"UiPath Standalone",                                "client":"Internal POC",           "employee":"Chandru S",            "start":"04/02/2026","end":"",           "status":"Internal POC",  "po":"",      "desc":"Explored UiPath On-premise Studio, Orchestrator, Test Suite",    "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":44, "name":"SHAREPOINT",                                       "client":"Internal POC",           "employee":"Rubika AE",            "start":"24/01/2026","end":"23/01/2026","status":"Internal POC",  "po":"",      "desc":"Web Development Project Tracker",                                "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":45, "name":"AWS - Knowledge Based Agent",                      "client":"Internal POC",           "employee":"Rubika AE",            "start":"02/02/2026","end":"05/02/2026","status":"Internal POC",  "po":"",      "desc":"Agentic Platform using Amazon Bedrock and S3 bucket",            "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":46, "name":"Himatsingha",                                      "client":"External POC",           "employee":"Narendra",             "start":"17/02/2026","end":"",           "status":"External POC",  "po":"",      "desc":"Reconciliation of Files to create a template",                   "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
    {"id":47, "name":"IBM RPA Tool",                                     "client":"Internal POC",           "employee":"Sharan",               "start":"",          "end":"",           "status":"Internal POC",  "po":"",      "desc":"",                                                             "manual_hrs":"","auto_hrs":"","cost_per_hr":"","hours_saved":"","cost_saved":"","roi_pct":"","is_new":False,"is_active":True},
]

ALL_STATUSES  = ["R&M","UAT","In Progress","Completed","PDD","Discontinued","Internal POC","External POC","Important","Presales"]
STATUS_STYLES = {
    "R&M":          {"bg":"#EFF6FF","text":"#1D4ED8","dot":"#3B82F6"},
    "UAT":          {"bg":"#FFFBEB","text":"#92400E","dot":"#F59E0B"},
    "Completed":    {"bg":"#ECFDF5","text":"#065F46","dot":"#10B981"},
    "In Progress":  {"bg":"#ECFEFF","text":"#155E75","dot":"#06B6D4"},
    "PDD":          {"bg":"#FFF7ED","text":"#9A3412","dot":"#F97316"},
    "Discontinued": {"bg":"#FEF2F2","text":"#991B1B","dot":"#EF4444"},
    "Internal POC": {"bg":"#F5F3FF","text":"#5B21B6","dot":"#8B5CF6"},
    "External POC": {"bg":"#FDF2F8","text":"#9D174D","dot":"#EC4899"},
    "Important":    {"bg":"#FFF1F2","text":"#BE123C","dot":"#F43F5E"},
    "Presales":     {"bg":"#F0F9FF","text":"#0369A1","dot":"#0EA5E9"},
}
STATUS_CHART_COLORS = ["#3B82F6","#F59E0B","#06B6D4","#10B981","#F97316","#EF4444","#8B5CF6","#EC4899","#F43F5E","#0EA5E9"]

SYSTEM_PROMPT = """You are an AI Project Manager Agent for Qualesce (RPA automation company).
BASE PORTFOLIO: 46 projects across Raychem(1), Swagekklok-CA(5), Swagelok-AL(7), TEPL(19), Internal POC(13), External POC(1).
STATUS MIX: R&M(10), UAT(6), Completed(8), In Progress(3), POC(14), Discontinued(3), PDD(2), Important(flagged critical).
STATUSES: R&M (Run & Maintain), UAT (User Acceptance Testing), In Progress, Completed, PDD (Pre-Due Diligence), Discontinued, Internal POC, External POC, Important (high-priority flagged tasks needing immediate attention).
TEAM: 16 members — Akhila Kovuri, Avinash, Chethan B N, Faiyaz, Mathan, Nandukanth, Narendra, Nischal, Radhika, Sharan, Shiv Shankar, Sivin, Sushma, Vikas, Chandru S, Rubika AE.
ROI FORMULA: Hours Saved = Manual Hrs - Auto Hrs | Cost Saved = Hours Saved x Cost/Hr | ROI% = (Hours Saved / Manual Hrs) x 100

FORMATTING RULES (always follow):
- When listing multiple projects, people, or items with attributes → use a markdown table with | column | headers |
- When explaining steps, reasons, or a summary → use bullet points (- item)
- Never write long prose paragraphs — always break into bullets
- Show ROI formula steps when calculating
- Be concise and data-driven"""

# ── EXCEL HELPERS ─────────────────────────────────────────────────────────────
def save_to_excel(df: pd.DataFrame):
    import tempfile, shutil
    out = df.copy()
    for col in EXCEL_COLS:
        if col not in out.columns:
            out[col] = ""
    _poc_statuses_excel = {"Presales", "Internal POC", "External POC"}
    presales_df = out[out["status"].str.strip().isin(_poc_statuses_excel)][EXCEL_COLS].reset_index(drop=True)
    license_records = auth.get_all_licenses()
    license_df = pd.DataFrame(license_records) if license_records else pd.DataFrame(
        columns=["id", "tool_name", "no_of_licenses", "start_date", "end_date", "created_at"]
    )
    user_records = auth.get_all_users()
    user_df = pd.DataFrame(user_records, columns=["id","name","email","role","is_active","created_at"]) \
              if user_records else pd.DataFrame(columns=["id","name","email","role","is_active","created_at"])
    sold_records = auth.get_all_sold_licenses()
    sold_df = pd.DataFrame(sold_records) if sold_records else pd.DataFrame(
        columns=["id","tool_name","client","no_of_licenses","start_date","end_date","notes","created_at"]
    )
    # Write to a temp file first, then replace — avoids PermissionError when Excel has the file open
    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(EXCEL_PATH))
    os.close(tmp_fd)
    try:
        with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
            out[EXCEL_COLS].to_excel(writer, sheet_name="Project Details", index=False)
            presales_df.to_excel(writer, sheet_name="Presales_POC", index=False)
            license_df.to_excel(writer, sheet_name="License", index=False)
            sold_df.to_excel(writer, sheet_name="Sold_License", index=False)
            user_df.to_excel(writer, sheet_name="Users", index=False)
        shutil.move(tmp_path, EXCEL_PATH)
    except PermissionError:
        # projects.xlsx is open in Excel — keep temp file as fallback and surface a clear warning
        st.warning(
            "⚠️ **Could not save to projects.xlsx** — the file is open in Excel. "
            "Please close Excel and click **Sync** to reload, or your changes are held in memory only.",
            icon=None,
        )
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise

def load_from_excel() -> pd.DataFrame:
    if os.path.exists(EXCEL_PATH):
        try:
            return pd.read_excel(EXCEL_PATH, sheet_name="Project Details", dtype=str, engine="openpyxl").fillna("")
        except Exception:
            try:
                return pd.read_excel(EXCEL_PATH, dtype=str, engine="openpyxl").fillna("")
            except Exception:
                # File is corrupted — back it up and rebuild from BASE_PROJECTS
                import shutil
                try:
                    shutil.move(EXCEL_PATH, EXCEL_PATH + ".corrupted.bak")
                except Exception:
                    os.remove(EXCEL_PATH)
    df = pd.DataFrame(BASE_PROJECTS)
    save_to_excel(df)
    return df

def excel_mtime() -> float:
    return os.path.getmtime(EXCEL_PATH) if os.path.exists(EXCEL_PATH) else 0.0


# ── USERS EXCEL HELPERS ───────────────────────────────────────────────────────
def _load_users_excel_passwords() -> dict:
    """Return {email_lower: plain_password} from users.xlsx if the file exists."""
    if not os.path.exists(USERS_EXCEL_PATH):
        return {}
    try:
        df = pd.read_excel(USERS_EXCEL_PATH, dtype=str, engine="openpyxl").fillna("")
        if "Email" in df.columns and "Password" in df.columns:
            return {str(r["Email"]).strip().lower(): str(r["Password"])
                    for _, r in df.iterrows() if str(r["Email"]).strip()}
    except Exception:
        pass
    return {}


def sync_users_excel(password_updates: dict = None):
    """Write users.xlsx (Name, Email, Role, Password, Active).
    password_updates = {email: plain_text_password} for newly set passwords."""
    users = auth.get_all_users()
    existing_pw = _load_users_excel_passwords()
    if password_updates:
        existing_pw.update({k.strip().lower(): v for k, v in password_updates.items()})
    rows = [
        {
            "Name": u["name"],
            "Email": u["email"],
            "Role": u["role"],
            "Password": existing_pw.get(u["email"].strip().lower(), ""),
            "Active": "Yes" if u["is_active"] else "No",
        }
        for u in users
    ]
    pd.DataFrame(rows).to_excel(USERS_EXCEL_PATH, index=False, engine="openpyxl")


def compute_roi(manual, auto, cost):
    try:
        m, a, c = float(manual), float(auto), float(cost)
        if m > 0:
            saved = max(0.0, m - a)
            return {"saved": saved, "cost": saved * c, "pct": round((saved / m) * 100)}
    except (ValueError, TypeError):
        pass
    return None

def get_api_key() -> str:
    try:
        return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        pass
    return os.environ.get("ANTHROPIC_API_KEY", "")

def is_new(row) -> bool:
    return str(row.get("is_new","")).lower() in ["true","1","yes"]

# ── HTML HELPERS ──────────────────────────────────────────────────────────────
esc = html.escape   # shorthand — always escape user-sourced values before HTML injection

def _parse_dmy(s: str):
    try: return datetime.strptime(str(s).strip(), "%d/%m/%Y").date()
    except: return None

def _parse_ymd(s: str):
    try: return datetime.strptime(str(s).strip(), "%Y-%m-%d").date()
    except: return None

def _due_cell(due_str: str) -> str:
    v = str(due_str).strip()
    if not v:
        return '<span style="font-size:10px;color:#CBD5E1">—</span>'
    d = _parse_dmy(v)
    if not d:
        return f'<span style="font-size:11px;color:#64748B">{esc(v)}</span>'
    diff = (d - date.today()).days
    if diff < 0:
        return f'<span style="font-size:10px;font-weight:700;background:#FEF2F2;color:#991B1B;padding:2px 5px;border-radius:4px">{esc(v)}</span>'
    if diff <= 7:
        return f'<span style="font-size:10px;font-weight:700;background:#FFFBEB;color:#92400E;padding:2px 5px;border-radius:4px">{esc(v)}</span>'
    return f'<span style="font-size:11px;color:#64748B">{esc(v)}</span>'

def badge_html(status: str) -> str:
    s = STATUS_STYLES.get(status, {"bg":"#F1F5F9","text":"#475569","dot":"#94A3B8"})
    return (f'<span style="display:inline-flex;align-items:center;gap:4px;padding:3px 10px;'
            f'border-radius:20px;font-size:11px;font-weight:700;background:{s["bg"]};color:{s["text"]}">'
            f'<span style="width:6px;height:6px;border-radius:50%;background:{s["dot"]};'
            f'display:inline-block"></span>{esc(status)}</span>')

def cell(val, size: str = "11px", color: str = "#374151") -> str:
    """Render a safe, consistently-styled table cell span."""
    return f'<span style="font-size:{size};color:{color}">{esc(str(val))}</span>'

def _inline_md(t: str) -> str:
    t = re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', t)
    t = re.sub(r'\*(.+?)\*', r'<em>\1</em>', t)
    t = re.sub(r'`([^`]+)`', r'<code style="background:#F1F5F9;padding:1px 4px;border-radius:3px;font-size:11px;font-family:monospace">\1</code>', t)
    return t

def _is_table_separator(line: str) -> bool:
    return bool(re.match(r'^\|[\s\-|:]+\|$', line.strip()))

def _parse_table_row(line: str) -> list:
    cells = line.strip().strip('|').split('|')
    return [c.strip() for c in cells]

def md_to_html(text: str) -> str:
    raw_lines = str(text).split('\n')
    out, in_list, i = [], False, 0
    while i < len(raw_lines):
        line = raw_lines[i]
        # ── Markdown table detection ──────────────────────────────────────────
        if (line.strip().startswith('|') and
                i + 1 < len(raw_lines) and _is_table_separator(raw_lines[i + 1])):
            if in_list: out.append('</ul>'); in_list = False
            header_cells = _parse_table_row(line)
            i += 2
            body_rows = []
            while i < len(raw_lines) and raw_lines[i].strip().startswith('|'):
                body_rows.append(_parse_table_row(raw_lines[i]))
                i += 1
            th = ''.join(
                f'<th style="padding:6px 12px;text-align:left;font-size:11px;'
                f'font-weight:700;text-transform:uppercase;letter-spacing:.4px;'
                f'color:#475569;background:#F1F5F9;border-bottom:2px solid #E2E8F0">'
                f'{html.escape(c)}</th>' for c in header_cells)
            rows_html = ''
            for ri, row in enumerate(body_rows):
                bg = '#ffffff' if ri % 2 == 0 else '#F8FAFC'
                td = ''.join(
                    f'<td style="padding:6px 12px;font-size:12px;color:#334155;'
                    f'border-bottom:1px solid #F1F5F9">{_inline_md(html.escape(c))}</td>'
                    for c in row)
                rows_html += f'<tr style="background:{bg}">{td}</tr>'
            out.append(
                f'<div style="overflow-x:auto;margin:8px 0">'
                f'<table style="width:100%;border-collapse:collapse;border:1px solid #E2E8F0;'
                f'border-radius:8px;overflow:hidden;font-family:inherit">'
                f'<thead><tr>{th}</tr></thead>'
                f'<tbody>{rows_html}</tbody>'
                f'</table></div>')
            continue
        # ── Everything else ───────────────────────────────────────────────────
        esc = html.escape(line)
        m = re.match(r'^(#{1,3}) (.+)$', esc)
        if m:
            if in_list: out.append('</ul>'); in_list = False
            sz = {1: '15px', 2: '14px', 3: '13px'}[len(m.group(1))]
            out.append(f'<div style="font-size:{sz};font-weight:700;margin:6px 0 2px">{_inline_md(m.group(2))}</div>')
        elif re.match(r'^[-*] ', esc):
            if not in_list: out.append('<ul style="margin:4px 0;padding-left:18px">'); in_list = True
            out.append(f'<li style="margin:2px 0">{_inline_md(esc[2:])}</li>')
        elif re.match(r'^\d+\. ', esc):
            if not in_list: out.append('<ul style="margin:4px 0;padding-left:18px">'); in_list = True
            out.append(f'<li style="margin:2px 0">{_inline_md(re.sub(r"^\d+[.] ", "", esc))}</li>')
        elif not esc.strip():
            if in_list: out.append('</ul>'); in_list = False
            out.append('<br>')
        else:
            if in_list: out.append('</ul>'); in_list = False
            out.append(_inline_md(esc) + '<br>')
        i += 1
    if in_list:
        out.append('</ul>')
    return ''.join(out)


# ── SESSION STATE ─────────────────────────────────────────────────────────────
if "projects" not in st.session_state:
    st.session_state.projects = load_from_excel()
if "is_active" not in st.session_state.projects.columns:
    st.session_state.projects["is_active"] = True
if "proj_type" not in st.session_state.projects.columns:
    st.session_state.projects["proj_type"] = ""
if "due_date" not in st.session_state.projects.columns:
    st.session_state.projects["due_date"] = ""
if "excel_mtime" not in st.session_state:
    st.session_state.excel_mtime = excel_mtime()
if "messages" not in st.session_state:
    st.session_state.messages = [{"role": "assistant", "content":
        "Hello! I'm your **AI Project Manager Agent**.\n\n"
        "I have live access to all **46 Qualesce projects** across Raychem, TEPL, "
        "Swagekklok-California, Swagelok-Alabama and internal/external POCs.\n\n"
        "Ask me anything about projects, team workload, status breakdown, or ROI!"}]
if "next_id" not in st.session_state:
    ids = pd.to_numeric(st.session_state.projects.get("id", pd.Series([])), errors="coerce").dropna()
    st.session_state.next_id = int(ids.max()) + 1 if not ids.empty else max(r["id"] for r in BASE_PROJECTS) + 1
if "active_tab"           not in st.session_state: st.session_state.active_tab           = "dashboard"
if "dash_slicer"          not in st.session_state: st.session_state.dash_slicer          = None
if "show_modal"           not in st.session_state: st.session_state.show_modal           = None
if "confirm_delete"       not in st.session_state: st.session_state.confirm_delete       = None
if "toast"                not in st.session_state: st.session_state.toast                = None
if "dismissed_notifs"     not in st.session_state: st.session_state.dismissed_notifs     = set()
if "show_notif_detail"    not in st.session_state: st.session_state.show_notif_detail    = None
if "project_filter_preset"  not in st.session_state: st.session_state.project_filter_preset  = "All"
if "presales_filter_preset" not in st.session_state: st.session_state.presales_filter_preset = "All"
if "lc_edit_id"            not in st.session_state: st.session_state.lc_edit_id            = None
if "sl_edit_id"            not in st.session_state: st.session_state.sl_edit_id            = None
if "dash_client_filter"    not in st.session_state: st.session_state.dash_client_filter    = "All"
if "dash_slicers_expanded" not in st.session_state: st.session_state.dash_slicers_expanded = False
if "current_user"         not in st.session_state: st.session_state.current_user         = None
if "reset_pwd_uid"        not in st.session_state: st.session_state.reset_pwd_uid        = None
if "user_edit_id"         not in st.session_state: st.session_state.user_edit_id         = None
if "task_comment_view" not in st.session_state: st.session_state.task_comment_view = None
if "poc_row_edit"     not in st.session_state: st.session_state.poc_row_edit     = None

# ── HELPERS ───────────────────────────────────────────────────────────────────
def get_stats(d):
    new_mask = d["is_new"].astype(str).str.lower().isin(["true","1","yes"]) if "is_new" in d.columns else pd.Series([False]*len(d))
    hrs  = pd.to_numeric(d.get("hours_saved", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    cost = pd.to_numeric(d.get("cost_saved",  pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
    def c(s): return int(d["status"].str.contains(s, na=False).sum())
    return dict(total=len(d), rm=c("R&M"), uat=c("UAT"), completed=c("Completed"),
                in_progress=c("In Progress"), poc=c("POC"), pdd=c("PDD"),
                discontinued=c("Discontinued"), important=c("Important"), new_added=int(new_mask.sum()),
                total_hrs=float(hrs), total_cost=float(cost))

def call_claude(api_key, msgs, df):
    client = anthropic.Anthropic(api_key=api_key)
    proj_ctx = "\n\nLIVE PROJECT DATA:\n" + "\n".join(
        f"- {r['name']} | {r['client']} | {r['employee']} | {r['status']}"
        for _, r in df.iterrows())
    api_msgs = [{"role": m["role"], "content": m["content"]} for m in msgs[-12:]]
    while api_msgs and api_msgs[0]["role"] != "user":
        api_msgs = api_msgs[1:]
    if not api_msgs:
        return "Please ask me a question to get started!"
    resp = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=500,
        system=SYSTEM_PROMPT + proj_ctx,
        messages=api_msgs)
    return resp.content[0].text

# ── STYLES ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
html,body,[class*="css"]{
  font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','Inter',sans-serif!important;
  background:#F1F5F9!important;
  color:#1E293B!important;
  -webkit-font-smoothing:antialiased!important;
  -moz-osx-font-smoothing:grayscale!important;
  text-rendering:optimizeLegibility!important;
}
#MainMenu,footer,header{visibility:hidden}
.block-container{padding:0 1.5rem 2rem!important;max-width:100%!important}
section[data-testid="stSidebar"]{display:none!important}

/* ── KPI Cards ── */
.kpi-wrap{
  text-align:center;padding:18px 12px;border-radius:12px;background:#FFFFFF;
  border:1px solid #E2E8F0;
  box-shadow:0 2px 8px rgba(15,23,42,.07);
  cursor:pointer}
.kpi-wrap:hover{box-shadow:0 6px 20px rgba(15,23,42,.12)}
.kpi-num{font-family:'Courier New',Courier,monospace;font-size:28px;font-weight:700;margin:8px 0 4px;letter-spacing:-1px}
.kpi-lbl{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#94A3B8}

/* ── Top Navigation ── */
.q-nav{
  background:#0F172A;
  padding:0 28px;
  display:flex;align-items:center;justify-content:space-between;
  height:62px;position:sticky;top:0;z-index:100;
  box-shadow:0 2px 12px rgba(0,0,0,.30);
  margin:0 -1.5rem 24px}

/* ── Slicer rows ── */
.srow{
  padding:12px 16px;border-bottom:1px solid #F1F5F9;
  display:flex;justify-content:space-between;align-items:start}
.srow:nth-child(even){background:#F8FAFC}
.srow:hover{background:#EFF6FF}

/* ── Project table rows ── */
.prow{padding:10px;border-bottom:1px solid #F1F5F9;background:#fff}
.prow:nth-child(even){background:#F8FAFC}
.prow:hover{background:#EFF6FF}

/* ── Chat bubbles ── */
@keyframes slideInRight{from{opacity:0;transform:translateX(30px)}to{opacity:1;transform:translateX(0)}}
@keyframes slideInLeft{from{opacity:0;transform:translateX(-30px)}to{opacity:1;transform:translateX(0)}}
@keyframes typingPulse{0%,80%,100%{transform:scale(0);opacity:.3}40%{transform:scale(1);opacity:1}}
.chat-row{display:flex;align-items:flex-end;gap:8px;margin:8px 0}
.chat-row.user-row{flex-direction:row-reverse}
.chat-avatar{
  width:30px;height:30px;border-radius:50%;display:flex;align-items:center;
  justify-content:center;font-size:14px;flex-shrink:0;font-weight:700}
.avatar-user{background:#DBEAFE;color:#1D4ED8}
.avatar-bot{background:#DCFCE7;color:#16A34A}
.chat-user{
  background:#EFF6FF;border:1px solid #BFDBFE;
  border-radius:16px 16px 4px 16px;padding:12px 16px;font-size:13px;line-height:1.6;
  max-width:80%;animation:slideInRight .3s ease-out}
.chat-bot{
  background:#F0FDF4;border:1px solid #BBF7D0;
  border-radius:16px 16px 16px 4px;padding:12px 16px;font-size:13px;line-height:1.6;
  max-width:80%;animation:slideInLeft .3s ease-out}
.typing-indicator{
  display:flex;align-items:center;gap:10px;
  background:#F0FDF4;border:1px solid #BBF7D0;
  border-radius:16px 16px 16px 4px;padding:12px 16px;
  width:fit-content;animation:slideInLeft .3s ease-out}
.typing-dots{display:flex;gap:4px;align-items:center}
.typing-dots span{
  width:7px;height:7px;border-radius:50%;background:#16A34A;display:inline-block}
.typing-dots span:nth-child(1){animation:typingPulse 1.2s infinite ease-in-out}
.typing-dots span:nth-child(2){animation:typingPulse 1.2s infinite ease-in-out .2s}
.typing-dots span:nth-child(3){animation:typingPulse 1.2s infinite ease-in-out .4s}

/* ── ROI Banner ── */
.roi-banner{
  background:linear-gradient(135deg,#0F2D52,#1E3A5F);
  border:1px solid rgba(37,99,235,.35);
  border-radius:12px;padding:18px 26px;
  display:flex;gap:32px;align-items:center;margin-bottom:20px;
  box-shadow:0 4px 16px rgba(15,23,42,.18)}

/* ── Streamlit buttons ── */
div[data-testid="stButton"] > button{
  border-radius:8px!important;
  font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif!important;
  font-weight:600!important;
  font-size:12px!important;
  letter-spacing:.2px!important}

/* ── Notification popup ── */
.notif-popup{
  border-radius:12px;padding:20px 24px;margin-bottom:20px;
  box-shadow:0 4px 24px rgba(15,23,42,.12);}
.notif-alert{}

/* ── Streamlit container borders ── */
div[data-testid="stVerticalBlockBorderWrapper"]{
  border-color:#E2E8F0!important;border-radius:12px!important}

/* ── Table action icon buttons (✏ edit · 🗑 delete · 🔑 warn) ── */
div[data-testid="stMarkdownContainer"]:has(.act-edit-marker) ~ div[data-testid="stButton"] > button{
  background:#EFF6FF!important;border:1.5px solid #BFDBFE!important;
  color:#1D4ED8!important;font-size:15px!important;
  min-height:30px!important;padding:2px 8px!important;
  transition:background .15s,border-color .15s!important}
div[data-testid="stMarkdownContainer"]:has(.act-edit-marker) ~ div[data-testid="stButton"] > button:hover{
  background:#DBEAFE!important;border-color:#93C5FD!important}
div[data-testid="stMarkdownContainer"]:has(.act-del-marker) ~ div[data-testid="stButton"] > button{
  background:#FFF1F2!important;border:1.5px solid #FECACA!important;
  color:#DC2626!important;font-size:15px!important;
  min-height:30px!important;padding:2px 8px!important;
  transition:background .15s,border-color .15s!important}
div[data-testid="stMarkdownContainer"]:has(.act-del-marker) ~ div[data-testid="stButton"] > button:hover{
  background:#FEE2E2!important;border-color:#FCA5A5!important}
div[data-testid="stMarkdownContainer"]:has(.act-warn-marker) ~ div[data-testid="stButton"] > button{
  background:#FFFBEB!important;border:1.5px solid #FDE68A!important;
  color:#92400E!important;font-size:15px!important;
  min-height:30px!important;padding:2px 8px!important;
  transition:background .15s,border-color .15s!important}
div[data-testid="stMarkdownContainer"]:has(.act-warn-marker) ~ div[data-testid="stButton"] > button:hover{
  background:#FEF3C7!important;border-color:#FCD34D!important}

/* ── Login ── */
.login-hint{text-align:center;font-size:11px;color:#94A3B8;margin-top:12px}

/* ── Task progress bar ── */
.progress-bar-outer{background:#E2E8F0;border-radius:10px;height:7px;overflow:hidden;margin:4px 0}
.progress-bar-inner{height:7px;border-radius:10px}

/* ── Role badge ── */
.role-badge{font-size:9px;font-weight:700;padding:2px 8px;border-radius:10px;text-transform:uppercase}

/* ── KPI expand animation ── */
@keyframes kpi-slide-in{
  0%  {opacity:0;transform:translateY(-14px) scale(0.90)}
  60% {opacity:1;transform:translateY(2px)   scale(1.02)}
  100%{opacity:1;transform:translateY(0)     scale(1)}
}
.kpi-anim{animation:kpi-slide-in .38s cubic-bezier(0.34,1.3,0.64,1) both}

/* ── Expand arrow button ── */
.expand-btn button{
  border-radius:50%!important;
  width:38px!important;height:38px!important;
  padding:0!important;font-size:16px!important;
  background:#F1F5F9!important;border:1px solid #CBD5E1!important;
  color:#475569!important;font-weight:700!important}
.expand-btn button:hover{background:#E2E8F0!important}

/* ── HD table ── */
.hd-table{width:100%;border-collapse:collapse;font-size:12px}
.hd-table th{font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.7px;color:#94A3B8;padding:6px 8px;border-bottom:2px solid #E2E8F0;white-space:nowrap}
.hd-table td{padding:8px 8px;border-bottom:1px solid #F1F5F9;vertical-align:middle;color:#374151}
.hd-table tr:hover td{background:#F0F4FF}
.hd-table tr:nth-child(even) td{background:#F8FAFC}
.hd-table tr:nth-child(even):hover td{background:#EFF6FF}
</style>
""", unsafe_allow_html=True)

# ── LOGIN GATE ───────────────────────────────────────────────────────────────
def _render_login():
    _, col, _ = st.columns([1, 1.2, 1])
    with col:
        with st.container(border=True):
            st.markdown("""
            <div style="text-align:center;padding:16px 0 20px">
              <div style="font-size:36px;font-weight:900;color:#3B82F6;letter-spacing:-2px;margin-bottom:8px">Q</div>
              <div style="font-size:20px;font-weight:800;color:#0F172A;letter-spacing:-.3px">QUALESCE</div>
              <div style="font-size:12px;color:#64748B;margin-top:4px">AI Project Manager Platform</div>
            </div>
            """, unsafe_allow_html=True)
            with st.form("login_form"):
                email    = st.text_input("Email Address", placeholder="you@company.com")
                password = st.text_input("Password", type="password", placeholder="••••••••")
                submitted = st.form_submit_button("Sign In", use_container_width=True, type="primary")
            if submitted:
                if not email.strip() or not password:
                    st.error("Email and password are required.")
                else:
                    user = auth.authenticate(email, password)
                    if user:
                        st.session_state.current_user = user
                        st.session_state.active_tab   = "tasks" if user["role"] == "employee" else "dashboard"
                        st.rerun()
                    else:
                        st.error("Invalid credentials or account is inactive.")
            st.markdown('<div class="login-hint"> </div>',
                        unsafe_allow_html=True)

if st.session_state.current_user is None:
    _render_login()
    st.stop()

cu   = st.session_state.current_user
role = cu["role"]

# ── NAV ───────────────────────────────────────────────────────────────────────
df    = st.session_state.projects
stats = get_stats(df)

_new_badge = f"&nbsp;<span style='color:#34D399;font-weight:600'>+{stats['new_added']} new</span>" if stats["new_added"] else ""
st.markdown(
    f'<div class="q-nav">'
    f'<div style="display:flex;align-items:center;gap:14px">'
    f'<div style="width:38px;height:38px;background:linear-gradient(135deg,#3B82F6,#6366F1);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:14px;font-weight:800;color:#fff;letter-spacing:-0.5px;box-shadow:0 0 0 1px rgba(255,255,255,.12)">Q</div>'
    f'<div>'
    f'<div style="font-family:\'JetBrains Mono\',monospace;font-weight:700;font-size:13px;color:#F1F5F9;letter-spacing:2px;text-transform:uppercase">QUALESCE</div>'
    f'<div style="font-size:9px;color:#94A3B8;letter-spacing:1.2px;text-transform:uppercase;font-weight:500;margin-top:1px">AI Project Manager</div>'
    f'</div>'
    f'</div>'
    f'<div style="font-size:12px;color:#94A3B8;display:flex;align-items:center;gap:10px">'
    f'<span style="width:7px;height:7px;border-radius:50%;background:#10B981;box-shadow:0 0 8px #10B981;display:inline-block"></span>'
    f'<b style="color:#E2E8F0;font-weight:600">{stats["total"]}</b>'
    f'<span>projects live</span>'
    f'{_new_badge}'
    f'&nbsp;<span style="color:#475569">|</span>&nbsp;'
    f'<span style="color:#E2E8F0;font-weight:600">{esc(cu["name"])}</span>'
    f'<span style="background:#1E3A8A;color:#93C5FD;font-size:9px;font-weight:700;padding:2px 8px;border-radius:10px;text-transform:uppercase">{esc(cu["role"])}</span>'
    f'</div>'
    f'</div>',
    unsafe_allow_html=True
)

# ── TOAST ─────────────────────────────────────────────────────────────────────
if st.session_state.toast:
    t = st.session_state.toast
    colors = {"success": ("#064E3B","#10B981"), "error": ("#7F1D1D","#EF4444"), "info": ("#1E3A8A","#3B82F6")}
    bg, border = colors.get(t.get("type","success"), ("#064E3B","#10B981"))
    st.markdown(f'<div style="background:{bg};border:1px solid {border};border-radius:10px;'
                f'padding:11px 18px;color:#fff;font-size:13px;font-weight:600;margin-bottom:12px">'
                f'{esc(t["msg"])}</div>', unsafe_allow_html=True)
    st.session_state.toast = None

# ── TOP BAR: TABS + ACTIONS ───────────────────────────────────────────────────
if role == "employee":
    _tab_defs = [("tasks", "My Tasks")]
elif role == "sales":
    _tab_defs = [("dashboard", "Dashboard"), ("presales", "Presales/POC")]
elif role in ("lead", "manager"):
    _tab_defs = [("dashboard", "Dashboard"), ("projects", "Projects"),
                 ("presales", "Presales/POC"), ("license", "License"),
                 ("agent", "AI Agent"), ("tasks", "Tasks")]
else:
    _tab_defs = [("dashboard", "Dashboard"), ("projects", "Projects"),
                 ("presales", "Presales/POC"), ("license", "License"),
                 ("agent", "AI Agent"), ("users", "Users"), ("tasks", "Tasks")]

if st.session_state.active_tab not in [t[0] for t in _tab_defs]:
    st.session_state.active_tab = _tab_defs[0][0]

_n = len(_tab_defs)
if role == "admin":
    nav_c = st.columns([1] * _n + [0.9, 0.6, 0.55])
elif role in ("lead", "manager"):
    nav_c = st.columns([1] * _n + [0.55])
else:
    nav_c = st.columns([1] * _n + [0.55])

for _i, (_tid, _tlabel) in enumerate(_tab_defs):
    _active = st.session_state.active_tab == _tid
    _badge  = f" +{stats['new_added']}" if _tid == "projects" and stats.get("new_added") else ""
    if nav_c[_i].button(f"{_tlabel}{_badge}", key=f"tab_{_tid}",
                        type="primary" if _active else "secondary",
                        use_container_width=True):
        st.session_state.active_tab = _tid
        st.rerun()

if role == "admin":
    if nav_c[_n].button("Add Project", type="primary", use_container_width=True):
        st.session_state.show_modal = "add"
        st.rerun()
    if nav_c[_n + 1].button("Sync", use_container_width=True):
        st.session_state.projects = load_from_excel()
        st.session_state.excel_mtime = excel_mtime()
        ids = pd.to_numeric(st.session_state.projects.get("id", pd.Series([])), errors="coerce").dropna()
        st.session_state.next_id = int(ids.max()) + 1 if not ids.empty else max(r["id"] for r in BASE_PROJECTS) + 1
        st.session_state.toast = {"msg": "Synced from Excel!", "type": "success"}
        st.rerun()
    if nav_c[_n + 2].button("Logout", use_container_width=True):
        st.session_state.current_user = None
        st.rerun()
elif role in ("lead", "manager"):
    if nav_c[_n].button("Logout", use_container_width=True):
        st.session_state.current_user = None
        st.rerun()
else:
    if nav_c[_n].button("Logout", use_container_width=True):
        st.session_state.current_user = None
        st.rerun()

if excel_mtime() != st.session_state.excel_mtime:
    st.warning("Excel file changed externally — click **Sync Excel** to reload.")

st.markdown("---")
df = st.session_state.projects   # re-bind after possible sync
_HDR_STYLE = 'font-size:9px;font-weight:700;text-transform:uppercase;color:#94A3B8;letter-spacing:.7px;padding:5px 0;border-bottom:2px solid #E2E8F0'

# ══════════════════════════════════════════════════════════════════════════════
# MODAL: ADD / EDIT
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.show_modal is not None and role in ("admin", "lead", "manager"):
    mode     = "add" if st.session_state.show_modal == "add" else "edit"
    edit_row = {} if mode == "add" else st.session_state.show_modal.get("edit", {})

    # Build sorted unique employee list from current data
    all_employees = sorted(set(
        n.strip()
        for raw in st.session_state.projects.get("employee", pd.Series(dtype=str)).dropna()
        for n in str(raw).replace("&", ",").split(",")
        if n.strip()
    ))
    # Include leads in the employee pool for lead selection
    if "lead" in st.session_state.projects.columns:
        all_employees = sorted(set(all_employees) | set(
            str(l).strip()
            for l in st.session_state.projects["lead"].dropna()
            if str(l).strip()
        ))
    # Build sorted unique client list from current data
    all_clients = sorted(set(
        str(c).strip()
        for c in st.session_state.projects.get("client", pd.Series(dtype=str)).dropna()
        if str(c).strip()
    ))
    EMP_NEW    = "── Type new name ──"
    CLIENT_NEW = "── Type new client ──"
    client_options = all_clients + [CLIENT_NEW]

    title = "Add New Project" if mode == "add" else "Edit Project"
    st.markdown(f"### {title}")
    with st.container(border=True):
        c1, c2 = st.columns(2)
        name = c1.text_input("Project Name *", value=edit_row.get("name",""))

        # Client: searchable selectbox + optional free-text override
        current_client = edit_row.get("client","")
        client_idx     = client_options.index(current_client) if current_client in client_options else len(client_options) - 1
        client_select  = c2.selectbox(
            "Client * (search or select)",
            options=client_options,
            index=client_idx,
            help="Start typing to search existing clients. Choose the last option to enter a new client."
        )
        if client_select == CLIENT_NEW:
            client = c2.text_input("Enter new client name *", value="", placeholder="e.g. Acme Corp")
        else:
            client = client_select

        # Lead: searchable selectbox (single person — project lead)
        lead_options_full = [""] + all_employees + [EMP_NEW]
        current_lead = edit_row.get("lead", "")
        lead_idx = lead_options_full.index(current_lead) if current_lead in lead_options_full else 0
        lead_select = c1.selectbox(
            "Lead (search or select)",
            options=lead_options_full,
            index=lead_idx,
            help="Select the project lead. Start typing to search existing team members."
        )
        if lead_select == EMP_NEW:
            lead = c1.text_input("Enter new lead name", value="", placeholder="e.g. Jane Smith")
        else:
            lead = lead_select

        idx    = ALL_STATUSES.index(edit_row["status"]) if edit_row.get("status") in ALL_STATUSES else 0
        status = c2.selectbox("Status", ALL_STATUSES, index=idx)

        _PROJ_TYPES = ["", "RPA", "AI Agent", "Presales"]
        _pt_val = edit_row.get("proj_type", "")
        _pt_idx = _PROJ_TYPES.index(_pt_val) if _pt_val in _PROJ_TYPES else 0
        proj_type = c1.selectbox("Type", _PROJ_TYPES, index=_pt_idx,
                                 format_func=lambda x: "— Select type —" if x == "" else x)

        # Employees: multi-select — one or more team members assigned to the project
        current_emp_raw  = str(edit_row.get("employee",""))
        current_emp_list = [n.strip() for n in current_emp_raw.replace("&", ",").split(",") if n.strip()]
        valid_emp_defaults = [e for e in current_emp_list if e in all_employees]
        selected_emps = st.multiselect(
            "Employees * (select one or more)",
            options=all_employees,
            default=valid_emp_defaults,
            help="Search and select all team members assigned to this project."
        )
        new_emp_name = st.text_input(
            "Add new employee name (optional)",
            value="",
            placeholder="e.g. John Doe — leave blank if not needed"
        )
        if new_emp_name.strip():
            emp = ", ".join(selected_emps + [new_emp_name.strip()])
        else:
            emp = ", ".join(selected_emps)

        _dc1, _dc2, _dc3 = st.columns(3)
        _s_default = _parse_dmy(edit_row.get("start", ""))
        _start_dt = _dc1.date_input("Start Date", value=_s_default, key="modal_start", format="DD/MM/YYYY")
        start = _start_dt.strftime("%d/%m/%Y") if _start_dt else ""

        _e_raw = edit_row.get("end", "").strip()
        _is_ongoing = not bool(_e_raw)
        _ongoing = _dc2.checkbox("Ongoing (no end date)", value=_is_ongoing, key="modal_ongoing")
        if _ongoing:
            end = ""
        else:
            _e_default = _parse_dmy(_e_raw) or date.today()
            _end_dt = _dc2.date_input("End Date", value=_e_default, key="modal_end", format="DD/MM/YYYY")
            end = _end_dt.strftime("%d/%m/%Y") if _end_dt else ""

        _d_raw = edit_row.get("due_date", "").strip()
        _due_dt_default = _parse_dmy(_d_raw) if _d_raw else None
        _due_dt = _dc3.date_input("Due Date (optional)", value=_due_dt_default, key="modal_due", format="DD/MM/YYYY")
        due_date = _due_dt.strftime("%d/%m/%Y") if _due_dt else ""

        po     = c1.text_input("PO Number",           value=edit_row.get("po",""))
        desc   = c2.text_input("Description",         value=edit_row.get("desc",""))
        _is_active_raw = str(edit_row.get("is_active", "True")).strip().lower()
        is_active = c1.checkbox("Active", value=(_is_active_raw not in ["false","0","no"]))

        st.markdown("**ROI Calculator** *(optional)*")
        r1, r2, r3 = st.columns(3)
        manual_hrs  = r1.text_input("Manual Hrs",  value=edit_row.get("manual_hrs",""))
        auto_hrs    = r2.text_input("Auto Hrs",    value=edit_row.get("auto_hrs",""))
        cost_per_hr = r3.text_input("Cost/Hr (₹)", value=edit_row.get("cost_per_hr",""))

        roi = compute_roi(manual_hrs, auto_hrs, cost_per_hr)
        if roi:
            st.success(f"ROI: **{roi['pct']}%** | Hrs Saved: **{roi['saved']}** | Cost Saved: **₹{roi['cost']:,.0f}**")

        s1, s2 = st.columns(2)
        save_clicked   = s1.button("Save",   type="primary", use_container_width=True, key="modal_save")
        cancel_clicked = s2.button("Cancel",  use_container_width=True, key="modal_cancel")

        if cancel_clicked:
            st.session_state.show_modal = None
            st.rerun()

        if save_clicked:
            errors = []
            if not name or len(name.strip()) < 3:  errors.append("Project name must be at least 3 characters.")
            if not client.strip():                  errors.append("Client is required.")
            if not emp.strip():                     errors.append("Employee is required.")
            if errors:
                for e in errors: st.error(e)
            else:
                if mode == "add":
                    new_row = {
                        "id": st.session_state.next_id,
                        "name": name.strip(), "client": client.strip(),
                        "lead": lead.strip(), "employee": emp.strip(),
                        "status": status, "proj_type": proj_type,
                        "start": start, "end": end, "due_date": due_date, "po": po, "desc": desc.strip(),
                        "manual_hrs": manual_hrs, "auto_hrs": auto_hrs, "cost_per_hr": cost_per_hr,
                        "hours_saved": str(roi["saved"]) if roi else "",
                        "cost_saved":  str(roi["cost"])  if roi else "",
                        "roi_pct":     str(roi["pct"])   if roi else "",
                        "is_new": True,
                        "is_active": is_active,
                    }
                    st.session_state.projects = pd.concat(
                        [st.session_state.projects, pd.DataFrame([new_row])], ignore_index=True)
                    st.session_state.next_id += 1
                    roi_line = f" | ROI {roi['pct']}%" if roi else ""
                    st.session_state.messages.append({"role":"user","content":
                        f"New project added: {name} | {client} | {emp} | {status}{roi_line}. Confirm and give a brief health insight."})
                    st.session_state.toast = {"msg": f'"{name}" added!', "type": "success"}
                else:
                    eid = str(edit_row.get("id",""))
                    records = []
                    for r in st.session_state.projects.to_dict("records"):
                        if str(r.get("id","")) == eid:
                            r.update({"name":name.strip(),"client":client.strip(),
                                      "lead":lead.strip(),"employee":emp.strip(),
                                      "status":status,"proj_type":proj_type,
                                      "start":start,"end":end,"due_date":due_date,"po":po,"desc":desc.strip(),
                                      "manual_hrs":manual_hrs,"auto_hrs":auto_hrs,"cost_per_hr":cost_per_hr,
                                      "hours_saved":str(roi["saved"]) if roi else r.get("hours_saved",""),
                                      "cost_saved": str(roi["cost"])  if roi else r.get("cost_saved",""),
                                      "roi_pct":    str(roi["pct"])   if roi else r.get("roi_pct",""),
                                      "is_active":  is_active})
                        records.append(r)
                    st.session_state.projects = pd.DataFrame(records)
                    st.session_state.toast = {"msg": f'"{name}" updated!', "type": "success"}

                save_to_excel(st.session_state.projects)
                st.session_state.excel_mtime = excel_mtime()
                st.session_state.show_modal = None
                st.rerun()

    st.markdown("---")

# ── CONFIRM DELETE ────────────────────────────────────────────────────────────
if st.session_state.confirm_delete and role == "admin":
    cd = st.session_state.confirm_delete
    st.warning(f"Delete \"{cd['name']}\"? This cannot be undone.")
    da, db, _ = st.columns([1,1,4])
    if da.button("Yes, Delete", type="primary", use_container_width=True, key="yes_del"):
        # Delete by id (not name) to avoid deleting two projects with the same name
        st.session_state.projects = st.session_state.projects[
            st.session_state.projects["id"].astype(str) != str(cd["id"])].reset_index(drop=True)
        save_to_excel(st.session_state.projects)
        st.session_state.excel_mtime = excel_mtime()
        st.session_state.messages.append({"role":"assistant",
            "content": f'"{cd["name"]}" removed. Dashboard updated.'})
        st.session_state.toast = {"msg": f'"{cd["name"]}" deleted.', "type": "info"}
        st.session_state.confirm_delete = None
        st.rerun()
    if db.button("Cancel", use_container_width=True, key="no_del"):
        st.session_state.confirm_delete = None
        st.rerun()
    st.markdown("---")

df = st.session_state.projects

# ── Jinja2 chat templates ─────────────────────────────────────────────────────
_TMPL_USER_MSG = Template("""
<div class="chat-row user-row">
  <div class="chat-avatar avatar-user">U</div>
  <div class="chat-user">{{ content }}</div>
</div>
""")

_TMPL_BOT_MSG = Template("""
<div class="chat-row">
  <div class="chat-avatar avatar-bot">Q</div>
  <div class="chat-bot">{{ content }}</div>
</div>
""")

_TMPL_TYPING = Template("""
<div class="chat-row">
  <div class="chat-avatar avatar-bot">Q</div>
  <div class="typing-indicator">
    <div class="typing-dots"><span></span><span></span><span></span></div>
  </div>
</div>
""")

# ══════════════════════════════════════════════════════════════════════════════
# TAB: DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.active_tab == "dashboard" and role not in ("employee",):
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#0F172A;margin-bottom:4px;letter-spacing:-.3px">Project Portfolio Dashboard</h2>', unsafe_allow_html=True)
    st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:10px">Click any status card to drill into projects &amp; team members</p>', unsafe_allow_html=True)

    # ── CLIENT FILTER (top, under heading) ───────────────────────────────────
    _all_clients = sorted(set(
        str(c).strip() for c in df["client"].dropna() if str(c).strip()
    )) if "client" in df.columns else []
    _cf_col1, _cf_col2 = st.columns([2, 5])
    with _cf_col1:
        _sel_client = st.selectbox(
            "Filter by Client",
            options=["All"] + _all_clients,
            index=(["All"] + _all_clients).index(st.session_state.dash_client_filter)
                  if st.session_state.dash_client_filter in (["All"] + _all_clients) else 0,
            key="dash_client_select",
            help="Filter dashboard by client."
        )
    if _sel_client != st.session_state.dash_client_filter:
        st.session_state.dash_client_filter = _sel_client
        st.rerun()
    if st.session_state.dash_client_filter != "All":
        _cf_col2.markdown(
            f'<div style="background:#EFF6FF;border:1px solid #BFDBFE;border-radius:8px;'
            f'padding:8px 14px;font-size:12px;color:#1D4ED8;font-weight:600;margin-top:4px">'
            f'Showing projects for <b>{st.session_state.dash_client_filter}</b>'
            f'</div>',
            unsafe_allow_html=True
        )

    # Pre-compute client-filtered df so KPI cards and all panels reflect the filter
    _dash_df_pre = df.copy()
    if st.session_state.dash_client_filter != "All" and "client" in _dash_df_pre.columns:
        _dash_df_pre = _dash_df_pre[
            _dash_df_pre["client"].str.strip() == st.session_state.dash_client_filter
        ]
    stats = get_stats(_dash_df_pre)

    # ── KPI SLICER CARDS ──────────────────────────────────────────────────────
    _dev_statuses = {"In Progress"}
    _active_dev_mask = _dash_df_pre["status"].isin(_dev_statuses)
    if "is_active" in _dash_df_pre.columns:
        _active_dev_mask = _active_dev_mask & (
            ~_dash_df_pre["is_active"].astype(str).str.strip().str.lower().isin(["false","0","no"])
        )
    _active_dev_count = int(_active_dev_mask.sum())

    _kpi_extra = [
        ("R&M",         stats["rm"],          "RM", "#3B82F6", "R&M"),
        ("UAT",         stats["uat"],         "UA", "#F59E0B", "UAT"),
        ("Completed",   stats["completed"],   "CP", "#10B981", "Completed"),
        ("In Progress", stats["in_progress"], "IP", "#06B6D4", "In Progress"),
    ]

    def _kpi_card(col, label, val, icon, color, key, compact=False, animate=False, anim_delay=0):
        active     = st.session_state.dash_slicer == key
        bg         = f"linear-gradient(135deg,{color}18,{color}08)" if active else "#FFFFFF"
        border     = f"2px solid {color}" if active else "1px solid #E2E8F0"
        shadow     = f"0 6px 20px {color}44" if active else "0 4px 14px rgba(15,23,42,.08)"
        _dot       = f"<div style='width:7px;height:7px;border-radius:50%;background:{color};margin:4px auto 0;box-shadow:0 0 6px {color}'></div>" if active else ""
        anim_class = "kpi-anim" if animate else ""
        anim_style = f"animation-delay:{anim_delay}s;" if animate and anim_delay else ""
        pad        = "padding:8px 10px;" if compact else ""
        col.markdown(
            f'<div class="kpi-wrap {anim_class}" style="background:{bg};border:{border};box-shadow:{shadow};{pad}{anim_style}">'
            f'<div style="font-size:{"13px" if compact else "20px"};font-weight:800;color:{color}">{icon}</div>'
            f'<div class="kpi-num" style="color:{color};{"font-size:20px;" if compact else ""}">{val}</div>'
            f'<div class="kpi-lbl">{label}</div>'
            f'{_dot}'
            f'</div>',
            unsafe_allow_html=True
        )
        if col.button("✓" if (active and compact) else ("✓ Active" if active else ("▼ Filters" if compact else "Filter")),
                      key=f"kpi_{label}", use_container_width=True,
                      type="primary" if active else "secondary"):
            st.session_state.dash_slicer = None if active else key
            st.rerun()

    # ── Fixed 6-column row (structure never changes — prevents layout shift) ───
    _row = st.columns([0.9, 0.16, 0.9, 0.9, 0.9, 0.9])

    # "All" card — same size as other slicers
    _kpi_card(_row[0], "All", _active_dev_count, "ALL", "#3B82F6", "__active_dev__",
              compact=False, animate=False)

    # Arrow button — centred vertically
    _row[1].markdown("<div style='height:22px'></div>", unsafe_allow_html=True)
    _arrow_label = "▼" if st.session_state.dash_slicers_expanded else "▶"
    if _row[1].button(_arrow_label, key="expand_slicers",
                      help="Show / hide filter slicers",
                      use_container_width=True):
        st.session_state.dash_slicers_expanded = not st.session_state.dash_slicers_expanded
        st.rerun()

    # Extra slicers — cols 2-5 stay in layout but only get content when expanded
    if st.session_state.dash_slicers_expanded:
        for _i, (_col, (_lbl, _val, _ico, _clr, _key)) in enumerate(zip(_row[2:], _kpi_extra)):
            _kpi_card(_col, _lbl, _val, _ico, _clr, _key,
                      animate=True, anim_delay=round(_i * 0.08, 2))

    st.markdown("<br>", unsafe_allow_html=True)

    # Use pre-computed client-filtered df for all panels below
    dash_df    = _dash_df_pre
    dash_stats = stats

    st.markdown("<br>", unsafe_allow_html=True)

    # ── NOTIFICATION ALERT PANEL ──────────────────────────────────────────────
    NOTIF_DEFS = [
        {
            "key":    "Important",
            "label":  "Important Tasks",
            "icon":   "!",
            "color":  "#F43F5E",
            "bg":     "#FFF1F2",
            "border": "#FDA4AF",
            "note":   "High-priority tasks requiring immediate attention",
        },
    ]

    # Compute project lists for each alert type (respects lead filter)
    def get_alert_projects(status_key):
        mask = dash_df["status"].str.contains(status_key, na=False)
        return dash_df[mask]

    active_notifs = [n for n in NOTIF_DEFS if n["key"] not in st.session_state.dismissed_notifs]
    notif_data    = {n["key"]: get_alert_projects(n["key"]) for n in active_notifs}
    visible_notifs = [n for n in active_notifs if len(notif_data[n["key"]]) > 0]

    if visible_notifs:
        total_important = len(notif_data.get("Important", pd.DataFrame()))
        if total_important > 0 and "Important" not in st.session_state.dismissed_notifs:
            st.markdown(
                f'<div class="notif-alert" style="background:#FFF1F2;border:2px solid #F43F5E;'
                f'border-radius:10px;padding:10px 16px;display:flex;align-items:center;'
                f'gap:10px;margin-bottom:8px">'
                f'<span style="font-size:13px;font-weight:900;color:#F43F5E;padding:2px 7px;background:#FEE2E2;border-radius:6px">!</span>'
                f'<span style="font-weight:800;color:#BE123C;font-size:13px">ALERT:</span>'
                f'<span style="color:#9F1239;font-size:12px">'
                f'<b>{total_important}</b> project(s) marked as <b>Important</b> require immediate attention!</span>'
                f'</div>',
                unsafe_allow_html=True
            )

        notif_cols = st.columns(len(visible_notifs))
        for col, notif in zip(notif_cols, visible_notifs):
            proj_list   = notif_data[notif["key"]]
            proj_count  = len(proj_list)
            preview     = proj_list["name"].head(3).tolist()
            preview_str = "  •  ".join(preview) + ("  …" if proj_count > 3 else "")
            is_active   = st.session_state.show_notif_detail == notif["key"]

            col.markdown(f"""
            <div class="notif-alert" style="background:{notif['bg']};border:1.5px solid {notif['border']};
              border-radius:12px;padding:12px 14px;min-height:90px">
              <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px">
                <span style="font-size:16px">{notif['icon']}</span>
                <span style="font-size:10px;font-weight:800;color:{notif['color']};
                  background:{notif['color']}18;padding:2px 8px;border-radius:20px">{proj_count} projects</span>
              </div>
              <div style="font-size:12px;font-weight:800;color:#1E293B;margin-bottom:2px">{notif['label']}</div>
              <div style="font-size:10px;color:#64748B;margin-bottom:6px">{notif['note']}</div>
              <div style="font-size:9.5px;color:{notif['color']};font-style:italic;
                white-space:nowrap;overflow:hidden;text-overflow:ellipsis">{preview_str}</div>
            </div>""".strip(), unsafe_allow_html=True)

            btn_c1, btn_c2 = col.columns(2)
            detail_label = "Hide" if is_active else "Details"
            if btn_c1.button(detail_label, key=f"notif_detail_{notif['key']}", use_container_width=True,
                             type="primary" if is_active else "secondary"):
                st.session_state.show_notif_detail = None if is_active else notif["key"]
                st.rerun()
            if btn_c2.button("Dismiss", key=f"notif_dismiss_{notif['key']}", use_container_width=True,
                             help="Dismiss this notification"):
                st.session_state.dismissed_notifs.add(notif["key"])
                if st.session_state.show_notif_detail == notif["key"]:
                    st.session_state.show_notif_detail = None
                st.rerun()

        # ── NOTIFICATION POPUP DETAIL ─────────────────────────────────────────
        if st.session_state.show_notif_detail:
            nd_key   = st.session_state.show_notif_detail
            nd_info  = next((n for n in NOTIF_DEFS if n["key"] == nd_key), None)
            nd_projs = notif_data.get(nd_key, pd.DataFrame())
            if nd_info and not nd_projs.empty:
                st.markdown(f"""
                <div class="notif-popup" style="background:{nd_info['bg']};
                  border:2px solid {nd_info['border']}">
                  <div style="display:flex;align-items:center;justify-content:space-between;
                    margin-bottom:14px">
                    <div style="display:flex;align-items:center;gap:10px">
                      <span style="font-size:13px;font-weight:900;padding:2px 8px;background:{nd_info['border']};color:{nd_info['color']};border-radius:6px">{nd_info['icon']}</span>
                      <div>
                        <div style="font-size:14px;font-weight:800;color:#1E293B">
                          {nd_info['label']} — {len(nd_projs)} Projects</div>
                        <div style="font-size:11px;color:#64748B">{nd_info['note']}</div>
                      </div>
                    </div>
                  </div>
                </div>""".strip(), unsafe_allow_html=True)

                # Project table inside popup
                pop_hdr = st.columns([0.4, 3.0, 2.0, 2.2, 1.4, 1.2, 1.2])
                for ph, pl in zip(pop_hdr, ["ID","Project Name","Client","Employee","Status","Start","End"]):
                    ph.markdown(f'<div style="font-size:9px;font-weight:700;text-transform:uppercase;'
                                f'color:{nd_info["color"]};letter-spacing:.5px;padding:3px 0;'
                                f'border-bottom:2px solid {nd_info["border"]}">{pl}</div>',
                                unsafe_allow_html=True)

                for _, prow in nd_projs.iterrows():
                    pc = st.columns([0.4, 3.0, 2.0, 2.2, 1.4, 1.2, 1.2])
                    pc[0].markdown(cell(prow.get("id",""), size="10px", color="#94A3B8"), unsafe_allow_html=True)
                    pc[1].markdown(f'<span style="font-size:11px;font-weight:700;color:#111827">'
                                   f'{esc(str(prow.get("name","")))}</span>', unsafe_allow_html=True)
                    pc[2].markdown(cell(prow.get("client",""), size="11px"), unsafe_allow_html=True)
                    pc[3].markdown(cell(prow.get("employee",""), size="11px"), unsafe_allow_html=True)
                    pc[4].markdown(badge_html(str(prow.get("status",""))), unsafe_allow_html=True)
                    pc[5].markdown(cell(prow.get("start",""), size="10px", color="#64748B"), unsafe_allow_html=True)
                    pc[6].markdown(cell(prow.get("end","") or "Ongoing", size="10px", color="#64748B"), unsafe_allow_html=True)

                st.markdown("<br>", unsafe_allow_html=True)
                pa, pb, pc_col = st.columns([1.5, 1.5, 3])
                if pa.button("Open in Projects Tab", key="notif_goto_projects",
                             type="primary", use_container_width=True):
                    st.session_state.project_filter_preset = nd_key
                    st.session_state.active_tab            = "projects"
                    st.session_state.show_notif_detail     = None
                    st.rerun()
                if pb.button("Set Dashboard Filter", key="notif_set_slicer",
                             use_container_width=True):
                    st.session_state.dash_slicer       = nd_key
                    st.session_state.show_notif_detail = None
                    st.rerun()
                if pc_col.button("Close Panel", key="notif_close_popup",
                                 use_container_width=True):
                    st.session_state.show_notif_detail = None
                    st.rerun()
                st.markdown("---")

    if st.session_state.dismissed_notifs:
        if st.button("Restore Notifications", key="restore_notifs",
                     help="Re-show all dismissed alerts"):
            st.session_state.dismissed_notifs = set()
            st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)

    # ── ROI BANNER ────────────────────────────────────────────────────────────
    if dash_stats["total_hrs"] > 0:
        st.markdown(f"""
        <div class="roi-banner">
          <span style="font-size:13px;font-weight:800;color:#6EE7B7;letter-spacing:1px">ROI</span>
          <div>
            <div style="font-size:10px;color:#6EE7B7;font-weight:700;letter-spacing:1px;text-transform:uppercase">
              Cumulative ROI This Session</div>
            <div style="display:flex;gap:28px;margin-top:6px">
              <span><b style="font-size:22px;color:#10B981;font-family:'JetBrains Mono',monospace">{dash_stats['total_hrs']:.0f}</b>
                <span style="color:#6EE7B7;font-size:12px;margin-left:4px">hrs saved</span></span>
              <span><b style="font-size:22px;color:#10B981;font-family:'JetBrains Mono',monospace">&#8377;{dash_stats['total_cost']:,.0f}</b>
                <span style="color:#6EE7B7;font-size:12px;margin-left:4px">cost saved</span></span>
            </div>
          </div>
        </div>""".strip(), unsafe_allow_html=True)

    # ── CHARTS ────────────────────────────────────────────────────────────────
    _client_label = st.session_state.dash_client_filter
    _slicer_key   = st.session_state.dash_slicer

    # Apply slicer filter on top of client filter so the pie reflects both
    _pie_df = dash_df
    _slicer_label = None
    if _slicer_key is not None:
        if _slicer_key == "__active_dev__":
            _sm = dash_df["status"].isin({"In Progress"})
            if "is_active" in dash_df.columns:
                _sm = _sm & (~dash_df["is_active"].astype(str).str.strip().str.lower().isin(["false","0","no"]))
            _pie_df, _slicer_label = dash_df[_sm], "Active Development"
        elif _slicer_key == "__new__":
            _sm = dash_df["is_new"].astype(str).str.lower().isin(["true","1","yes"]) if "is_new" in dash_df.columns else pd.Series([False]*len(dash_df))
            _pie_df, _slicer_label = dash_df[_sm], "New Added"
        elif _slicer_key == "POC":
            _pie_df, _slicer_label = dash_df[dash_df["status"].str.contains("POC", na=False)], "POC"
        else:
            _pie_df, _slicer_label = dash_df[dash_df["status"].str.contains(_slicer_key, na=False)], _slicer_key

    _cl_part = f" — {_client_label}" if _client_label != "All" else " — All Clients"
    _sl_part = f" · {_slicer_label}" if _slicer_label else ""

    _chart_c1, _chart_c2 = st.columns(2)

    with _chart_c1:
        with st.container(border=True):
            _pie_title = f"Status Breakdown{_cl_part}{_sl_part}"
            st.markdown(f'<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin-bottom:10px">{_pie_title}</div>', unsafe_allow_html=True)
            if not _pie_df.empty:
                _sc = _pie_df["status"].value_counts().reset_index()
                _sc.columns = ["status", "count"]
                _color_map = {s: STATUS_STYLES.get(s, {"dot": "#94A3B8"})["dot"] for s in _sc["status"]}
                fig = px.pie(_sc, names="status", values="count", color="status",
                             color_discrete_map=_color_map, hole=0.45)
                fig.update_layout(margin=dict(t=0,b=0,l=0,r=0), height=240,
                                  legend=dict(font=dict(size=9), orientation="v"),
                                  paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("No data for current filter.")

    with _chart_c2:
        with st.container(border=True):
            _bar_title = f"Projects by Client{_cl_part}{_sl_part}"
            st.markdown(f'<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin-bottom:10px">{_bar_title}</div>', unsafe_allow_html=True)
            _bar_src = _pie_df if not _pie_df.empty else dash_df
            if not _bar_src.empty and "client" in _bar_src.columns:
                _ccounts = (_bar_src.groupby("client").size()
                            .reset_index(name="count")
                            .sort_values("count", ascending=True))
                _bar_fig = go.Figure(go.Bar(
                    x=_ccounts["count"], y=_ccounts["client"],
                    orientation="h",
                    marker=dict(color="#3B82F6", opacity=0.85),
                    text=_ccounts["count"], textposition="outside",
                    textfont=dict(size=10),
                ))
                _bar_fig.update_layout(
                    margin=dict(t=0, b=0, l=0, r=30), height=240,
                    xaxis=dict(visible=False),
                    yaxis=dict(tickfont=dict(size=10)),
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                )
                st.plotly_chart(_bar_fig, use_container_width=True)
            else:
                st.info("No client data available.")

    # ── PROJECT DETAIL PANEL (always visible; slicer narrows the view) ───────────
    _detail_key = st.session_state.dash_slicer

    if _detail_key is not None:
        key = _detail_key
        if key == "__active_dev__":
            _ad_mask = dash_df["status"].isin({"In Progress"})
            if "is_active" in dash_df.columns:
                _ad_mask = _ad_mask & (
                    ~dash_df["is_active"].astype(str).str.strip().str.lower().isin(["false","0","no"])
                )
            sliced, slicer_label = dash_df[_ad_mask], "Active Development"
        elif key == "__new__":
            new_mask = dash_df["is_new"].astype(str).str.lower().isin(["true","1","yes"]) if "is_new" in dash_df.columns else pd.Series([False]*len(dash_df))
            sliced, slicer_label = dash_df[new_mask], "New Added"
        elif key == "POC":
            sliced, slicer_label = dash_df[dash_df["status"].str.contains("POC", na=False)], "POC (Internal + External)"
        else:
            sliced, slicer_label = dash_df[dash_df["status"].str.contains(key, na=False)], key

        # For Development-related slicers, show only active projects
        _dev_keys = {"In Progress", "PDD", "Important"}
        if key in _dev_keys and "is_active" in sliced.columns:
            sliced = sliced[~sliced["is_active"].astype(str).str.strip().str.lower().isin(["false","0","no"])]
    else:
        # No slicer active — show all projects for the current client filter
        sliced = dash_df
        _cl = st.session_state.dash_client_filter
        slicer_label = f"All — {_cl}" if _cl != "All" else "All Projects"

    # Build team map
    emp_map = {}
    for _, row in sliced.iterrows():
        for n in str(row.get("employee","")).replace("&",",").split(","):
            n = n.strip()
            if not n: continue
            if n not in emp_map: emp_map[n] = {"projects":[], "clients":set()}
            emp_map[n]["projects"].append(row["name"])
            emp_map[n]["clients"].add(str(row.get("client","")))
    team_list = sorted(emp_map.items(), key=lambda x: -len(x[1]["projects"]))

    st.markdown("<br>", unsafe_allow_html=True)
    hc1, hc2 = st.columns([5, 1])
    _style_key = _detail_key if _detail_key not in [None, "__new__", "POC", "__active_dev__"] else (
        "In Progress" if _detail_key == "__active_dev__" else
        "Completed"   if _detail_key == "__new__"        else
        "Internal POC" if _detail_key == "POC"           else "R&M"
    )
    hc1.markdown(f"""
    <div style="display:flex;align-items:center;gap:12px;padding:10px 16px;
      background:#fff;border:1px solid #E2E8F0;border-radius:10px">
      {badge_html(slicer_label if slicer_label in STATUS_STYLES else "R&M")}
      <span style="color:#64748B;font-size:12px;font-weight:500">
        <b style="color:#0F172A">{len(sliced)}</b> projects &nbsp;·&nbsp;
        <b style="color:#0F172A">{len(team_list)}</b> team members assigned
        &nbsp;·&nbsp; <b style="color:#64748B">{slicer_label}</b></span>
    </div>""".strip(), unsafe_allow_html=True)
    if _detail_key is not None:
        if hc2.button("Clear Slicer", use_container_width=True, key="clear_slicer"):
            st.session_state.dash_slicer = None
            st.rerun()

    pl, pr = st.columns([1.6, 1])

    # ── Project detail cards ──────────────────────────────────────────────
    with pl:
        with st.container(border=True):
            st.markdown(f'<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;padding-bottom:8px;border-bottom:1px solid #E2E8F0">Project Details — {len(sliced)} records</div>', unsafe_allow_html=True)
            if sliced.empty:
                st.info("No projects in this category.")
            else:
                for i, (_, row) in enumerate(sliced.iterrows()):
                    roi_badge  = ""
                    if str(row.get("roi_pct","")).strip():
                        roi_badge = f'<span style="font-size:10px;background:#064E3B;color:#10B981;border-radius:4px;padding:2px 8px;font-weight:800;margin-left:6px">ROI {esc(str(row["roi_pct"]))}%</span>'
                    new_badge  = '<span style="font-size:9px;background:#10B981;color:#fff;border-radius:4px;padding:1px 5px;font-weight:800;margin-left:4px">NEW</span>' if is_new(row) else ""
                    _lead      = esc(str(row.get("lead","")).strip())
                    _start     = esc(str(row.get("start","")))
                    _end       = esc(str(row.get("end","")) or "Ongoing")
                    _due_raw   = str(row.get("due_date","")).strip()
                    _po        = esc(str(row.get("po","")))
                    _desc      = esc(str(row.get("desc","")))

                    meta_spans = [f'<span>{esc(str(row.get("client","")))} </span>']
                    if _lead:
                        meta_spans.append(f'<span>Lead: <b style="color:#2563EB">{_lead}</b></span>')
                    meta_spans.append(f'<span>{esc(str(row.get("employee","")))} </span>')
                    if _start:
                        meta_spans.append(f'<span>{_start} to {_end}</span>')
                    if _due_raw:
                        _due_d = _parse_dmy(_due_raw)
                        _due_color = "#DC2626" if (_due_d and (_due_d - date.today()).days < 0) else "#92400E" if (_due_d and (_due_d - date.today()).days <= 7) else "#64748B"
                        meta_spans.append(f'<span>Due: <b style="color:{_due_color}">{esc(_due_raw)}</b></span>')
                    if _po:
                        meta_spans.append(f'<span>PO #{_po}</span>')
                    meta_html = "".join(meta_spans)
                    desc_html = f'<div style="font-size:10px;color:#64748B;font-style:italic">{_desc}</div>' if _desc else ""
                    row_bg    = "#fff" if i % 2 == 0 else "#F8FAFC"

                    st.markdown(
                        f'<div class="srow" style="background:{row_bg}">'
                        f'<div style="flex:1">'
                        f'<div style="font-size:12px;font-weight:700;color:#111827;margin-bottom:4px">{esc(str(row.get("name","")))}{new_badge}</div>'
                        f'<div style="display:flex;flex-wrap:wrap;gap:10px;font-size:10px;color:#64748B;margin-bottom:3px">{meta_html}</div>'
                        f'{desc_html}{roi_badge}'
                        f'</div>'
                        f'<div style="flex-shrink:0;margin-left:10px">{badge_html(str(row.get("status","")))}</div>'
                        f'</div>',
                        unsafe_allow_html=True,
                    )

    # ── Team panel ────────────────────────────────────────────────────────
    AVATAR_COLS = [("#1E3A8A","#3B82F6"),("#451A03","#F59E0B"),("#064E3B","#10B981"),
                   ("#1E1B4B","#8B5CF6"),("#7F1D1D","#EF4444"),("#0C4A6E","#06B6D4"),
                   ("#78350F","#F97316"),("#500724","#EC4899")]
    with pr:
        with st.container(border=True):
            st.markdown('<div style="font-size:9px;color:#94A3B8;font-weight:600;text-transform:uppercase;letter-spacing:.8px;padding-bottom:8px;border-bottom:1px solid #E2E8F0">Team Responsible</div>', unsafe_allow_html=True)
            if not team_list:
                st.info("No team members.")
            else:
                for i, (name, info) in enumerate(team_list):
                    bg_c, ac = AVATAR_COLS[i % len(AVATAR_COLS)]
                    clients_str = " · ".join(esc(c) for c in sorted(info["clients"]))
                    st.markdown(f"""
                    <div style="display:flex;align-items:center;gap:10px;padding:10px 4px;border-bottom:1px solid #F1F5F9">
                      <div style="width:36px;height:36px;border-radius:10px;flex-shrink:0;
                        background:linear-gradient(135deg,{bg_c},{ac}44);border:1px solid {ac}55;
                        display:flex;align-items:center;justify-content:center;
                        font-size:14px;font-weight:800;color:{ac}">{esc(name[0].upper())}</div>
                      <div style="flex:1;min-width:0">
                        <div style="font-size:13px;font-weight:700;color:#111827">{esc(name)}</div>
                        <div style="font-size:10px;color:#64748B;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">
                          {clients_str}</div>
                      </div>
                      <div style="width:26px;height:26px;border-radius:7px;flex-shrink:0;
                        background:{ac}22;border:1px solid {ac}44;display:flex;align-items:center;
                        justify-content:center;font-size:13px;font-weight:800;color:{ac};
                        font-family:'JetBrains Mono',monospace">{len(info["projects"])}</div>
                    </div>""".strip(), unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TAB: PROJECTS
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "projects" and role != "employee":

    # ── Shared filter bar (client, project search, active/inactive) ───────────
    f1, f2, f3 = st.columns([2, 1.5, 1.5])
    search_q      = f1.text_input("Search", placeholder="Search projects…",
                                  label_visibility="collapsed")
    client_filter = f2.selectbox("Client", ["All"] + sorted(df["client"].dropna().unique().tolist()),
                                 label_visibility="collapsed")
    active_filter = f3.selectbox("Status", ["All", "Active", "Inactive"],
                                 label_visibility="collapsed")

    # ── Helper: apply shared filters to a sub-set of projects ─────────────────
    def _apply_filters(subset):
        out = subset.copy()
        if search_q:
            q = search_q.lower()
            sc = [c for c in ["name","employee","lead","client","desc"] if c in out.columns]
            mask = (out[sc].fillna("").astype(str)
                    .apply(lambda col: col.str.lower().str.contains(q, regex=False))
                    .any(axis=1))
            out = out[mask]
        if client_filter != "All":
            out = out[out["client"] == client_filter]
        if active_filter != "All" and "is_active" in out.columns:
            inactive_vals = {"false", "0", "no"}
            raw_col = out["is_active"].astype(str).str.strip().str.lower()
            if active_filter == "Active":
                out = out[~raw_col.isin(inactive_vals)]
            else:
                out = out[raw_col.isin(inactive_vals)]
        return out

    # ── Helper: render active/inactive pill ───────────────────────────────────
    def _active_pill(row):
        raw = str(row.get("is_active","True")).strip().lower()
        if raw in ["false","0","no"]:
            return '<span style="display:inline-flex;align-items:center;gap:3px;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;background:#FEF2F2;color:#991B1B"><span style="width:5px;height:5px;border-radius:50%;background:#EF4444;display:inline-block"></span>Inactive</span>'
        return '<span style="display:inline-flex;align-items:center;gap:3px;padding:2px 8px;border-radius:20px;font-size:10px;font-weight:700;background:#ECFDF5;color:#065F46"><span style="width:5px;height:5px;border-radius:50%;background:#10B981;display:inline-block"></span>Active</span>'

    # ── Helper: render a project table for a given filtered DataFrame ──────────
    _ROW_BG = {
        "Important": "#FFF1F2",
        "Completed": "#ECFDF5",
        "R&M":       "#EFF6FF",
    }

    def _render_project_table(filtered, tab_key=""):
        st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 10px"><b>{len(filtered)}</b> projects</p>',
                    unsafe_allow_html=True)
        if filtered.empty:
            st.info("No projects match the current filters.")
            return

        is_admin   = (role == "admin")
        can_edit   = role in ("admin", "lead", "manager")
        col_widths = [10, 0.4, 0.4] if is_admin else ([10, 0.4] if can_edit else [10])

        # Helper: type badge
        def _type_badge(pt):
            if pt == "RPA":
                return '<span style="font-size:9px;font-weight:700;background:#DBEAFE;color:#1D4ED8;padding:1px 6px;border-radius:4px;margin-left:4px">RPA</span>'
            if pt == "AI Agent":
                return '<span style="font-size:9px;font-weight:700;background:#F3E8FF;color:#7C3AED;padding:1px 6px;border-radius:4px;margin-left:4px">AI</span>'
            return ""

        # Header row
        hcols = st.columns(col_widths)
        hcols[0].markdown(
            f'<div style="display:flex;gap:0;align-items:center">'
            f'<div style="width:3%;{_HDR_STYLE}">ID</div>'
            f'<div style="width:19%;{_HDR_STYLE}">Project Name</div>'
            f'<div style="width:11%;{_HDR_STYLE}">Client</div>'
            f'<div style="width:9%;{_HDR_STYLE}">Lead</div>'
            f'<div style="width:14%;{_HDR_STYLE}">Employee</div>'
            f'<div style="width:6%;{_HDR_STYLE}">Type</div>'
            f'<div style="width:7%;{_HDR_STYLE}">Start</div>'
            f'<div style="width:7%;{_HDR_STYLE}">End</div>'
            f'<div style="width:8%;{_HDR_STYLE}">Due Date</div>'
            f'<div style="width:6%;{_HDR_STYLE}">PO</div>'
            f'<div style="width:5%;{_HDR_STYLE}">Active</div>'
            f'</div>',
            unsafe_allow_html=True
        )
        if can_edit:
            hcols[1].markdown(f'<div style="{_HDR_STYLE}"></div>', unsafe_allow_html=True)
        if is_admin:
            hcols[2].markdown(f'<div style="{_HDR_STYLE}"></div>', unsafe_allow_html=True)

        # Data rows — one st.columns per row, all text in one markdown
        for _, row in filtered.iterrows():
            row_status = str(row.get("status",""))
            bg = next((_ROW_BG[s] for s in _ROW_BG if s in row_status), "#FFFFFF")
            new_tag = ' <span style="font-size:9px;font-weight:700;background:#DBEAFE;color:#1D4ED8;padding:1px 5px;border-radius:4px">NEW</span>' if is_new(row) else ""
            type_badge = _type_badge(str(row.get("proj_type","")).strip())
            lead_val = str(row.get("lead","")).strip()
            lead_html = (f'<span style="font-size:11px;font-weight:600;color:#2563EB">{esc(lead_val)}</span>'
                         if lead_val else '<span style="font-size:11px;color:#CBD5E1">—</span>')
            raw_active = str(row.get("is_active","True")).strip().lower()
            active_html = (
                '<span style="font-size:10px;font-weight:700;background:#FEF2F2;color:#991B1B;padding:2px 6px;border-radius:20px">Inactive</span>'
                if raw_active in ["false","0","no"] else
                '<span style="font-size:10px;font-weight:700;background:#ECFDF5;color:#065F46;padding:2px 6px;border-radius:20px">Active</span>'
            )
            rcols = st.columns(col_widths)
            rcols[0].markdown(
                f'<div style="display:flex;gap:0;align-items:center;background:{bg};padding:7px 0;border-bottom:1px solid #F1F5F9">'
                f'<div style="width:3%;font-size:10px;color:#94A3B8">{esc(str(row.get("id","")))}</div>'
                f'<div style="width:19%;font-size:12px;font-weight:600;color:#111827">{esc(str(row.get("name","")))}{new_tag}{type_badge}</div>'
                f'<div style="width:11%;font-size:12px;color:#374151">{esc(str(row.get("client","")))}</div>'
                f'<div style="width:9%">{lead_html}</div>'
                f'<div style="width:14%;font-size:11px;color:#374151">{esc(str(row.get("employee","")))}</div>'
                f'<div style="width:6%">{_type_badge(str(row.get("proj_type","")).strip()) or cell("—","10px","#CBD5E1")}</div>'
                f'<div style="width:7%;font-size:11px;color:#64748B">{esc(str(row.get("start","")))}</div>'
                f'<div style="width:7%;font-size:11px;color:#64748B">{esc(str(row.get("end","")))}</div>'
                f'<div style="width:8%">{_due_cell(str(row.get("due_date","")))}</div>'
                f'<div style="width:6%;font-size:11px;color:#94A3B8">{esc(str(row.get("po","")))}</div>'
                f'<div style="width:5%">{active_html}</div>'
                f'</div>',
                unsafe_allow_html=True
            )
            rid = str(row.get("id",""))
            if can_edit:
                with rcols[1]:
                    st.markdown('<span class="act-edit-marker"></span>', unsafe_allow_html=True)
                    if st.button("✏", key=f"edit_{tab_key}_{rid}", help="Edit project", use_container_width=True):
                        st.session_state.show_modal = {"edit": row.to_dict()}
                        st.rerun()
            if is_admin:
                with rcols[2]:
                    st.markdown('<span class="act-del-marker"></span>', unsafe_allow_html=True)
                    if st.button("🗑", key=f"del_{tab_key}_{rid}", help="Delete project", use_container_width=True):
                        st.session_state.confirm_delete = {"id": rid, "name": str(row.get("name",""))}
                        st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        csv = filtered.to_csv(index=False)
        st.download_button("Export CSV", csv,
                           file_name=f"qualesce_{tab_key}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                           mime="text/csv", key=f"csv_{tab_key}")

    # ── Sub-tab status mapping ─────────────────────────────────────────────────
    _DEV_STATUSES        = {"In Progress", "PDD", "Important"}
    _RM_STATUSES         = {"R&M"}
    _COMPLETED_STATUSES  = {"Completed", "Discontinued"}
    _UAT_STATUSES        = {"UAT"}

    tab_dev, tab_rm, tab_completed, tab_uat = st.tabs([
        "Development",
        "R&M",
        "Completed",
        "UAT",
    ])

    with tab_dev:
        dev_df   = df[df["status"].isin(_DEV_STATUSES)]
        _render_project_table(_apply_filters(dev_df), tab_key="dev")

    with tab_rm:
        rm_df    = df[df["status"].isin(_RM_STATUSES)]
        _render_project_table(_apply_filters(rm_df), tab_key="rm")

    with tab_completed:
        comp_df  = df[df["status"].isin(_COMPLETED_STATUSES)]
        _render_project_table(_apply_filters(comp_df), tab_key="completed")

    with tab_uat:
        uat_df   = df[df["status"].isin(_UAT_STATUSES)]
        _render_project_table(_apply_filters(uat_df), tab_key="uat")

# ══════════════════════════════════════════════════════════════════════════════
# TAB: PRESALES
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "presales" and role not in ("employee",):
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#0F172A;margin-bottom:4px">Presales / POC</h2>', unsafe_allow_html=True)
    st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:16px">Presales pipeline and proof-of-concept projects</p>', unsafe_allow_html=True)

    _POC_DEFAULT = {"Presales", "Internal POC", "External POC"}
    _POC_CLIENTS = {"Internal POC", "External POC"}

    PS_ROW_BG = {
        "Important":    "#FFF1F2",
        "Presales":     "#F0F9FF",
        "Internal POC": "#F5F3FF",
        "External POC": "#FDF2F8",
        "Completed":    "#ECFDF5",
        "In Progress":  "#ECFEFF",
        "Discontinued": "#FEF2F2",
    }

    def _ps_type_badge(pt):
        if pt == "RPA":
            return '<span style="font-size:9px;font-weight:700;background:#DBEAFE;color:#1D4ED8;padding:1px 6px;border-radius:4px">RPA</span>'
        if pt == "AI Agent":
            return '<span style="font-size:9px;font-weight:700;background:#F3E8FF;color:#7C3AED;padding:1px 6px;border-radius:4px">AI</span>'
        if pt == "Presales":
            return '<span style="font-size:9px;font-weight:700;background:#FEF9C3;color:#854D0E;padding:1px 6px;border-radius:4px">Pre</span>'
        return '<span style="font-size:10px;color:#CBD5E1">—</span>'

    def _render_poc_table(data, tab_key):
        _is_adm = (role == "admin")
        _can_ed = role in ("admin", "lead", "manager")
        _cw = [10, 0.4, 0.4] if _is_adm else ([10, 0.4] if _can_ed else [10])
        if data.empty:
            st.info("No projects found.")
            return
        st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 12px"><b>{len(data)}</b> project(s)</p>',
                    unsafe_allow_html=True)
        _hc = st.columns(_cw)
        _hc[0].markdown(
            f'<div style="display:flex;gap:0;align-items:center">'
            f'<div style="width:3%;{_HDR_STYLE}">ID</div>'
            f'<div style="width:17%;{_HDR_STYLE}">Project Name</div>'
            f'<div style="width:10%;{_HDR_STYLE}">Client</div>'
            f'<div style="width:9%;{_HDR_STYLE}">Lead</div>'
            f'<div style="width:12%;{_HDR_STYLE}">Employee</div>'
            f'<div style="width:6%;{_HDR_STYLE}">Type</div>'
            f'<div style="width:10%;{_HDR_STYLE}">Status</div>'
            f'<div style="width:7%;{_HDR_STYLE}">Start</div>'
            f'<div style="width:7%;{_HDR_STYLE}">End</div>'
            f'<div style="width:8%;{_HDR_STYLE}">Due Date</div>'
            f'<div style="width:11%;{_HDR_STYLE}">Notes</div>'
            f'</div>',
            unsafe_allow_html=True
        )
        if _can_ed: _hc[1].markdown(f'<div style="{_HDR_STYLE}"></div>', unsafe_allow_html=True)
        if _is_adm: _hc[2].markdown(f'<div style="{_HDR_STYLE}"></div>', unsafe_allow_html=True)
        for _, _row in data.iterrows():
            _rstat = str(_row.get("status",""))
            _bg = next((PS_ROW_BG[s] for s in PS_ROW_BG if s in _rstat), "#FFFFFF")
            _new_tag = (' <span style="font-size:9px;font-weight:700;background:#DBEAFE;color:#1D4ED8;'
                        'padding:1px 5px;border-radius:4px">NEW</span>') if is_new(_row) else ""
            _lv = str(_row.get("lead","")).strip()
            _lead_html = (f'<span style="font-size:11px;font-weight:600;color:#2563EB">{esc(_lv)}</span>'
                          if _lv else '<span style="font-size:11px;color:#CBD5E1">—</span>')
            _rid = str(_row.get("id",""))
            _inline_active = (st.session_state.poc_row_edit == _rid)
            _notes_val = str(_row.get("desc","")).strip()
            _notes_disp = (f'<span style="font-size:11px;color:#374151">{esc(_notes_val)}</span>'
                           if _notes_val else '<span style="font-size:11px;color:#CBD5E1">—</span>')
            _rc = st.columns(_cw)
            _rc[0].markdown(
                f'<div style="display:flex;gap:0;align-items:center;background:{_bg};padding:7px 0;border-bottom:1px solid #F1F5F9">'
                f'<div style="width:3%;font-size:10px;color:#94A3B8">{esc(str(_row.get("id","")))}</div>'
                f'<div style="width:17%;font-size:12px;font-weight:600;color:#111827">{esc(str(_row.get("name","")))}{_new_tag}</div>'
                f'<div style="width:10%;font-size:12px;color:#374151">{esc(str(_row.get("client","")))}</div>'
                f'<div style="width:9%">{_lead_html}</div>'
                f'<div style="width:12%;font-size:11px;color:#374151">{esc(str(_row.get("employee","")))}</div>'
                f'<div style="width:6%">{_ps_type_badge(str(_row.get("proj_type","")).strip())}</div>'
                f'<div style="width:10%">{badge_html(str(_row.get("status","")))}</div>'
                f'<div style="width:7%;font-size:11px;color:#64748B">{esc(str(_row.get("start","")))}</div>'
                f'<div style="width:7%;font-size:11px;color:#64748B">{esc(str(_row.get("end","")))}</div>'
                f'<div style="width:8%">{_due_cell(str(_row.get("due_date","")))}</div>'
                f'<div style="width:11%">{_notes_disp}</div>'
                f'</div>',
                unsafe_allow_html=True
            )
            if _can_ed:
                with _rc[1]:
                    if not _inline_active:
                        st.markdown('<span class="act-edit-marker"></span>', unsafe_allow_html=True)
                    if st.button("✕" if _inline_active else "✏", key=f"{tab_key}_edit_{_rid}",
                                 help="Cancel" if _inline_active else "Edit notes",
                                 use_container_width=True):
                        if _inline_active:
                            st.session_state.poc_row_edit = None
                        else:
                            st.session_state.poc_row_edit = _rid
                        st.rerun()
            if _is_adm:
                with _rc[2]:
                    st.markdown('<span class="act-del-marker"></span>', unsafe_allow_html=True)
                    if st.button("🗑", key=f"{tab_key}_del_{_rid}", help="Delete", use_container_width=True):
                        st.session_state.confirm_delete = {"id": _rid, "name": str(_row.get("name",""))}
                        st.rerun()
            if _inline_active and _can_ed:
                with st.container():
                    _ic1, _ic2 = st.columns([3, 1])
                    _new_comment = _ic1.text_area(
                        "Notes / Comment", value=_notes_val,
                        key=f"{tab_key}_inline_comment_{_rid}", height=72,
                        label_visibility="collapsed", placeholder="Add notes or comment…"
                    )
                    _b1, _b2, _b3 = _ic2.columns(3)
                    if _b1.button("💾", key=f"{tab_key}_save_cmt_{_rid}", help="Save comment"):
                        _proj_idx = st.session_state.projects.index[
                            st.session_state.projects["id"].astype(str) == _rid
                        ]
                        if len(_proj_idx) > 0:
                            st.session_state.projects.at[_proj_idx[0], "desc"] = _new_comment.strip()
                            save_to_excel(st.session_state.projects)
                            st.session_state.excel_mtime = excel_mtime()
                        st.session_state.poc_row_edit = None
                        st.session_state.toast = {"msg": "Comment saved!", "type": "success"}
                        st.rerun()
                    if _b2.button("✏️", key=f"{tab_key}_full_edit_{_rid}", help="Full edit"):
                        st.session_state.poc_row_edit = None
                        st.session_state.show_modal = {"edit": _row.to_dict()}
                        st.rerun()
                    if _b3.button("✕", key=f"{tab_key}_cancel_cmt_{_rid}", help="Cancel"):
                        st.session_state.poc_row_edit = None
                        st.rerun()
        st.markdown("<br>", unsafe_allow_html=True)
        st.download_button("Export CSV", data.to_csv(index=False),
                           file_name=f"qualesce_{tab_key}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                           mime="text/csv", key=f"csv_{tab_key}")

    # ── Create New form (admin / lead / manager only) ─────────────────────────
    if role in ("admin", "lead", "manager"):
        with st.expander("+ Create New Presales / POC Entry", expanded=False):
            with st.container():
                _pc1, _pc2 = st.columns(2)
                _ps_new_name   = _pc1.text_input("Project Name *", key="ps_new_name")
                _ps_new_client = _pc2.text_input("Client Name *",  key="ps_new_client")

                _pc3, _pc4 = st.columns(2)
                _ps_all_emp = sorted(set(
                    n.strip()
                    for raw in df.get("employee", pd.Series(dtype=str)).dropna()
                    for n in str(raw).replace("&", ",").split(",")
                    if n.strip()
                ))
                _ps_all_leads = sorted(set(
                    str(l).strip() for l in df.get("lead", pd.Series(dtype=str)).dropna() if str(l).strip()
                )) if "lead" in df.columns else []

                _ps_new_lead = _pc3.selectbox(
                    "Lead", [""] + _ps_all_leads + ["── Type new ──"],
                    key="ps_new_lead"
                )
                if _ps_new_lead == "── Type new ──":
                    _ps_new_lead = _pc3.text_input("Enter lead name", key="ps_new_lead_txt")

                _ps_new_emp_sel = _pc4.multiselect(
                    "Employee(s)", options=_ps_all_emp, key="ps_new_emp_sel"
                )
                _ps_new_emp_txt = _pc4.text_input(
                    "Add new employee (optional)", key="ps_new_emp_txt",
                    placeholder="leave blank if not needed"
                )
                _ps_new_emp = ", ".join(_ps_new_emp_sel + ([_ps_new_emp_txt.strip()] if _ps_new_emp_txt.strip() else []))

                _pc5, _pc6 = st.columns(2)
                _PS_NEW_TYPES    = ["", "RPA", "AI Agent", "Presales"]
                _ps_new_type     = _pc5.selectbox("Type", _PS_NEW_TYPES, key="ps_new_type",
                                                   format_func=lambda x: "— Select type —" if x == "" else x)
                _PS_NEW_STATUSES = ["Internal POC", "External POC",
                                    "In Progress", "Completed", "Discontinued"]
                _ps_new_status   = _pc6.selectbox("Status", _PS_NEW_STATUSES, key="ps_new_status")

                _pc7, _pc8, _pc9 = st.columns([1.5, 1.5, 2])
                _ps_new_start_dt = _pc7.date_input("Start Date (optional)", value=None,
                                                    key="ps_new_start", format="DD/MM/YYYY")
                _ps_new_end_dt   = _pc8.date_input("End Date (optional)", value=None,
                                                    key="ps_new_end", format="DD/MM/YYYY")
                _ps_new_comment  = _pc9.text_area("Notes / Comment", key="ps_new_comment", height=68)
                _ps_new_start = _ps_new_start_dt.strftime("%d/%m/%Y") if _ps_new_start_dt else ""
                _ps_new_end   = _ps_new_end_dt.strftime("%d/%m/%Y")   if _ps_new_end_dt   else ""

                if st.button("Save New Entry", type="primary", key="ps_new_save"):
                    _ps_errs = []
                    if not _ps_new_name.strip():   _ps_errs.append("Project name is required.")
                    if not _ps_new_client.strip(): _ps_errs.append("Client name is required.")
                    if _ps_errs:
                        for _e in _ps_errs: st.error(_e)
                    else:
                        _ps_row = {
                            "id": st.session_state.next_id,
                            "name": _ps_new_name.strip(),
                            "client": _ps_new_client.strip(),
                            "lead": _ps_new_lead.strip() if _ps_new_lead != "── Type new ──" else "",
                            "employee": _ps_new_emp,
                            "status": _ps_new_status,
                            "proj_type": _ps_new_type,
                            "start": _ps_new_start,
                            "end": _ps_new_end,
                            "po": "", "desc": _ps_new_comment.strip(),
                            "manual_hrs": "", "auto_hrs": "", "cost_per_hr": "",
                            "hours_saved": "", "cost_saved": "", "roi_pct": "",
                            "is_new": True, "is_active": True,
                        }
                        st.session_state.projects = pd.concat(
                            [st.session_state.projects, pd.DataFrame([_ps_row])], ignore_index=True
                        )
                        st.session_state.next_id += 1
                        save_to_excel(st.session_state.projects)
                        st.session_state.excel_mtime = excel_mtime()
                        st.session_state.toast = {"msg": f'"{_ps_new_name.strip()}" added!', "type": "success"}
                        st.rerun()

    # ── Sub-tabs ──────────────────────────────────────────────────────────────
    _ps_t2, _ps_t3, _ps_t4 = st.tabs(["In Progress", "Completed", "Discontinued"])

    _POC_MASK = df["client"].isin(_POC_CLIENTS) | (df["proj_type"].fillna("").str.strip() == "Presales")
    _ACTIVE_STATUSES = {"In Progress", "Presales", "Internal POC", "External POC"}

    with _ps_t2:
        _ip_df = df[_POC_MASK & df["status"].isin(_ACTIVE_STATUSES)].copy()
        st.markdown('<p style="color:#64748B;font-size:12px;margin:0 0 12px">'
                    'POC / Presales projects currently in development</p>', unsafe_allow_html=True)
        _render_poc_table(_ip_df, "poc_ip")

    with _ps_t3:
        _done_df = df[_POC_MASK & (df["status"] == "Completed")].copy()
        st.markdown('<p style="color:#64748B;font-size:12px;margin:0 0 12px">'
                    'Successfully completed POC / Presales projects</p>', unsafe_allow_html=True)
        _render_poc_table(_done_df, "poc_done")

    with _ps_t4:
        _disc_df = df[_POC_MASK & (df["status"] == "Discontinued")].copy()
        st.markdown('<p style="color:#64748B;font-size:12px;margin:0 0 12px">'
                    'Discontinued POC / Presales projects</p>', unsafe_allow_html=True)
        _render_poc_table(_disc_df, "poc_disc")

# ══════════════════════════════════════════════════════════════════════════════
# TAB: LICENSE
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "license" and role != "employee":
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#0F172A;margin-bottom:4px">License Management</h2>', unsafe_allow_html=True)
    st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:16px">Track purchased and sold licenses</p>', unsafe_allow_html=True)

    def _lc_expiry_badge(end_date: str) -> str:
        if not end_date:
            return '<span style="font-size:10px;color:#94A3B8">—</span>'
        try:
            exp = datetime.strptime(end_date, "%Y-%m-%d").date()
            today = datetime.now().date()
            diff  = (exp - today).days
            if diff < 0:
                return (f'<span style="background:#FEF2F2;color:#991B1B;font-size:10px;font-weight:700;'
                        f'padding:2px 8px;border-radius:10px">Expired</span>')
            elif diff <= 30:
                return (f'<span style="background:#FFFBEB;color:#92400E;font-size:10px;font-weight:700;'
                        f'padding:2px 8px;border-radius:10px">Expiring in {diff}d</span>')
            else:
                return (f'<span style="background:#ECFDF5;color:#065F46;font-size:10px;font-weight:700;'
                        f'padding:2px 8px;border-radius:10px">Active</span>')
        except ValueError:
            return f'<span style="font-size:11px;color:#64748B">{esc(end_date)}</span>'

    _licenses_all     = auth.get_all_licenses()
    _sold_licenses_all = auth.get_all_sold_licenses()

    # Tool names from purchased licenses (for Sold License dropdown)
    _purchased_tool_names = sorted({l["tool_name"].strip() for l in _licenses_all if l["tool_name"].strip()})

    _lc_tab1, _lc_tab2 = st.tabs(["Purchased License", "Sold License"])

    # ══════════════════════════════════════════════════════════════════════════
    # SUB-TAB 1 — PURCHASED LICENSE
    # ══════════════════════════════════════════════════════════════════════════
    with _lc_tab1:
        # ── Edit form ────────────────────────────────────────────────────────
        if st.session_state.lc_edit_id is not None:
            _lc_rec = next((x for x in _licenses_all if x["id"] == st.session_state.lc_edit_id), None)
            if _lc_rec:
                with st.container(border=True):
                    st.markdown('<div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:10px">Edit Purchased License</div>', unsafe_allow_html=True)
                    _ec1, _ec2 = st.columns(2)
                    _e_tool  = _ec1.text_input("Tool Name *", value=_lc_rec["tool_name"], key="lc_e_tool")
                    _e_seats = _ec2.number_input("No. of Licenses *", min_value=1, value=int(_lc_rec["no_of_licenses"]), step=1, key="lc_e_seats")
                    _ec3, _ec4 = st.columns(2)
                    _e_start_dt = _ec3.date_input("Start Date", value=_parse_ymd(_lc_rec["start_date"]), key="lc_e_start", format="YYYY-MM-DD")
                    _e_end_dt   = _ec4.date_input("End Date", value=_parse_ymd(_lc_rec["end_date"]), key="lc_e_end", format="YYYY-MM-DD")
                    _e_start = _e_start_dt.strftime("%Y-%m-%d") if _e_start_dt else ""
                    _e_end   = _e_end_dt.strftime("%Y-%m-%d") if _e_end_dt else ""
                    _eb1, _eb2 = st.columns([1, 4])
                    if _eb1.button("Save Changes", type="primary", key="lc_save_edit"):
                        if not _e_tool.strip():
                            st.error("Tool name is required.")
                        else:
                            auth.update_license(st.session_state.lc_edit_id, _e_tool, int(_e_seats), _e_start, _e_end)
                            save_to_excel(st.session_state.projects)
                            st.session_state.lc_edit_id = None
                            st.session_state.toast = {"msg": "License updated!", "type": "success"}
                            st.rerun()
                    if _eb2.button("Cancel", key="lc_cancel_edit"):
                        st.session_state.lc_edit_id = None
                        st.rerun()

        # ── Add License form ─────────────────────────────────────────────────
        with st.expander("Add Purchased License", expanded=False):
            _lc1, _lc2 = st.columns(2)
            _n_tool  = _lc1.text_input("Tool Name *", key="lc_n_tool")
            _n_seats = _lc2.number_input("No. of Licenses *", min_value=1, value=1, step=1, key="lc_n_seats")
            _lc3, _lc4 = st.columns(2)
            _n_start_dt = _lc3.date_input("Start Date (optional)", value=None, key="lc_n_start", format="YYYY-MM-DD")
            _n_end_dt   = _lc4.date_input("End Date (optional)", value=None, key="lc_n_end", format="YYYY-MM-DD")
            _n_start = _n_start_dt.strftime("%Y-%m-%d") if _n_start_dt else ""
            _n_end   = _n_end_dt.strftime("%Y-%m-%d") if _n_end_dt else ""
            if st.button("Add License", type="primary", key="lc_add_btn"):
                if not _n_tool.strip():
                    st.error("Tool name is required.")
                else:
                    auth.create_license(_n_tool, int(_n_seats), _n_start, _n_end)
                    save_to_excel(st.session_state.projects)
                    st.session_state.toast = {"msg": f'License "{_n_tool}" added!', "type": "success"}
                    st.rerun()

        # ── Purchased License table ──────────────────────────────────────────
        st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 12px"><b>{len(_licenses_all)}</b> license(s) tracked</p>', unsafe_allow_html=True)
        if not _licenses_all:
            st.info("No licenses added yet. Use the form above to add one.")
        else:
            _lhdr = st.columns([0.3, 2.5, 1.2, 1.5, 1.5, 1.4, 0.4, 0.4])
            for _lc, _ll in zip(_lhdr, ["#", "Tool Name", "No. of Licenses", "Start Date", "End Date", "Status", "", ""]):
                _lc.markdown(f'<div style="font-size:9px;font-weight:600;text-transform:uppercase;color:#94A3B8;'
                             f'letter-spacing:.6px;padding:5px 0;border-bottom:2px solid #E2E8F0">{_ll}</div>',
                             unsafe_allow_html=True)
            for _lic in _licenses_all:
                _lr = st.columns([0.3, 2.5, 1.2, 1.5, 1.5, 1.4, 0.4, 0.4])
                _lr[0].markdown(cell(_lic["id"], size="10px", color="#94A3B8"), unsafe_allow_html=True)
                _lr[1].markdown(f'<span style="font-size:13px;font-weight:700;color:#111827">{esc(_lic["tool_name"])}</span>', unsafe_allow_html=True)
                _lr[2].markdown(f'<span style="font-size:13px;font-weight:600;color:#2563EB">{_lic["no_of_licenses"]}</span>', unsafe_allow_html=True)
                _lr[3].markdown(cell(_lic["start_date"] or "—", size="12px", color="#64748B"), unsafe_allow_html=True)
                _lr[4].markdown(cell(_lic["end_date"] or "—", size="12px", color="#64748B"), unsafe_allow_html=True)
                _lr[5].markdown(_lc_expiry_badge(_lic["end_date"]), unsafe_allow_html=True)
                if role == "admin":
                    with _lr[6]:
                        st.markdown('<span class="act-edit-marker"></span>', unsafe_allow_html=True)
                        if st.button("✏", key=f"lc_e_{_lic['id']}", help="Edit license", use_container_width=True):
                            st.session_state.lc_edit_id = _lic["id"]
                            st.session_state.sl_edit_id = None
                            st.rerun()
                    with _lr[7]:
                        st.markdown('<span class="act-del-marker"></span>', unsafe_allow_html=True)
                        if st.button("🗑", key=f"lc_d_{_lic['id']}", help="Delete license", use_container_width=True):
                            auth.delete_license(_lic["id"])
                            save_to_excel(st.session_state.projects)
                            st.session_state.toast = {"msg": f'License "{_lic["tool_name"]}" deleted.', "type": "info"}
                            st.rerun()

    # ══════════════════════════════════════════════════════════════════════════
    # SUB-TAB 2 — SOLD LICENSE
    # ══════════════════════════════════════════════════════════════════════════
    with _lc_tab2:
        # ── Edit form ────────────────────────────────────────────────────────
        if st.session_state.sl_edit_id is not None:
            _sl_rec = next((x for x in _sold_licenses_all if x["id"] == st.session_state.sl_edit_id), None)
            if _sl_rec:
                with st.container(border=True):
                    st.markdown('<div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:10px">Edit Sold License</div>', unsafe_allow_html=True)
                    _se1, _se2 = st.columns(2)
                    _sl_tool_opts = _purchased_tool_names or [""]
                    _sl_e_tool_idx = _sl_tool_opts.index(_sl_rec["tool_name"]) if _sl_rec["tool_name"] in _sl_tool_opts else 0
                    _sl_e_tool   = _se1.selectbox("Tool Name *", _sl_tool_opts, index=_sl_e_tool_idx, key="sl_e_tool")
                    _sl_e_client = _se2.text_input("Client *", value=_sl_rec["client"], key="sl_e_client")
                    _se3, _se4 = st.columns(2)
                    _sl_e_seats  = _se3.number_input("No. of Licenses *", min_value=1, value=int(_sl_rec["no_of_licenses"]), step=1, key="sl_e_seats")
                    _sl_e_notes  = _se4.text_input("Notes", value=_sl_rec["notes"], key="sl_e_notes")
                    _se5, _se6 = st.columns(2)
                    _sl_e_start_dt = _se5.date_input("Start Date", value=_parse_ymd(_sl_rec["start_date"]), key="sl_e_start", format="YYYY-MM-DD")
                    _sl_e_end_dt   = _se6.date_input("End Date", value=_parse_ymd(_sl_rec["end_date"]), key="sl_e_end", format="YYYY-MM-DD")
                    _sl_e_start = _sl_e_start_dt.strftime("%Y-%m-%d") if _sl_e_start_dt else ""
                    _sl_e_end   = _sl_e_end_dt.strftime("%Y-%m-%d") if _sl_e_end_dt else ""
                    _sb1, _sb2 = st.columns([1, 4])
                    if _sb1.button("Save Changes", type="primary", key="sl_save_edit"):
                        if not _sl_e_tool or not _sl_e_client.strip():
                            st.error("Tool name and client are required.")
                        else:
                            auth.update_sold_license(st.session_state.sl_edit_id, _sl_e_tool,
                                                     _sl_e_client, int(_sl_e_seats),
                                                     _sl_e_start, _sl_e_end, _sl_e_notes)
                            save_to_excel(st.session_state.projects)
                            st.session_state.sl_edit_id = None
                            st.session_state.toast = {"msg": "Sold license updated!", "type": "success"}
                            st.rerun()
                    if _sb2.button("Cancel", key="sl_cancel_edit"):
                        st.session_state.sl_edit_id = None
                        st.rerun()

        # ── Add Sold License form ────────────────────────────────────────────
        with st.expander("Add Sold License", expanded=False):
            if not _purchased_tool_names:
                st.info("Add at least one purchased license first — the tool name list comes from there.")
            else:
                _sa1, _sa2 = st.columns(2)
                _sl_n_tool   = _sa1.selectbox("Tool Name *", _purchased_tool_names, key="sl_n_tool")
                _sl_n_client = _sa2.text_input("Client *", key="sl_n_client")
                _sa3, _sa4 = st.columns(2)
                _sl_n_seats  = _sa3.number_input("No. of Licenses *", min_value=1, value=1, step=1, key="sl_n_seats")
                _sl_n_notes  = _sa4.text_input("Notes (optional)", key="sl_n_notes")
                _sa5, _sa6 = st.columns(2)
                _sl_n_start_dt = _sa5.date_input("Start Date (optional)", value=None, key="sl_n_start", format="YYYY-MM-DD")
                _sl_n_end_dt   = _sa6.date_input("End Date (optional)", value=None, key="sl_n_end", format="YYYY-MM-DD")
                _sl_n_start = _sl_n_start_dt.strftime("%Y-%m-%d") if _sl_n_start_dt else ""
                _sl_n_end   = _sl_n_end_dt.strftime("%Y-%m-%d") if _sl_n_end_dt else ""
                if st.button("Add Sold License", type="primary", key="sl_add_btn"):
                    if not _sl_n_client.strip():
                        st.error("Client is required.")
                    else:
                        auth.create_sold_license(_sl_n_tool, _sl_n_client, int(_sl_n_seats),
                                                 _sl_n_start, _sl_n_end, _sl_n_notes)
                        save_to_excel(st.session_state.projects)
                        st.session_state.toast = {"msg": f'Sold license "{_sl_n_tool}" added!', "type": "success"}
                        st.rerun()

        # ── Sold License table ───────────────────────────────────────────────
        st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 12px"><b>{len(_sold_licenses_all)}</b> sold license record(s)</p>', unsafe_allow_html=True)
        if not _sold_licenses_all:
            st.info("No sold licenses recorded yet. Use the form above to add one.")
        else:
            _slhdr = st.columns([0.3, 2.0, 2.0, 1.0, 1.4, 1.4, 1.4, 1.8, 0.4, 0.4])
            for _slc, _sll in zip(_slhdr, ["#", "Tool Name", "Client", "Licenses", "Start Date", "End Date", "Status", "Notes", "", ""]):
                _slc.markdown(f'<div style="font-size:9px;font-weight:600;text-transform:uppercase;color:#94A3B8;'
                              f'letter-spacing:.6px;padding:5px 0;border-bottom:2px solid #E2E8F0">{_sll}</div>',
                              unsafe_allow_html=True)
            for _sl in _sold_licenses_all:
                _slr = st.columns([0.3, 2.0, 2.0, 1.0, 1.4, 1.4, 1.4, 1.8, 0.4, 0.4])
                _slr[0].markdown(cell(_sl["id"], size="10px", color="#94A3B8"), unsafe_allow_html=True)
                _slr[1].markdown(f'<span style="font-size:12px;font-weight:700;color:#111827">{esc(_sl["tool_name"])}</span>', unsafe_allow_html=True)
                _slr[2].markdown(f'<span style="font-size:12px;color:#374151">{esc(_sl["client"])}</span>', unsafe_allow_html=True)
                _slr[3].markdown(f'<span style="font-size:13px;font-weight:600;color:#2563EB">{_sl["no_of_licenses"]}</span>', unsafe_allow_html=True)
                _slr[4].markdown(cell(_sl["start_date"] or "—", size="12px", color="#64748B"), unsafe_allow_html=True)
                _slr[5].markdown(cell(_sl["end_date"] or "—", size="12px", color="#64748B"), unsafe_allow_html=True)
                _slr[6].markdown(_lc_expiry_badge(_sl["end_date"]), unsafe_allow_html=True)
                _slr[7].markdown(cell(_sl["notes"] or "—", size="11px", color="#64748B"), unsafe_allow_html=True)
                if role == "admin":
                    with _slr[8]:
                        st.markdown('<span class="act-edit-marker"></span>', unsafe_allow_html=True)
                        if st.button("✏", key=f"sl_e_{_sl['id']}", help="Edit sold license", use_container_width=True):
                            st.session_state.sl_edit_id = _sl["id"]
                            st.session_state.lc_edit_id = None
                            st.rerun()
                    with _slr[9]:
                        st.markdown('<span class="act-del-marker"></span>', unsafe_allow_html=True)
                        if st.button("🗑", key=f"sl_d_{_sl['id']}", help="Delete sold license", use_container_width=True):
                            auth.delete_sold_license(_sl["id"])
                            save_to_excel(st.session_state.projects)
                            st.session_state.toast = {"msg": f'Sold license deleted.', "type": "info"}
                            st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# TAB: AI AGENT
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "agent" and role in ("admin", "lead", "manager"):
    api_key = get_api_key()
    if not api_key:
        api_key = st.text_input("Anthropic API Key", type="password",
                                help="Get your key at console.anthropic.com")

    if not api_key:
        st.info("Enter your Anthropic API Key above to use the AI Agent.")
    else:
        # Chat history
        for msg in st.session_state.messages:
            content = md_to_html(msg["content"])
            if msg["role"] == "user":
                st.markdown(_TMPL_USER_MSG.render(content=content), unsafe_allow_html=True)
            else:
                st.markdown(_TMPL_BOT_MSG.render(content=content), unsafe_allow_html=True)

        # Quick question buttons
        st.markdown('<div style="margin:12px 0 6px;font-size:10px;font-weight:600;color:#94A3B8;text-transform:uppercase;letter-spacing:.6px">Quick Questions</div>', unsafe_allow_html=True)
        quick_qs = [
            "Which projects are In Progress?",
            "Show team workload summary",
            "How many UAT projects?",
            "List all TEPL projects",
            "What is the ROI formula?",
        ]
        qcols = st.columns(len(quick_qs))
        for col, q in zip(qcols, quick_qs):
            if col.button(q, key=f"qq_{q[:14]}", use_container_width=True):
                st.session_state.messages.append({"role":"user","content":q})
                st.markdown(_TMPL_TYPING.render(), unsafe_allow_html=True)
                try:
                    reply = call_claude(api_key, st.session_state.messages, df)
                    st.session_state.messages.append({"role":"assistant","content":reply})
                except Exception as e:
                    st.session_state.messages.append({"role":"assistant","content":f"Error: {e}"})
                st.rerun()

        # Chat input
        user_input = st.chat_input("Ask anything about projects, team, ROI…")
        if user_input:
            st.session_state.messages.append({"role":"user","content":user_input})
            st.markdown(_TMPL_TYPING.render(), unsafe_allow_html=True)
            try:
                reply = call_claude(api_key, st.session_state.messages, df)
                st.session_state.messages.append({"role":"assistant","content":reply})
            except Exception as e:
                st.session_state.messages.append({"role":"assistant","content":f"Error: {e}"})
            st.rerun()

        if st.button("Clear Chat", key="clear_chat"):
            st.session_state.messages = [st.session_state.messages[0]]
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# TAB: USER MANAGEMENT  (admin only)
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "users" and role == "admin":
    st.markdown('<h2 style="font-size:20px;font-weight:700;color:#0F172A;margin-bottom:4px">User Management</h2>', unsafe_allow_html=True)
    st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:16px">Create accounts, assign roles, and manage password resets</p>', unsafe_allow_html=True)

    _users_cache = auth.get_all_users()

    # ── Import users from Excel ───────────────────────────────────────────────
    with st.expander("Import Users from Excel", expanded=False):
        st.markdown(
            '<p style="color:#64748B;font-size:12px;margin-bottom:10px">'
            'Upload an Excel file with columns: <b>Name</b>, <b>Email</b>, <b>Password</b>, '
            '<b>Role</b> (optional — defaults to <i>employee</i>). '
            'Existing users (same email) will be skipped.</p>',
            unsafe_allow_html=True
        )
        _tmpl_df = pd.DataFrame([
            {"Name": "Alice Smith", "Email": "alice@example.com", "Password": "Alice@123", "Role": "employee"},
            {"Name": "Bob Lead",    "Email": "bob@example.com",   "Password": "Bob@456",   "Role": "lead"},
        ])
        import io as _io
        _tmpl_buf = _io.BytesIO()
        _tmpl_df.to_excel(_tmpl_buf, index=False, engine="openpyxl")
        st.download_button(
            "Download Template", data=_tmpl_buf.getvalue(),
            file_name="users_import_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_user_tmpl"
        )
        _uploaded = st.file_uploader("Upload Excel", type=["xlsx", "xls"], key="import_users_file")
        if _uploaded:
            try:
                _imp_df = pd.read_excel(_uploaded, dtype=str, engine="openpyxl").fillna("")
                _imp_df.columns = [c.strip() for c in _imp_df.columns]
                _required = {"Name", "Email", "Password"}
                if not _required.issubset(set(_imp_df.columns)):
                    st.error(f"Missing columns. Required: {', '.join(_required)}")
                else:
                    _existing_emails = {u["email"].strip().lower() for u in _users_cache}
                    _created, _skipped, _errors = [], [], []
                    _pw_map = {}
                    for _, _ir in _imp_df.iterrows():
                        _iname  = str(_ir["Name"]).strip()
                        _iemail = str(_ir["Email"]).strip().lower()
                        _ipass  = str(_ir["Password"]).strip()
                        _irole  = str(_ir.get("Role", "employee")).strip().lower()
                        if _irole not in auth.ROLES:
                            _irole = "employee"
                        if not _iname or not _iemail or "@" not in _iemail:
                            _errors.append(f"Invalid row: {_iname} / {_iemail}")
                            continue
                        if len(_ipass) < 6:
                            _errors.append(f"Password too short for {_iemail} (min 6 chars)")
                            continue
                        if _iemail in _existing_emails:
                            _skipped.append(_iemail)
                            continue
                        try:
                            auth.create_user(_iname, _iemail, _ipass, _irole)
                            _pw_map[_iemail] = _ipass
                            _created.append(_iname)
                            _existing_emails.add(_iemail)
                        except Exception as _ex:
                            _errors.append(f"{_iemail}: {_ex}")
                    if _pw_map:
                        sync_users_excel(_pw_map)
                        save_to_excel(st.session_state.projects)
                    if _created:
                        st.success(f"Created {len(_created)} user(s): {', '.join(_created)}")
                    if _skipped:
                        st.info(f"Skipped {len(_skipped)} existing email(s): {', '.join(_skipped)}")
                    for _e in _errors:
                        st.error(_e)
                    if _created:
                        st.rerun()
            except Exception as _ex:
                st.error(f"Could not read file: {_ex}")

    # ── Create user form ──────────────────────────────────────────────────────
    with st.expander("Create New User", expanded=False):
        with st.container():
            ua, ub = st.columns(2)
            nu_name  = ua.text_input("Full Name *",      key="nu_name")
            nu_email = ub.text_input("Email Address *",  key="nu_email")
            uc2, ud = st.columns(2)
            nu_pass  = uc2.text_input("Password *",      type="password", key="nu_pass")
            nu_role  = ud.selectbox("Role",              auth.ROLES, key="nu_role")
            if st.button("Create User", type="primary", key="create_user_btn"):
                _errs = []
                if not nu_name.strip():                        _errs.append("Name is required.")
                if not nu_email.strip() or "@" not in nu_email: _errs.append("Valid email is required.")
                if not nu_pass or len(nu_pass) < 6:            _errs.append("Password must be at least 6 characters.")
                if _errs:
                    for _e in _errs: st.error(_e)
                else:
                    try:
                        auth.create_user(nu_name.strip(), nu_email.strip(), nu_pass, nu_role)
                        sync_users_excel({nu_email.strip().lower(): nu_pass})
                        save_to_excel(st.session_state.projects)
                        st.session_state.toast = {"msg": f'User "{nu_name.strip()}" created!', "type": "success"}
                        st.rerun()
                    except Exception as _ex:
                        st.error(f"Could not create user: {_ex}")

    # ── Edit user form (shown when a row's Edit button is clicked) ───────────
    if st.session_state.user_edit_id is not None:
        _eu_all  = _users_cache
        _eu_rec  = next((u for u in _eu_all if u["id"] == st.session_state.user_edit_id), None)
        if _eu_rec:
            with st.container(border=True):
                st.markdown(f'<div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:10px">Edit User — <span style="color:#2563EB">{esc(_eu_rec["name"])}</span></div>', unsafe_allow_html=True)
                _ea, _eb = st.columns(2)
                _eu_name  = _ea.text_input("Full Name *",    value=_eu_rec["name"],  key="eu_name")
                _eu_email = _eb.text_input("Email *",        value=_eu_rec["email"], key="eu_email")
                _ec, _ed  = st.columns(2)
                _eu_role  = _ec.selectbox("Role", auth.ROLES,
                                          index=auth.ROLES.index(_eu_rec["role"]) if _eu_rec["role"] in auth.ROLES else 0,
                                          key="eu_role")
                _ed.write("")
                _es1, _es2 = st.columns([1, 4])
                if _es1.button("Save", type="primary", key="eu_save"):
                    _errs = []
                    if not _eu_name.strip():                           _errs.append("Name is required.")
                    if not _eu_email.strip() or "@" not in _eu_email:  _errs.append("Valid email is required.")
                    if _errs:
                        for _e in _errs: st.error(_e)
                    else:
                        try:
                            auth.update_user(st.session_state.user_edit_id, _eu_name, _eu_email, _eu_role)
                            sync_users_excel()
                            save_to_excel(st.session_state.projects)
                            st.session_state.user_edit_id = None
                            st.session_state.toast = {"msg": f'User "{_eu_name.strip()}" updated!', "type": "success"}
                            st.rerun()
                        except Exception as _ex:
                            st.error(f"Could not update user: {_ex}")
                if _es2.button("Cancel", key="eu_cancel"):
                    st.session_state.user_edit_id = None
                    st.rerun()
            st.markdown("---")

    # ── Password reset form (shown when a row's Reset button is clicked) ──────
    _rp_uid = st.session_state.get("reset_pwd_uid")
    if _rp_uid:
        _rp_users = _users_cache
        _rp_user  = next((u for u in _rp_users if u["id"] == _rp_uid), None)
        if _rp_user:
            with st.container(border=True):
                st.markdown(f'<div style="font-size:13px;font-weight:700;color:#0F172A;margin-bottom:8px">Reset password for <span style="color:#2563EB">{esc(_rp_user["name"])}</span></div>', unsafe_allow_html=True)
                rpa, rpb = st.columns([2, 1])
                _new_pwd = rpa.text_input("New Password (min 6 chars)", type="password", key="rp_new_pwd")
                rpb.write("")
                rpc, rpd = st.columns(2)
                if rpc.button("Save Password", type="primary", key="rp_save"):
                    if _new_pwd and len(_new_pwd) >= 6:
                        auth.reset_password(_rp_uid, _new_pwd)
                        sync_users_excel({_rp_user["email"].strip().lower(): _new_pwd})
                        st.session_state.reset_pwd_uid = None
                        st.session_state.toast = {"msg": "Password reset successfully!", "type": "success"}
                        st.rerun()
                    else:
                        st.error("Password must be at least 6 characters.")
                if rpd.button("Cancel", key="rp_cancel"):
                    st.session_state.reset_pwd_uid = None
                    st.rerun()
            st.markdown("---")

    # ── Users table ───────────────────────────────────────────────────────────
    _all_users = _users_cache
    st.markdown(f'<p style="color:#64748B;font-size:12px;margin:6px 0 10px"><b>{len(_all_users)}</b> registered users</p>', unsafe_allow_html=True)

    _uhdr = st.columns([0.3, 1.6, 2.2, 1.0, 0.7, 0.45, 0.45, 0.45, 0.45])
    for _col, _lbl in zip(_uhdr, ["ID", "Name", "Email", "Role", "Active", "", "", "", ""]):
        _col.markdown(f'<div style="font-size:9px;font-weight:600;text-transform:uppercase;color:#94A3B8;letter-spacing:.6px;padding:5px 0;border-bottom:2px solid #E2E8F0">{_lbl}</div>', unsafe_allow_html=True)

    _role_colors = {"admin": "#1D4ED8", "lead": "#065F46", "manager": "#92400E", "employee": "#374151", "sales": "#0369A1"}
    for _u in _all_users:
        _uc = st.columns([0.3, 1.6, 2.2, 1.0, 0.7, 0.45, 0.45, 0.45, 0.45])
        _uc[0].markdown(cell(_u["id"], size="10px", color="#94A3B8"), unsafe_allow_html=True)
        _uc[1].markdown(f'<span style="font-size:12px;font-weight:600;color:#111827">{esc(_u["name"])}</span>', unsafe_allow_html=True)
        _uc[2].markdown(cell(_u["email"]), unsafe_allow_html=True)
        _rc = _role_colors.get(_u["role"], "#374151")
        _uc[3].markdown(f'<span style="font-size:11px;font-weight:700;color:{_rc}">{_u["role"].upper()}</span>', unsafe_allow_html=True)
        _uc[4].markdown(f'<span style="font-size:11px;font-weight:700;color:{"#10B981" if _u["is_active"] else "#EF4444"}">{"Yes" if _u["is_active"] else "No"}</span>', unsafe_allow_html=True)

        with _uc[5]:
            st.markdown('<span class="act-edit-marker"></span>', unsafe_allow_html=True)
            if st.button("✏", key=f"eu_{_u['id']}", help="Edit user", use_container_width=True):
                st.session_state.user_edit_id = _u["id"]
                st.session_state.reset_pwd_uid = None
                st.rerun()

        with _uc[6]:
            st.markdown('<span class="act-warn-marker"></span>', unsafe_allow_html=True)
            if st.button("🔑", key=f"rp_{_u['id']}", help="Reset password", use_container_width=True):
                st.session_state.reset_pwd_uid = _u["id"]
                st.session_state.user_edit_id = None
                st.rerun()

        _tog_lbl = "🔒" if _u["is_active"] else "🔓"
        _tog_tip = "Deactivate" if _u["is_active"] else "Activate"
        with _uc[7]:
            st.markdown('<span class="act-warn-marker"></span>', unsafe_allow_html=True)
            if st.button(_tog_lbl, key=f"tog_{_u['id']}", help=_tog_tip, use_container_width=True):
                if _u["id"] != cu["id"]:
                    auth.set_active(_u["id"], not _u["is_active"])
                    sync_users_excel()
                    save_to_excel(st.session_state.projects)
                    st.session_state.toast = {"msg": f'User {"deactivated" if _u["is_active"] else "activated"}.', "type": "info"}
                    st.rerun()
                else:
                    st.warning("You cannot deactivate your own account.")

        with _uc[8]:
            st.markdown('<span class="act-del-marker"></span>', unsafe_allow_html=True)
            if st.button("🗑", key=f"du_{_u['id']}", help="Delete user", use_container_width=True):
                if _u["id"] != cu["id"]:
                    auth.delete_user(_u["id"])
                    sync_users_excel()
                    save_to_excel(st.session_state.projects)
                    st.session_state.toast = {"msg": f'User "{_u["name"]}" deleted.', "type": "info"}
                    st.rerun()
                else:
                    st.warning("You cannot delete your own account.")

# ══════════════════════════════════════════════════════════════════════════════
# TAB: TASKS  (all roles — employees see only their own tasks)
# ══════════════════════════════════════════════════════════════════════════════
elif st.session_state.active_tab == "tasks":
    _STAT_COLORS = {
        "Not Started": "#94A3B8", "In Progress": "#3B82F6",
        "Completed": "#10B981",   "On Hold": "#F59E0B",
    }

    if role == "employee":
        # ── Employee view: own tasks + progress update ────────────────────────
        st.markdown('<h2 style="font-size:20px;font-weight:700;color:#0F172A;margin-bottom:4px">My Tasks</h2>', unsafe_allow_html=True)
        st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:16px">Tasks assigned to you — update your progress here</p>', unsafe_allow_html=True)

        _my_tasks = auth.get_user_tasks(cu["id"])
        if not _my_tasks:
            st.info("No tasks assigned to you yet.")
        else:
            st.markdown(f'<p style="color:#64748B;font-size:12px;margin-bottom:12px"><b>{len(_my_tasks)}</b> task(s) assigned to you</p>', unsafe_allow_html=True)
            for _t in _my_tasks:
                with st.container(border=True):
                    _tl, _tr = st.columns([3, 1.2])
                    _pct = int(_t["progress"])
                    _bar_c = "#10B981" if _pct == 100 else "#3B82F6"
                    with _tl:
                        st.markdown(f'<div style="font-size:14px;font-weight:700;color:#111827;margin-bottom:4px">{esc(_t["title"])}</div>', unsafe_allow_html=True)
                        if _t["description"]:
                            st.markdown(f'<div style="font-size:12px;color:#64748B;margin-bottom:6px;font-style:italic">{esc(_t["description"])}</div>', unsafe_allow_html=True)
                        _date_meta = f'Assigned by: <b>{esc(_t["assigned_by"])}</b>'
                        if _t.get("start_date"):
                            _date_meta += f' &nbsp;·&nbsp; Start: <b>{esc(_t["start_date"])}</b>'
                        if _t.get("due_date"):
                            _date_meta += f' &nbsp;·&nbsp; Due: <b>{esc(_t["due_date"])}</b>'
                        st.markdown(f'<div style="font-size:11px;color:#64748B;margin-bottom:6px">{_date_meta}</div>', unsafe_allow_html=True)
                        st.markdown(f'<div class="progress-bar-outer"><div class="progress-bar-inner" style="width:{_pct}%;background:{_bar_c}"></div></div>'
                                    f'<div style="font-size:10px;color:#64748B;margin-top:2px">{_pct}% complete</div>', unsafe_allow_html=True)
                    with _tr:
                        _new_prog = st.slider("Progress %", 0, 100, _pct, step=5, key=f"prog_{_t['id']}")
                        _stat_idx = auth.TASK_STATUSES.index(_t["status"]) if _t["status"] in auth.TASK_STATUSES else 0
                        _new_stat = st.selectbox("Status", auth.TASK_STATUSES, index=_stat_idx, key=f"stat_{_t['id']}")
                    _new_comment = st.text_area(
                        "Comments / Notes",
                        value=_t.get("comment", ""),
                        key=f"comment_{_t['id']}",
                        height=80,
                        placeholder="Add a note, update, or blocker…",
                    )
                    if st.button("Save Update", type="primary", key=f"save_p_{_t['id']}", use_container_width=True):
                        auth.update_task_progress(_t["id"], _new_prog, _new_stat, _new_comment)
                        st.session_state.toast = {"msg": "Task updated!", "type": "success"}
                        st.rerun()

                    # ── Weekly Update (locked after submission) ────────────────
                    _wk_start = auth.get_week_start()
                    _wk_end_dt = date.fromisoformat(_wk_start) + timedelta(days=6)
                    _wk_label = (f"{date.fromisoformat(_wk_start).strftime('%d %b')} – "
                                 f"{_wk_end_dt.strftime('%d %b %Y')}")
                    st.markdown(
                        f'<div style="font-size:11px;font-weight:700;color:#475569;'
                        f'border-top:1px solid #E2E8F0;margin-top:10px;padding-top:10px">'
                        f'Weekly Update — {_wk_label}</div>',
                        unsafe_allow_html=True)
                    _existing_wc = auth.get_user_week_comment(_t["id"], cu["id"], _wk_start)
                    if _existing_wc:
                        st.markdown(
                            f'<div style="background:#F8FAFC;border:1px solid #CBD5E1;border-radius:8px;'
                            f'padding:10px 14px;font-size:12px;color:#64748B;margin-top:4px">'
                            f'<span style="font-weight:700;color:#10B981">Submitted ✓</span>&nbsp; '
                            f'{esc(_existing_wc["comment"])}'
                            f'<br><span style="font-size:10px;color:#94A3B8">{esc(_existing_wc["created_at"])}</span>'
                            f'</div>',
                            unsafe_allow_html=True)
                    else:
                        _wc_text = st.text_area(
                            "Weekly update", height=70,
                            key=f"wc_{_t['id']}",
                            placeholder="Describe your progress this week…",
                            label_visibility="collapsed")
                        if st.button("Submit Weekly Update", key=f"wc_sub_{_t['id']}",
                                     use_container_width=True):
                            if _wc_text.strip():
                                auth.add_task_comment(_t["id"], cu["id"], _wc_text, _wk_start)
                                st.session_state.toast = {"msg": "Weekly update submitted!", "type": "success"}
                                st.rerun()
                            else:
                                st.warning("Please enter a comment before submitting.")

    else:
        # ── Admin / Lead / Manager: create + view all tasks ───────────────────
        st.markdown('<h2 style="font-size:20px;font-weight:700;color:#0F172A;margin-bottom:4px">Task Management</h2>', unsafe_allow_html=True)
        st.markdown('<p style="color:#64748B;font-size:12px;margin-bottom:16px">Assign and track tasks for your team</p>', unsafe_allow_html=True)

        def _render_my_tasks_panel(key_prefix):
            _my_tasks = auth.get_user_tasks(cu["id"])
            if not _my_tasks:
                st.info("No tasks assigned to you yet.")
            else:
                st.markdown(f'<p style="color:#64748B;font-size:12px;margin-bottom:12px"><b>{len(_my_tasks)}</b> task(s) assigned to you</p>', unsafe_allow_html=True)
                for _t in _my_tasks:
                    with st.container(border=True):
                        _tl, _tr = st.columns([3, 1.2])
                        _pct   = int(_t["progress"])
                        _bar_c = "#10B981" if _pct == 100 else "#3B82F6"
                        with _tl:
                            st.markdown(f'<div style="font-size:14px;font-weight:700;color:#111827;margin-bottom:4px">{esc(_t["title"])}</div>', unsafe_allow_html=True)
                            if _t["description"]:
                                st.markdown(f'<div style="font-size:12px;color:#64748B;margin-bottom:6px;font-style:italic">{esc(_t["description"])}</div>', unsafe_allow_html=True)
                            _lmeta = f'Assigned by: <b>{esc(_t["assigned_by"])}</b>'
                            if _t.get("start_date"):
                                _lmeta += f' &nbsp;·&nbsp; Start: <b>{esc(_t["start_date"])}</b>'
                            if _t.get("due_date"):
                                _lmeta += f' &nbsp;·&nbsp; Due: <b>{esc(_t["due_date"])}</b>'
                            st.markdown(f'<div style="font-size:11px;color:#64748B;margin-bottom:6px">{_lmeta}</div>', unsafe_allow_html=True)
                            st.markdown(
                                f'<div class="progress-bar-outer"><div class="progress-bar-inner" style="width:{_pct}%;background:{_bar_c}"></div></div>'
                                f'<div style="font-size:10px;color:#64748B;margin-top:2px">{_pct}% complete</div>',
                                unsafe_allow_html=True)
                        with _tr:
                            _new_prog = st.slider("Progress %", 0, 100, _pct, step=5, key=f"{key_prefix}prog_{_t['id']}")
                            _stat_idx = auth.TASK_STATUSES.index(_t["status"]) if _t["status"] in auth.TASK_STATUSES else 0
                            _new_stat = st.selectbox("Status", auth.TASK_STATUSES, index=_stat_idx, key=f"{key_prefix}stat_{_t['id']}")
                    _new_comment = st.text_area(
                        "Comments / Notes",
                        value=_t.get("comment", ""),
                        key=f"{key_prefix}comment_{_t['id']}",
                        height=80,
                        placeholder="Add a note, update, or blocker…",
                    )
                    if st.button("Save Update", type="primary", key=f"{key_prefix}save_{_t['id']}", use_container_width=True):
                        auth.update_task_progress(_t["id"], _new_prog, _new_stat, _new_comment)
                        st.session_state.toast = {"msg": "Task updated!", "type": "success"}
                        st.rerun()

                    # ── Weekly Update ──────────────────────────────────────────
                    _wk_start = auth.get_week_start()
                    _wk_end_dt = date.fromisoformat(_wk_start) + timedelta(days=6)
                    _wk_label = (f"{date.fromisoformat(_wk_start).strftime('%d %b')} – "
                                 f"{_wk_end_dt.strftime('%d %b %Y')}")
                    st.markdown(
                        f'<div style="font-size:11px;font-weight:700;color:#475569;'
                        f'border-top:1px solid #E2E8F0;margin-top:10px;padding-top:10px">'
                        f'Weekly Update — {_wk_label}</div>',
                        unsafe_allow_html=True)
                    _existing_wc = auth.get_user_week_comment(_t["id"], cu["id"], _wk_start)
                    if _existing_wc:
                        st.markdown(
                            f'<div style="background:#F8FAFC;border:1px solid #CBD5E1;border-radius:8px;'
                            f'padding:10px 14px;font-size:12px;color:#64748B;margin-top:4px">'
                            f'<span style="font-weight:700;color:#10B981">Submitted ✓</span>&nbsp; '
                            f'{esc(_existing_wc["comment"])}'
                            f'<br><span style="font-size:10px;color:#94A3B8">{esc(_existing_wc["created_at"])}</span>'
                            f'</div>',
                            unsafe_allow_html=True)
                    else:
                        _wc_text = st.text_area(
                            "Weekly update", height=70,
                            key=f"{key_prefix}wc_{_t['id']}",
                            placeholder="Describe your progress this week…",
                            label_visibility="collapsed")
                        if st.button("Submit Weekly Update", key=f"{key_prefix}wc_sub_{_t['id']}",
                                     use_container_width=True):
                            if _wc_text.strip():
                                auth.add_task_comment(_t["id"], cu["id"], _wc_text, _wk_start)
                                st.session_state.toast = {"msg": "Weekly update submitted!", "type": "success"}
                                st.rerun()
                            else:
                                st.warning("Please enter a comment before submitting.")

        def _render_all_tasks_panel():
            with st.expander("Assign New Task", expanded=False):
                _assignable = auth.get_employees_and_leads()
                if not _assignable:
                    st.warning("No employee or lead accounts found. Create users under the Users tab first.")
                else:
                    _ta1, _ta2 = st.columns(2)
                    _nt_title = _ta1.text_input("Task Title *", key="nt_title")
                    _emp_opts  = [f"{_e['name']}  [{_e['role'].upper()}]  ({_e['email']})" for _e in _assignable]
                    _emp_sel   = _ta2.selectbox("Assign To *", _emp_opts, key="nt_emp",
                                                help="Employees and Leads are listed here.")
                    _nt_desc   = st.text_area("Description (optional)", key="nt_desc", height=80)
                    _ta3, _ta4, _ta5 = st.columns(3)
                    _nt_start_dt = _ta3.date_input("Start Date (optional)", value=None, key="nt_start_dt", format="YYYY-MM-DD")
                    _nt_due_dt   = _ta4.date_input("Due Date (optional)", value=None, key="nt_due_dt", format="YYYY-MM-DD")
                    _ta5.text_input("Assigned By", value=cu["name"], disabled=True, key="nt_assigned_by")
                    _nt_start = _nt_start_dt.strftime("%Y-%m-%d") if _nt_start_dt else ""
                    _nt_due   = _nt_due_dt.strftime("%Y-%m-%d") if _nt_due_dt else ""
                    if st.button("Assign Task", type="primary", key="assign_task_btn"):
                        if not _nt_title.strip():
                            st.error("Task title is required.")
                        else:
                            _sel_idx = _emp_opts.index(_emp_sel)
                            _sel_emp = _assignable[_sel_idx]
                            auth.create_task(_nt_title, _nt_desc or "", _sel_emp["id"], cu["id"], _nt_due.strip(), _nt_start.strip())
                            st.session_state.toast = {"msg": f'Task assigned to {_sel_emp["name"]}!', "type": "success"}
                            st.rerun()

            _all_tasks = auth.get_all_tasks()
            st.markdown(
                f'<p style="color:#64748B;font-size:12px;margin:6px 0 12px">'
                f'<b>{len(_all_tasks)}</b> total tasks</p>',
                unsafe_allow_html=True)

            # ── Comment date-range filter ──────────────────────────────────────
            with st.container(border=True):
                st.markdown('<div style="font-size:11px;font-weight:700;color:#475569;margin-bottom:8px">Weekly Comment Filter</div>', unsafe_allow_html=True)
                _cf1, _cf2 = st.columns(2)
                _cm_from_dt = _cf1.date_input(
                    "From (week start)", key="cm_from",
                    value=date.today() - timedelta(weeks=4),
                    format="YYYY-MM-DD")
                _cm_to_dt = _cf2.date_input(
                    "To (week start)", key="cm_to",
                    value=date.today(),
                    format="YYYY-MM-DD")
                _cm_from_str = _cm_from_dt.strftime("%Y-%m-%d") if _cm_from_dt else None
                _cm_to_str   = _cm_to_dt.strftime("%Y-%m-%d") if _cm_to_dt else None

            if not _all_tasks:
                st.info("No tasks yet. Use the form above to assign tasks to employees.")
            else:
                def _render_task_rows(tasks, tab_sfx):
                    if not tasks:
                        st.info("No tasks in this category.")
                        return
                    _thdr = st.columns([2.0, 1.6, 1.5, 1.4, 0.9, 1.0, 1.0, 0.4])
                    for _col, _lbl in zip(_thdr, ["Task", "Assigned To", "Assigned By", "Status", "Progress", "Start Date", "Due Date", ""]):
                        _col.markdown(f'<div style="font-size:9px;font-weight:600;text-transform:uppercase;color:#94A3B8;letter-spacing:.6px;padding:5px 0;border-bottom:2px solid #E2E8F0">{_lbl}</div>', unsafe_allow_html=True)

                    for _t in tasks:
                        _tc = st.columns([2.0, 1.6, 1.5, 1.4, 0.9, 1.0, 1.0, 0.4])
                        _tdesc = _t["description"]
                        _tdesc_short = (_tdesc[:50] + "…") if len(_tdesc) > 50 else _tdesc
                        _tc[0].markdown(
                            f'<span style="font-size:12px;font-weight:600;color:#111827">{esc(_t["title"])}</span>'
                            + (f'<br><span style="font-size:10px;color:#64748B">{esc(_tdesc_short)}</span>' if _tdesc_short else ""),
                            unsafe_allow_html=True)
                        _tc[1].markdown(
                            f'<span style="font-size:12px">{esc(_t["assigned_to"])}</span>'
                            f'<br><span style="font-size:10px;color:#64748B">{esc(_t["assigned_to_email"])}</span>',
                            unsafe_allow_html=True)
                        _tc[2].markdown(
                            f'<span style="font-size:12px;color:#374151">{esc(_t["assigned_by"])}</span>',
                            unsafe_allow_html=True)
                        _tsc = _STAT_COLORS.get(_t["status"], "#94A3B8")
                        _tc[3].markdown(f'<span style="font-size:11px;font-weight:700;color:{_tsc}">{esc(_t["status"])}</span>', unsafe_allow_html=True)
                        _tpct = int(_t["progress"])
                        _tbar = "#10B981" if _tpct == 100 else "#3B82F6"
                        _tc[4].markdown(
                            f'<div class="progress-bar-outer"><div class="progress-bar-inner" style="width:{_tpct}%;background:{_tbar}"></div></div>'
                            f'<div style="font-size:10px;color:#64748B">{_tpct}%</div>',
                            unsafe_allow_html=True)
                        _tc[5].markdown(cell(_t.get("start_date") or "—", size="11px", color="#64748B"), unsafe_allow_html=True)
                        _tc[6].markdown(cell(_t["due_date"] or "—", size="11px", color="#64748B"), unsafe_allow_html=True)
                        with _tc[7]:
                            st.markdown('<span class="act-del-marker"></span>', unsafe_allow_html=True)
                            if st.button("🗑", key=f"dt_{tab_sfx}_{_t['id']}", help="Delete task", use_container_width=True):
                                auth.delete_task(_t["id"])
                                st.session_state.toast = {"msg": "Task deleted.", "type": "info"}
                                st.rerun()

                        # ── Per-task weekly comments expander ──────────────────
                        _wc_list = auth.get_task_comments_with_users(
                            task_id=_t["id"], from_date=_cm_from_str, to_date=_cm_to_str)
                        with st.expander(f"Weekly Comments ({len(_wc_list)})", expanded=False):
                            if not _wc_list:
                                st.info("No weekly comments in the selected period.")
                            else:
                                _wch = st.columns([1.2, 3.5, 2.0, 1.8])
                                for _c, _l in zip(_wch, ["Week", "Comment", "Employee", "Submitted"]):
                                    _c.markdown(f'<div style="font-size:9px;font-weight:700;text-transform:uppercase;color:#94A3B8;border-bottom:1px solid #E2E8F0;padding-bottom:4px">{_l}</div>', unsafe_allow_html=True)
                                for _wc in _wc_list:
                                    _wr = st.columns([1.2, 3.5, 2.0, 1.8])
                                    _wk_d = date.fromisoformat(_wc["week_start"])
                                    _wk_end = _wk_d + timedelta(days=6)
                                    _wr[0].markdown(f'<span style="font-size:10px;color:#475569">{_wk_d.strftime("%d %b")}–{_wk_end.strftime("%d %b")}</span>', unsafe_allow_html=True)
                                    _wr[1].markdown(f'<span style="font-size:11px;color:#111827">{esc(_wc["comment"])}</span>', unsafe_allow_html=True)
                                    _wr[2].markdown(f'<span style="font-size:11px;color:#374151">{esc(_wc["user_name"])}</span>', unsafe_allow_html=True)
                                    _wr[3].markdown(f'<span style="font-size:10px;color:#94A3B8">{esc(_wc["created_at"][:16])}</span>', unsafe_allow_html=True)

                def _render_tab_with_filters(base_tasks, tab_sfx):
                    _emp_names = sorted({t["assigned_to"] for t in base_tasks})
                    _ff1, _ff2 = st.columns([1.5, 2.5])
                    _emp_f  = _ff1.selectbox("Employee", ["All"] + _emp_names, key=f"emp_f_{tab_sfx}")
                    _name_f = _ff2.text_input("Task name", placeholder="Filter by task title…", key=f"name_f_{tab_sfx}")
                    _visible = list(base_tasks)
                    if _emp_f != "All":
                        _visible = [t for t in _visible if t["assigned_to"] == _emp_f]
                    if _name_f.strip():
                        _nq = _name_f.strip().lower()
                        _visible = [t for t in _visible if _nq in t["title"].lower()]
                    st.markdown(
                        f'<p style="color:#64748B;font-size:11px;margin:4px 0 8px">'
                        f'<b>{len(_visible)}</b> task(s)</p>',
                        unsafe_allow_html=True)
                    _render_task_rows(_visible, tab_sfx)

                # ── Status sub-tabs ────────────────────────────────────────────
                _ip_tasks   = [t for t in _all_tasks if t["status"] == "In Progress"]
                _comp_tasks = [t for t in _all_tasks if t["status"] == "Completed"]
                _hold_tasks = [t for t in _all_tasks if t["status"] == "On Hold"]

                _stab_all, _stab_ip, _stab_comp, _stab_hold = st.tabs([
                    f"All ({len(_all_tasks)})",
                    f"In Progress ({len(_ip_tasks)})",
                    f"Completed ({len(_comp_tasks)})",
                    f"On Hold ({len(_hold_tasks)})",
                ])
                with _stab_all:
                    _render_tab_with_filters(_all_tasks, "all")
                with _stab_ip:
                    _render_tab_with_filters(_ip_tasks, "ip")
                with _stab_comp:
                    _render_tab_with_filters(_comp_tasks, "comp")
                with _stab_hold:
                    _render_tab_with_filters(_hold_tasks, "hold")

                # ── All comments summary (collapsible) ────────────────────────
                _all_comments = auth.get_task_comments_with_users(from_date=_cm_from_str, to_date=_cm_to_str)
                with st.expander(f"All Weekly Comments Summary ({len(_all_comments)} entries)", expanded=False):
                    if not _all_comments:
                        st.info("No comments in the selected period.")
                    else:
                        _ach = st.columns([1.5, 2.5, 2.5, 1.8, 1.8])
                        for _c, _l in zip(_ach, ["Week", "Task", "Employee", "Comment", "Submitted"]):
                            _c.markdown(f'<div style="font-size:9px;font-weight:700;text-transform:uppercase;color:#94A3B8;border-bottom:1px solid #E2E8F0;padding-bottom:4px">{_l}</div>', unsafe_allow_html=True)
                        for _ac in _all_comments:
                            _ar = st.columns([1.5, 2.5, 2.5, 1.8, 1.8])
                            _wk_d = date.fromisoformat(_ac["week_start"])
                            _wk_end = _wk_d + timedelta(days=6)
                            _ar[0].markdown(f'<span style="font-size:10px;color:#475569">{_wk_d.strftime("%d %b")}–{_wk_end.strftime("%d %b")}</span>', unsafe_allow_html=True)
                            _ar[1].markdown(f'<span style="font-size:11px;font-weight:600;color:#111827">{esc(_ac["task_title"])}</span>', unsafe_allow_html=True)
                            _ar[2].markdown(f'<span style="font-size:11px;color:#374151">{esc(_ac["user_name"])}</span>', unsafe_allow_html=True)
                            _ar[3].markdown(f'<span style="font-size:11px;color:#64748B">{esc(_ac["comment"][:80])}{"…" if len(_ac["comment"])>80 else ""}</span>', unsafe_allow_html=True)
                            _ar[4].markdown(f'<span style="font-size:10px;color:#94A3B8">{esc(_ac["created_at"][:16])}</span>', unsafe_allow_html=True)

        if role == "lead":
            _ltab_mine, _ltab_all = st.tabs(["My Tasks", "All Tasks"])
            with _ltab_mine:
                _render_my_tasks_panel(key_prefix="lt_")
            with _ltab_all:
                _render_all_tasks_panel()
        else:
            _render_all_tasks_panel()
