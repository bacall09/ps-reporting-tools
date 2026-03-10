"""
Workload Health Score
Upload a Smartsheets DRS export + NetSuite time detail to generate
a weighted workload score per consultant across active FF projects.
"""
import streamlit as st
import pandas as pd
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Workload Health Score", page_icon="📊", layout="wide")

# ── Constants ─────────────────────────────────────────────────────────────────
NAVY     = "1e2c63"
TEAL     = "4472C4"
WHITE    = "FFFFFF"
LTGRAY   = "F2F2F2"
RED      = "E74C3C"
GREEN    = "27AE60"
AMBER    = "F39C12"

# ── Employee / Region config (mirrors utilization report) ─────────────────────
EMPLOYEE_LOCATION = {
    "Arestarkhov, Yaroslav":  "Czech Republic",
    "Carpen, Anamaria":       "Spain",
    "Centinaje, Rhodechild":  "Manila (PH)",
    "Dolha, Madalina":        "Faroe Islands",
    "Dolha":                  "Faroe Islands",
    "Cooke, Ellen":           "Northern Ireland",
    "Cruz, Daniel":           "Manila (PH)",
    "DiMarco, Nicole R":      "USA",
    "Gardner, Cheryll L":     "USA",
    "Hopkins, Chris":         "USA",
    "Ickler, Georganne":      "USA",
    "Isberg, Eric":           "USA",
    "Jordanova, Marija":      "North Macedonia",
    "Lappin, Thomas":         "Northern Ireland",
    "Longalong, Santiago":    "Manila (PH)",
    "Mohammad, Manaan":       "Canada",
    "Morris, Lisa":           "Sydney (NSW)",
    "Pallone, Daniel":        "Sydney (NSW)",
    "NAQVI, SYED":            "Canada",
    "Raykova, Silvia":        "Netherlands",
    "Selvakumar, Sajithan":   "Canada",
    "Snee, Stefanie J":       "USA",
    "Stone, Matt":            "USA",
    "Tuazon, Carol":          "Manila (PH)",
    "Zoric, Ivan":            "Serbia",
    "Murphy, Conor":          "USA",
    "Bell, Stuart":           "USA",
    "Cloete":                 "Netherlands",
    "Cloete, Bronwyn":        "Netherlands",
    "Hamilton C":             "USA",
    "Hamilton, Julie C":      "USA",
    "Strauss, John W":        "USA",
    "Swanson":                "USA",
    "Barrio, Nairobi":  "USA",
    "Porangada, Suraj":  "USA",
    "Hughes, Madalyn":  "USA",
    "Olson, Austin D":  "USA",
    "Finalle-Newton, Jesse":  "USA",
    "Church, Jason G":  "USA",
}
PS_REGION_OVERRIDE = {
    "NAQVI, SYED":  "EMEA",
    "Cruz, Daniel": "NOAM",
}
PS_REGION_MAP = {
    "Sydney (NSW)":     "APAC",
    "Manila (PH)":      "APAC",
    "UK":               "EMEA",
    "Spain":            "EMEA",
    "Netherlands":      "EMEA",
    "Northern Ireland": "EMEA",
    "Faroe Islands":    "EMEA",
    "North Macedonia":  "EMEA",
    "Czech Republic":   "EMEA",
    "Serbia":           "EMEA",
    "USA":              "NOAM",
    "Canada":           "NOAM",
}

# ── Phase weights ─────────────────────────────────────────────────────────────
PHASE_WEIGHTS = {
    "00. onboarding":                    1.0,
    "01. requirements and design":       2.5,
    "02. configuration":                 3.0,
    "03. enablement/training":           2.0,
    "04. uat":                           1.5,
    "05. prep for go-live":              1.5,
    "06. go-live":                       1.5,
    "07. data migration":                1.0,
    "08. ready for support transition":  0.5,
    "09. phase 2 scoping":               1.0,
    "10. complete/pending final billing": 0.0,
    "11. on hold":                       0.25,
    "12. ps review":                     0.25,
}

# Phases excluded from active workload score
INACTIVE_PHASES = {"10. complete/pending final billing", "11. on hold"}

# ── Workload thresholds ───────────────────────────────────────────────────────
def workload_level(score):
    if score <= 25:   return "Low",    GREEN
    if score <= 60:   return "Medium", AMBER
    return "High", RED

# ── Client health multiplier ──────────────────────────────────────────────────
def client_health_multiplier(responsiveness, sentiment):
    r = str(responsiveness).strip().lower() if pd.notna(responsiveness) else ""
    s = str(sentiment).strip().lower() if pd.notna(sentiment) else ""
    if ("negative" in s or "unresponsive" in r) and ("negative" in s and "unresponsive" in r):
        return 1.3
    if "negative" in s or "unresponsive" in r or "not responsive" in r:
        return 1.2
    if "highly engaged" in r and "positive" in s:
        return 0.9
    return 1.0

# ── Risk multiplier ───────────────────────────────────────────────────────────
def risk_multiplier(risk_level):
    r = str(risk_level).strip().lower() if pd.notna(risk_level) else ""
    if "high"   in r: return 1.2
    if "medium" in r: return 1.05
    return 1.0

# ── Employee → PS Region lookup ───────────────────────────────────────────────
def get_ps_region(name):
    if not name or str(name).strip().lower() in ("", "nan", "none"):
        return "Unknown"
    name = str(name).strip()
    if name in PS_REGION_OVERRIDE:
        return PS_REGION_OVERRIDE[name]
    loc = EMPLOYEE_LOCATION.get(name)
    if not loc:
        last = name.split(",")[0].strip()
        loc = next((v for k, v in EMPLOYEE_LOCATION.items() if k.split(",")[0].strip().lower() == last.lower()), None)
    return PS_REGION_MAP.get(loc, "Unknown") if loc else "Unknown"

# ── Phase weight lookup ───────────────────────────────────────────────────────
def get_phase_weight(phase):
    if not phase or str(phase).strip().lower() in ("", "nan", "none"):
        return 1.0, "Undefined"
    p = str(phase).strip().lower()
    for key, weight in PHASE_WEIGHTS.items():
        if p == key or p.startswith(key[:8]):
            return weight, str(phase).strip()
    return 1.0, str(phase).strip()

# ── SS column map ─────────────────────────────────────────────────────────────
SS_COL_MAP = {
    "project name":          "project_name",
    "project id":            "project_id",
    "overall rag":           "rag",
    "start date":            "start_date",
    "go live date":          "go_live_date",
    "% complete":            "pct_complete",
    "project type":          "project_type",
    "status":                "status",
    "project phase":         "phase",
    "client responsiveness": "client_responsiveness",
    "client sentiment":      "client_sentiment",
    "risk level":            "risk_level",
    "schedule health":       "schedule_health",
    "resource health":       "resource_health",
    "scope health":          "scope_health",
    "territory":             "territory",
    "actual hours":          "actual_hours",
    "budgeted hours":        "budgeted_hours",
    "budget":                "budget",
    "change order":          "change_order",
    "partner name":          "partner_name",
    "on hold reason":        "on_hold_reason",
}

MILESTONE_COLS = [
    "Intro. Email Sent", "Standard Config Start", "Enablement Session",
    "Session #1", "Session #2", "Prod Cutover", "Hypercare Start",
    "Close Out Remaining Tasks", "UAT Signoff", "Transition to Support",
]

# NS column map (subset needed for PM join)
NS_COL_MAP = {
    "employee":        "employee",
    "project":         "project",
    "project id":      "project_id",
    "billing type":    "billing_type",
    "project manager": "project_manager",
}

# ── Excel helpers ─────────────────────────────────────────────────────────────
def hex_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def style_cell(cell, bg, bold=False, align="left", font_color="000000", size=10):
    cell.fill = hex_fill(bg)
    cell.font = Font(bold=bold, color=font_color, name="Calibri", size=size)
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)

def write_title(ws, title, ncols, subtitle=None):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    c = ws.cell(1, 1, title)
    style_cell(c, NAVY, bold=True, align="left", font_color=WHITE, size=12)
    ws.row_dimensions[1].height = 28
    if subtitle:
        ws.append([""] * ncols)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=ncols)
        sc = ws.cell(2, 1, subtitle)
        style_cell(sc, "EBF5FB", bold=False, align="left", font_color="1A5276", size=9)
        ws.row_dimensions[2].height = 16

def style_header(ws, row, headers, bg=None):
    bg = bg or TEAL
    for col, h in enumerate(headers, 1):
        c = ws.cell(row, col, h)
        style_cell(c, bg, bold=True, align="center", font_color=WHITE, size=10)
        ws.row_dimensions[row].height = 22

def rag_color(rag):
    r = str(rag).strip().lower() if pd.notna(rag) else ""
    if "red"    in r: return "FDECED"
    if "amber"  in r or "yellow" in r: return "FEF9E7"
    if "green"  in r: return "EAF9F1"
    return LTGRAY

def level_color(level):
    if level == "High":   return "FDECED"
    if level == "Medium": return "FEF9E7"
    if level == "Low":    return "EAF9F1"
    return LTGRAY

def border_thin():
    s = Side(style="thin", color="D0D0D0")
    return Border(left=s, right=s, top=s, bottom=s)

# ── Data loading ──────────────────────────────────────────────────────────────
def load_ss(file):
    """Load and normalise Smartsheets export."""
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    df.columns = df.columns.str.strip()

    # Rename via SS_COL_MAP
    rename = {c: SS_COL_MAP[c.lower()] for c in df.columns if c.lower() in SS_COL_MAP}
    df = df.rename(columns=rename)

    # Capture milestone columns
    milestone_present = [c for c in df.columns if c in MILESTONE_COLS]

    # Ensure project_id is string for joining
    if "project_id" in df.columns:
        df["project_id"] = df["project_id"].astype(str).str.strip()

    return df, milestone_present


def load_ns(file):
    """Load NetSuite actuals — extract project_id → project_manager map."""
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    df.columns = df.columns.str.strip()
    rename = {c: NS_COL_MAP[c.lower()] for c in df.columns if c.lower() in NS_COL_MAP}
    df = df.rename(columns=rename)

    needed = [c for c in ["project_id", "project_manager", "project", "billing_type"] if c in df.columns]
    df = df[needed].drop_duplicates()

    if "project_id" in df.columns:
        df["project_id"] = df["project_id"].astype(str).str.strip()

    return df


# ── Scoring engine ────────────────────────────────────────────────────────────
def score_projects(ss_df, ns_df):
    """
    Join SS + NS, compute per-project weighted score (FF only, excl inactive).
    Returns scored DataFrame.
    """
    df = ss_df.copy()

    # Filter FF only — exclude T&M
    if "project_type" in df.columns:
        tm_mask = df["project_type"].str.lower().str.contains("t&m|time.*material", na=False, regex=True)
        df = df[~tm_mask].copy()

    # Exclude inactive phases from score (still kept in data, score = 0)
    def is_active(phase):
        if not phase or str(phase).strip().lower() in ("", "nan", "none"):
            return True
        return str(phase).strip().lower() not in INACTIVE_PHASES

    # Join PM from NS
    if ns_df is not None and "project_id" in ns_df.columns and "project_manager" in ns_df.columns:
        pm_map = ns_df.dropna(subset=["project_manager"]).drop_duplicates("project_id").set_index("project_id")["project_manager"]
        df["project_manager"] = df["project_id"].map(pm_map)
        df["pm_flag"] = df["project_manager"].isna()
    else:
        df["project_manager"] = None
        df["pm_flag"] = True

    # Phase weight
    df["phase_weight"] = df["phase"].apply(lambda p: get_phase_weight(p)[0])

    # Multipliers
    resp_col = "client_responsiveness" if "client_responsiveness" in df.columns else None
    sent_col = "client_sentiment"      if "client_sentiment"      in df.columns else None
    risk_col = "risk_level"            if "risk_level"            in df.columns else None

    df["client_health_mult"] = df.apply(
        lambda r: client_health_multiplier(
            r[resp_col] if resp_col else None,
            r[sent_col] if sent_col else None
        ), axis=1
    )
    df["risk_mult"] = df[risk_col].apply(risk_multiplier) if risk_col else 1.0

    # Active = not on hold by phase OR status
    def is_active_row(row):
        phase_inactive = not is_active(row.get("phase", ""))
        status_onhold  = str(row.get("status", "")).strip().lower() in ("on-hold", "on hold", "onhold")
        return not phase_inactive and not status_onhold

    df["active"] = df.apply(is_active_row, axis=1)

    # Total projects = all FF rows (excl Complete/Pending Final Billing)
    complete_phases = {"10. complete/pending final billing"}
    df["total_project"] = df["phase"].apply(
        lambda p: str(p).strip().lower() not in complete_phases if pd.notna(p) else True
    )

    df["weighted_score"] = df.apply(
        lambda r: round(r["phase_weight"] * r["client_health_mult"] * r["risk_mult"], 2)
        if r["active"] else 0.0, axis=1
    )

    # PS Region from employee lookup (consultant = project_manager from NS)
    df["ps_region"] = df["project_manager"].apply(get_ps_region)

    return df


def build_consultant_summary(scored_df):
    """Aggregate scored projects by consultant."""
    active = scored_df[scored_df["active"]].copy()
    total  = scored_df[scored_df["total_project"]].copy() if "total_project" in scored_df.columns else scored_df.copy()

    # Active project count (excludes on-hold by phase or status)
    active_counts = active.groupby("project_manager")["project_id"].nunique().rename("active_project_count")
    # Total project count (all FF, excl complete)
    total_counts  = total.groupby("project_manager")["project_id"].nunique().rename("total_project_count")

    grp = active.groupby("project_manager").agg(
        total_score=("weighted_score", "sum"),
        ps_region=("ps_region", "first"),
    ).reset_index()

    grp = grp.join(active_counts, on="project_manager").join(total_counts, on="project_manager")
    grp["active_project_count"] = grp["active_project_count"].fillna(0).astype(int)
    grp["total_project_count"]  = grp["total_project_count"].fillna(0).astype(int)
    grp["total_score"] = grp["total_score"].round(1)
    grp["workload_level"] = grp["total_score"].apply(lambda s: workload_level(s)[0])
    grp = grp.sort_values("total_score", ascending=False).reset_index(drop=True)

    # Flag missing PM
    missing = scored_df[scored_df["pm_flag"] == True]["project_id"].nunique()
    return grp, missing


# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(scored_df, consultant_df, missing_pm_count, as_of):
    wb = Workbook()
    wb.remove(wb.active)

    # ── Tab 1: Dashboard ──────────────────────────────────────────────────────
    ws_dash = wb.create_sheet("Dashboard")
    ws_dash.sheet_properties.tabColor = NAVY

    ws_dash.column_dimensions["A"].width = 28
    ws_dash.column_dimensions["B"].width = 14
    ws_dash.column_dimensions["C"].width = 14
    ws_dash.column_dimensions["D"].width = 14
    ws_dash.column_dimensions["E"].width = 14

    # Title
    ws_dash.merge_cells("A1:E1")
    c = ws_dash.cell(1, 1, "WORKLOAD HEALTH SCORE — PS Team Overview")
    style_cell(c, NAVY, bold=True, font_color=WHITE, size=13)
    ws_dash.row_dimensions[1].height = 32

    ws_dash.merge_cells("A2:E2")
    c2 = ws_dash.cell(2, 1, f"Fixed Fee active projects only  ·  Data as of {as_of}")
    style_cell(c2, "EBF5FB", bold=False, font_color="1A5276", size=9)
    ws_dash.row_dimensions[2].height = 16

    # Key metrics row
    total_consultants  = len(consultant_df)
    high_count    = len(consultant_df[consultant_df["workload_level"] == "High"])
    medium_count  = len(consultant_df[consultant_df["workload_level"] == "Medium"])
    low_count     = len(consultant_df[consultant_df["workload_level"] == "Low"])
    active_projects = scored_df[scored_df["active"]]["project_id"].nunique()
    total_projects  = scored_df[scored_df.get("total_project", pd.Series(True, index=scored_df.index))]["project_id"].nunique() if "total_project" in scored_df.columns else active_projects

    ws_dash.row_dimensions[4].height = 18
    ws_dash.cell(4, 1, "KEY METRICS").font = Font(bold=True, color=NAVY, size=10, name="Calibri")

    metrics = [
        ("Consultants Scored",  total_consultants, LTGRAY),
        ("Total FF Projects",   total_projects,    LTGRAY),
        ("Active Projects",     active_projects,   LTGRAY),
        ("High Workload",       high_count,        "FDECED"),
        ("Medium Workload",     medium_count,      "FEF9E7"),
    ]
    for col, (label, val, bg) in enumerate(metrics, 1):
        ws_dash.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col)
        ws_dash.merge_cells(start_row=6, start_column=col, end_row=6, end_column=col)
        lc = ws_dash.cell(5, col, label)
        vc = ws_dash.cell(6, col, val)
        style_cell(lc, bg, bold=False, align="center", font_color="555555", size=9)
        style_cell(vc, bg, bold=True,  align="center", font_color=NAVY,     size=16)
        ws_dash.row_dimensions[5].height = 16
        ws_dash.row_dimensions[6].height = 36

    # Consultant summary table
    ws_dash.row_dimensions[8].height = 18
    ws_dash.cell(8, 1, "CONSULTANT WORKLOAD SUMMARY").font = Font(bold=True, color=NAVY, size=10, name="Calibri")

    dash_headers = ["Consultant", "PS Region", "Total Projects", "Active Projects", "Weighted Score", "Workload Level"]
    style_header(ws_dash, 9, dash_headers, TEAL)
    for col, w in enumerate([28, 14, 16, 16, 16, 16], 1):
        ws_dash.column_dimensions[get_column_letter(col)].width = w

    for r_idx, row in consultant_df.iterrows():
        r = 10 + r_idx
        level = row["workload_level"]
        bg = level_color(level)
        vals = [row["project_manager"], row["ps_region"],
                row.get("total_project_count", 0), row.get("active_project_count", 0),
                row["total_score"], level]
        for col, val in enumerate(vals, 1):
            c = ws_dash.cell(r, col, val)
            style_cell(c, bg, align="center" if col > 1 else "left")
            c.border = border_thin()
        ws_dash.row_dimensions[r].height = 16

    if missing_pm_count > 0:
        flag_row = 10 + len(consultant_df) + 2
        ws_dash.merge_cells(start_row=flag_row, start_column=1, end_row=flag_row, end_column=5)
        fc = ws_dash.cell(flag_row, 1,
            f"⚠  {missing_pm_count} project(s) have no PM assigned in NetSuite — update NS to include in scoring")
        style_cell(fc, "FEF9E7", bold=False, font_color="7D6608", size=9)

    ws_dash.freeze_panes = "A10"

    # ── Tab 2: By Consultant ──────────────────────────────────────────────────
    ws_con = wb.create_sheet("By Consultant")
    ws_con.sheet_properties.tabColor = TEAL

    con_headers = ["Consultant", "PS Region", "Total Projects", "Active Projects",
                   "Weighted Score", "Workload Level", "High Risk Projects", "Avg Score/Project"]
    con_widths   = [28, 14, 16, 16, 16, 16, 18, 18]
    write_title(ws_con, "WORKLOAD HEALTH SCORE — By Consultant", len(con_headers))
    style_header(ws_con, 2, con_headers, TEAL)
    ws_con.auto_filter.ref = f"A2:{get_column_letter(len(con_headers))}2"
    for i, w in enumerate(con_widths, 1):
        ws_con.column_dimensions[get_column_letter(i)].width = w

    # High risk count per consultant
    high_risk = scored_df[(scored_df["active"]) & (scored_df.get("risk_level", pd.Series(dtype=str)).str.lower().str.contains("high", na=False))]
    high_risk_count = high_risk.groupby("project_manager")["project_id"].nunique().to_dict() if "risk_level" in scored_df.columns else {}

    for r_idx, row in consultant_df.iterrows():
        r = 3 + r_idx
        level = row["workload_level"]
        bg = level_color(level)
        act  = row.get("active_project_count", 0)
        tot  = row.get("total_project_count", 0)
        avg  = round(row["total_score"] / act, 2) if act > 0 else 0
        hr   = high_risk_count.get(row["project_manager"], 0)
        vals = [row["project_manager"], row["ps_region"], tot, act,
                row["total_score"], level, hr, avg]
        for col, val in enumerate(vals, 1):
            c = ws_con.cell(r, col, val)
            style_cell(c, bg, align="center" if col > 1 else "left")
            c.border = border_thin()
        ws_con.row_dimensions[r].height = 16
    ws_con.freeze_panes = "A3"

    # ── Tab 3: By Project ─────────────────────────────────────────────────────
    ws_proj = wb.create_sheet("By Project")
    ws_proj.sheet_properties.tabColor = "8E44AD"

    proj_base = ["Project", "Project ID", "Project Type", "Phase", "Territory",
                 "Project Manager", "PS Region", "Status", "Overall RAG",
                 "Phase Weight", "Client Health Mult", "Risk Mult", "Weighted Score",
                 "Schedule Health", "Risk Level", "Client Responsiveness", "Client Sentiment",
                 "Go Live Date", "% Complete", "Active"]
    proj_widths = [38, 12, 22, 28, 12, 24, 12, 14, 12, 12, 16, 10, 14, 18, 12, 22, 18, 14, 10, 8]

    write_title(ws_proj, "WORKLOAD HEALTH SCORE — By Project (Fixed Fee Active)", len(proj_base))
    style_header(ws_proj, 2, proj_base, "8E44AD")
    ws_proj.auto_filter.ref = f"A2:{get_column_letter(len(proj_base))}2"
    for i, w in enumerate(proj_widths, 1):
        ws_proj.column_dimensions[get_column_letter(i)].width = w

    proj_sorted = scored_df.sort_values(["project_manager", "weighted_score"], ascending=[True, False])

    for r_idx, (_, row) in enumerate(proj_sorted.iterrows()):
        r = 3 + r_idx
        level_bg = level_color(workload_level(row.get("weighted_score", 0))[0])
        rag_bg   = rag_color(row.get("rag", ""))

        def gv(col): return row.get(col, "") if pd.notna(row.get(col, "")) else ""

        vals = [
            gv("project_name"), gv("project_id"), gv("project_type"), gv("phase"),
            gv("territory"), gv("project_manager"), gv("ps_region"),
            gv("status"), gv("rag"),
            gv("phase_weight"), gv("client_health_mult"), gv("risk_mult"), gv("weighted_score"),
            gv("schedule_health"), gv("risk_level"),
            gv("client_responsiveness"), gv("client_sentiment"),
            gv("go_live_date").strftime("%Y-%m-%d") if pd.notna(gv("go_live_date")) and hasattr(gv("go_live_date"), "strftime") else gv("go_live_date"),
            gv("pct_complete"),
            "Yes" if row.get("active") else "No",
        ]
        for col, val in enumerate(vals, 1):
            c = ws_proj.cell(r, col, val)
            bg = rag_bg if col == 9 else level_bg if col == 13 else LTGRAY if r_idx % 2 == 0 else WHITE
            style_cell(c, bg, align="center" if col > 2 else "left")
            c.border = border_thin()
        ws_proj.row_dimensions[r].height = 16
    ws_proj.freeze_panes = "A3"

    # ── Tab 4: At-Risk ────────────────────────────────────────────────────────
    ws_risk = wb.create_sheet("At-Risk Projects")
    ws_risk.sheet_properties.tabColor = RED

    at_risk_df = scored_df[
        (scored_df["active"]) & (
            (scored_df.get("risk_level",       pd.Series(dtype=str)).str.lower().str.contains("high",    na=False)) |
            (scored_df.get("schedule_health",  pd.Series(dtype=str)).str.lower().str.contains("behind",  na=False)) |
            (scored_df.get("rag",              pd.Series(dtype=str)).str.lower().str.contains("red",     na=False)) |
            (scored_df.get("client_sentiment", pd.Series(dtype=str)).str.lower().str.contains("negative",na=False))
        )
    ].sort_values("weighted_score", ascending=False)

    risk_headers = ["Project", "Project ID", "Project Manager", "PS Region", "Phase",
                    "Weighted Score", "Risk Level", "Schedule Health", "RAG",
                    "Client Sentiment", "Client Responsiveness", "Go Live Date"]
    risk_widths   = [38, 12, 24, 12, 28, 14, 12, 18, 10, 18, 22, 14]

    write_title(ws_risk, "AT-RISK PROJECTS — High Risk / Behind Schedule / Red RAG",
                len(risk_headers),
                f"{len(at_risk_df)} project(s) flagged")
    style_header(ws_risk, 3, risk_headers, RED)
    ws_risk.auto_filter.ref = f"A3:{get_column_letter(len(risk_headers))}3"
    for i, w in enumerate(risk_widths, 1):
        ws_risk.column_dimensions[get_column_letter(i)].width = w

    for r_idx, (_, row) in enumerate(at_risk_df.iterrows()):
        r = 4 + r_idx
        def gv(col): return row.get(col, "") if pd.notna(row.get(col, "")) else ""
        vals = [
            gv("project_name"), gv("project_id"), gv("project_manager"), gv("ps_region"),
            gv("phase"), gv("weighted_score"), gv("risk_level"), gv("schedule_health"),
            gv("rag"), gv("client_sentiment"), gv("client_responsiveness"),
            gv("go_live_date").strftime("%Y-%m-%d") if pd.notna(gv("go_live_date")) and hasattr(gv("go_live_date"), "strftime") else gv("go_live_date"),
        ]
        for col, val in enumerate(vals, 1):
            c = ws_risk.cell(r, col, val)
            bg = rag_color(gv("rag")) if col == 9 else "FEF0EE"
            style_cell(c, bg, align="center" if col > 2 else "left")
            c.border = border_thin()
        ws_risk.row_dimensions[r].height = 16
    ws_risk.freeze_panes = "A4"

    # ── Tab 5: Phase Distribution Matrix ─────────────────────────────────────
    ws_phase = wb.create_sheet("Phase Distribution")
    ws_phase.sheet_properties.tabColor = "1ABC9C"

    phase_order = [
        "00. Onboarding",
        "01. Requirements and Design",
        "02. Configuration",
        "03. Enablement/Training",
        "04. UAT",
        "05. Prep for Go-Live",
        "06. Go-Live",
        "07. Data Migration",
        "08. Ready for Support Transition",
        "09. Phase 2 Scoping",
        "10. Complete/Pending Final Billing",
        "11. On Hold",
        "12. PS Review",
    ]

    active_only = scored_df[scored_df["active"]].copy()
    if "phase" in active_only.columns and "project_manager" in active_only.columns:
        pivot = active_only.groupby(["project_manager", "phase"]).size().unstack(fill_value=0)
        existing_phases = [p for p in phase_order if p in pivot.columns]
        other_phases    = [p for p in pivot.columns if p not in phase_order]
        pivot = pivot[existing_phases + other_phases]
        # Total projects per consultant (all FF excl complete)
        if "total_project" in scored_df.columns:
            total_only = scored_df[scored_df["total_project"]]
            total_proj_counts = total_only.groupby("project_manager")["project_id"].nunique().rename("Total Projects")
        else:
            total_proj_counts = active_only.groupby("project_manager")["project_id"].nunique().rename("Total Projects")

        active_proj_counts = active_only.groupby("project_manager")["project_id"].nunique().rename("Active Projects")

        pivot["Total Projects"]  = pivot.index.map(total_proj_counts).fillna(0).astype(int)
        pivot["Active Projects"] = pivot.index.map(active_proj_counts).fillna(0).astype(int)
        pivot["Total Score"]     = active_only.groupby("project_manager")["weighted_score"].sum().round(1)
        pivot["Workload Level"]  = pivot["Total Score"].apply(lambda s: workload_level(s)[0])
        pivot = pivot.sort_values("Total Score", ascending=False)
        pivot = pivot.reset_index()

        phase_headers = ["Consultant"] + list(pivot.columns[1:])
        write_title(ws_phase, "PHASE DISTRIBUTION MATRIX — Active FF Projects by Consultant",
                    len(phase_headers))
        style_header(ws_phase, 2, phase_headers, "1ABC9C")
        ws_phase.auto_filter.ref = f"A2:{get_column_letter(len(phase_headers))}2"
        ws_phase.column_dimensions["A"].width = 28
        for i in range(2, len(phase_headers) + 1):
            ws_phase.column_dimensions[get_column_letter(i)].width = 14

        for r_idx, row_data in pivot.iterrows():
            r = 3 + r_idx
            level = row_data.get("Workload Level", "")
            bg    = level_color(level)
            for col, val in enumerate(row_data.tolist(), 1):
                c = ws_phase.cell(r, col, val)
                style_cell(c, bg, align="center" if col > 1 else "left")
                c.border = border_thin()
            ws_phase.row_dimensions[r].height = 16
        ws_phase.freeze_panes = "B3"

    # ── Tab 6: Processed Data ─────────────────────────────────────────────────
    ws_raw = wb.create_sheet("Processed Data")
    raw_cols = [c for c in scored_df.columns]
    style_header(ws_raw, 1, raw_cols, NAVY)
    for r_idx, (_, row_data) in enumerate(scored_df.iterrows()):
        r = 2 + r_idx
        for col, val in enumerate(row_data.tolist(), 1):
            c = ws_raw.cell(r, col, str(val) if pd.notna(val) else "")
            c.font      = Font(size=9, name="Calibri")
            c.alignment = Alignment(horizontal="left")
        ws_raw.row_dimensions[r].height = 14

    # Tab order
    tab_order = ["Dashboard", "By Consultant", "By Project",
                 "At-Risk Projects", "Phase Distribution", "Processed Data"]
    wb._sheets = sorted(wb._sheets, key=lambda s: tab_order.index(s.title) if s.title in tab_order else 99)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Streamlit UI ──────────────────────────────────────────────────────────────
def main():
    st.markdown("""
        <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
        <style>
            html, body, [class*="css"] { font-family: 'Manrope', sans-serif !important; }
            h1, h2, h3, .stMarkdown, .stDataFrame, label, button { font-family: 'Manrope', sans-serif !important; }
        </style>
        <div style='background-color:#1e2c63;padding:24px 32px;border-radius:8px;margin-bottom:24px;font-family:Manrope,sans-serif'>
            <h1 style='color:white;margin:0;font-size:28px;font-family:Manrope,sans-serif'>Workload Health Score</h1>
            <p style='color:#aac4d0;margin:6px 0 0 0;font-size:14px;font-family:Manrope,sans-serif'>
                Weighted project scoring across active Fixed Fee engagements — by consultant and phase
            </p>
            <p style='color:#8ab0c0;margin:8px 0 0 0;font-size:12px;font-family:Manrope,sans-serif;line-height:1.6;'>
                Each active FF project is scored by <b>phase weight</b> × <b>client health multiplier</b> × <b>risk multiplier</b>.
                Thresholds: <b>Low</b> 1–25 pts &nbsp;·&nbsp; <b>Medium</b> 26–60 pts &nbsp;·&nbsp; <b>High</b> 61+ pts (flag to Director).
                T&amp;M projects and inactive phases (On Hold / Complete) are excluded from scoring.
            </p>
        </div>
    """, unsafe_allow_html=True)

    # Phase weight reference
    with st.expander("📋 View phase weights & scoring model"):
        weight_df = pd.DataFrame([
            {"Phase": k.title(), "Weight (pts)": v}
            for k, v in PHASE_WEIGHTS.items()
        ])
        st.dataframe(weight_df, hide_index=True, use_container_width=True)

    # ── Uploads ───────────────────────────────────────────────────────────────
    st.subheader("Step 1 — Upload Smartsheets DRS Export")
    st.caption("Required columns: Project ID, Project Name, Project Phase, Project Type, Territory, Status")
    ss_file = st.file_uploader(
        "Drop your file here or click to browse",
        type=["xlsx", "xls", "csv"],
        key="ss_upload"
    )

    st.subheader("Step 2 — Upload NetSuite Time Detail Export *(optional — provides PM assignment)*")
    st.caption("Used to join Project Manager to each project via Project ID")
    ns_file = st.file_uploader(
        "Drop your file here or click to browse",
        type=["xlsx", "xls", "csv"],
        key="ns_upload"
    )

    if not ss_file:
        st.info("👆 Upload your Smartsheets DRS export to continue.")
        return

    # ── Process ───────────────────────────────────────────────────────────────
    try:
        ss_df, milestone_cols = load_ss(ss_file)
        ns_df = load_ns(ns_file) if ns_file else None
        scored_df = score_projects(ss_df, ns_df)
        consultant_df, missing_pm = build_consultant_summary(scored_df)
        as_of = datetime.today().strftime("%B %d, %Y")
    except Exception as e:
        st.error(f"Error processing files: {e}")
        st.exception(e)
        return

    st.success("✅ Processing complete!")

    # ── Metrics ───────────────────────────────────────────────────────────────
    total    = len(consultant_df)
    high     = len(consultant_df[consultant_df["workload_level"] == "High"])
    medium   = len(consultant_df[consultant_df["workload_level"] == "Medium"])
    low      = len(consultant_df[consultant_df["workload_level"] == "Low"])
    active_projects = scored_df[scored_df["active"]]["project_id"].nunique()
    total_projects  = scored_df[scored_df["total_project"]]["project_id"].nunique() if "total_project" in scored_df.columns else active_projects

    st.markdown(f"<div style='font-size:13px;color:#a0a0a0;font-family:Manrope,sans-serif;margin-bottom:12px'>Data as of <strong style='color:#ffffff'>{as_of}</strong></div>", unsafe_allow_html=True)

    def metric_card(label, value, pill_txt=None, pill_fg=None):
        pill = ""
        if pill_txt and pill_fg:
            pill = f"<div style='display:inline-block;margin-top:6px;padding:2px 10px;border-radius:999px;background-color:{pill_fg}33;font-size:13px;font-family:Manrope,sans-serif;color:{pill_fg}'>&#8593; {pill_txt}</div>"
        return f"<div style='font-size:14px;color:#a0a0a0;font-family:Manrope,sans-serif;margin-bottom:4px'>{label}</div><div style='font-size:36px;font-weight:700;color:inherit;font-family:Manrope,sans-serif;line-height:1.1'>{value}</div>{pill}"

    m1, m2, m3, m4, m5 = st.columns(5)
    with m1: st.markdown(metric_card("Consultants Scored",  f"{total:,}"), unsafe_allow_html=True)
    with m2: st.markdown(metric_card("Total FF Projects",   f"{total_projects:,}"), unsafe_allow_html=True)
    with m3: st.markdown(metric_card("Active Projects",     f"{active_projects:,}", "Excl. On Hold", "#4472C4"), unsafe_allow_html=True)
    with m4: st.markdown(metric_card("High Workload",       f"{high}",  "At or over capacity", "#e74c3c"), unsafe_allow_html=True)
    with m5: st.markdown(metric_card("Medium Workload",     f"{medium}", "Monitor for changes", "#f39c12"), unsafe_allow_html=True)

    if missing_pm > 0:
        st.warning(f"⚠️ {missing_pm} project(s) could not be assigned a PM — Project ID not found in NetSuite export. Update NS to include these in consultant scoring.")

    st.markdown("---")

    # ── Preview tabs ──────────────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs(["By Consultant", "By Project", "At-Risk"])

    with tab1:
        display_con = consultant_df.rename(columns={
            "project_manager": "Consultant",
            "ps_region":       "PS Region",
            "project_count":   "Active Projects",
            "total_score":     "Weighted Score",
            "workload_level":  "Workload Level",
        })
        st.dataframe(display_con, hide_index=True, use_container_width=True)

    with tab2:
        proj_display_cols = ["project_name", "project_id", "project_type", "phase",
                             "project_manager", "territory", "weighted_score",
                             "schedule_health", "risk_level", "rag", "active"]
        avail = [c for c in proj_display_cols if c in scored_df.columns]
        st.dataframe(
            scored_df[avail].sort_values("weighted_score", ascending=False),
            hide_index=True, use_container_width=True
        )

    with tab3:
        at_risk = scored_df[
            (scored_df["active"]) & (
                (scored_df.get("risk_level",      pd.Series(dtype=str)).str.lower().str.contains("high",   na=False)) |
                (scored_df.get("schedule_health", pd.Series(dtype=str)).str.lower().str.contains("behind", na=False)) |
                (scored_df.get("rag",             pd.Series(dtype=str)).str.lower().str.contains("red",    na=False))
            )
        ]
        if len(at_risk) == 0:
            st.success("No at-risk projects flagged.")
        else:
            risk_cols = ["project_name", "project_manager", "phase", "weighted_score",
                         "risk_level", "schedule_health", "rag", "go_live_date"]
            avail_r = [c for c in risk_cols if c in at_risk.columns]
            st.dataframe(at_risk[avail_r].sort_values("weighted_score", ascending=False),
                         hide_index=True, use_container_width=True)

    # ── Download ──────────────────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Step 3 — Generate Report")
    excel_buf = build_excel(scored_df, consultant_df, missing_pm, as_of)
    fname = f"Workload_Health_Score_{datetime.today().strftime('%Y%m%d')}.xlsx"
    st.download_button(
        label="⬇️  Download Workload Health Score Report",
        data=excel_buf,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()