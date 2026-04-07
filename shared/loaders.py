"""
PS Tools — Shared Loaders & Calculation Helpers
Data loading, fuzzy matching, inactivity calc, template utilities.
"""
import re
import pandas as pd
import streamlit as st
from datetime import date, datetime
from rapidfuzz import fuzz

from shared.constants import PRODUCT_KEYWORDS, MILESTONE_COLS_MAP, PHASE_BENCHMARKS, SS_COL_MAP, SFDC_COL_MAP
from shared.config import DEFAULT_SCOPE

# Import the authoritative full column map and inactive phase set from template_utils.
# These are richer than the short SS_COL_MAP in constants — they include milestone cols,
# last_updated aliases, account columns, etc.
from shared.template_utils import SS_COL_MAP_OUT, INACTIVE_PHASES_OUT


def _extract_product_hints(text):
    """Return set of product keywords found in a string."""
    t = str(text).lower()
    return {kw for kw in PRODUCT_KEYWORDS if kw.lower() in t}


def _clean_account(text):
    """Normalise account name for fuzzy comparison — strip legal suffixes, punctuation."""
    t = str(text).lower()
    for stop in ["ltd","limited","inc","llc","plc","gmbh","the ","abf ","- za -","& co","co."]:
        t = t.replace(stop, " ")
    return re.sub(r"[^a-z0-9 ]", " ", t).split()


def fuzzy_match_sfdc(df_sfdc, project_name, account_name):
    """
    Find best SFDC row(s) for a DRS project using:
    1. Exact opportunity name match
    2. Fuzzy account name + product keyword overlap
    Returns matching rows and a confidence label.
    """
    if df_sfdc is None or df_sfdc.empty:
        return pd.DataFrame(), None

    # ── Exact opp name match ───────────────────────────────────────────────
    if "opportunity" in df_sfdc.columns:
        exact = df_sfdc[df_sfdc["opportunity"].astype(str).str.strip().str.lower()
                        == str(project_name).strip().lower()]
        if not exact.empty:
            return exact, "Exact match"

    # ── Fuzzy account + product keyword match ─────────────────────────────
    drs_words    = set(_clean_account(account_name))
    drs_products = _extract_product_hints(project_name)

    best_score  = 0
    best_rows   = pd.DataFrame()
    best_label  = None
    best_opp_id = None
    best_opp_nm = None

    for _, row in df_sfdc.iterrows():
        sfdc_account  = str(row.get("account",  ""))
        sfdc_opp      = str(row.get("opportunity", ""))

        # Account fuzzy score — word overlap + token set ratio for partial name matches
        sfdc_words    = set(_clean_account(sfdc_account))
        common        = drs_words & sfdc_words
        word_score    = len(common) / max(len(drs_words), 1) * 100
        fuzz_score    = fuzz.token_set_ratio(" ".join(drs_words), " ".join(sfdc_words))
        acct_score    = max(word_score, fuzz_score * 0.7)  # blend both signals

        # Product keyword overlap
        sfdc_products = _extract_product_hints(sfdc_opp)
        prod_match    = bool(drs_products & sfdc_products)

        # Combined score — account similarity weighted, product match bonus
        score = acct_score + (30 if prod_match else 0)

        if score > best_score:
            best_score   = score
            best_opp_id  = row.get("opportunity_id", None)
            best_opp_nm  = row.get("opportunity",    None)
            best_label   = (
                f"Fuzzy match ({int(acct_score)}% account · {'✅ product match' if prod_match else '⚠️ no product match'})"
            )

    # Only return if confident enough
    if best_score >= 60:
        # Return ALL contacts for the matched opportunity (not just the best-scoring row)
        if best_opp_id and "opportunity_id" in df_sfdc.columns:
            best_rows = df_sfdc[df_sfdc["opportunity_id"] == best_opp_id]
        elif best_opp_nm and "opportunity" in df_sfdc.columns:
            best_rows = df_sfdc[df_sfdc["opportunity"] == best_opp_nm]
        if not best_rows.empty:
            return best_rows, best_label

    # ── Last resort: account name only, high threshold ────────────────────
    if "account" in df_sfdc.columns and account_name:
        df_sfdc["_acct_score"] = df_sfdc["account"].apply(
            lambda x: fuzz.token_set_ratio(str(account_name).lower(), str(x).lower())
        )
        top = df_sfdc[df_sfdc["_acct_score"] >= 75].sort_values("_acct_score", ascending=False)
        df_sfdc.drop(columns=["_acct_score"], inplace=True)
        if not top.empty:
            return top, f"Account name match ({top.iloc[0].get('_acct_score', '?')}% similarity)"

    return pd.DataFrame(), None


# ── Active employee list (sourced from EMPLOYEE_ROLES — leavers excluded) ────
ACTIVE_EMPLOYEES = [
    "Arestarkhov, Yaroslav", "Barrio, Nairobi", "Bell, Stuart", "Cadelina",
    "Carpen, Anamaria", "Centinaje, Rhodechild", "Church, Jason G", "Cooke, Ellen",
    "Cruz, Daniel", "DiMarco, Nicole R", "Dolha, Madalina", "Dunn, Steven",
    "Finalle-Newton, Jesse", "Gardner, Cheryll L", "Hopkins, Chris", "Hughes, Madalyn",
    "Ickler, Georganne", "Isberg, Eric", "Jordanova, Marija", "Lappin, Thomas",
    "Law, Brandon", "Longalong, Santiago", "Mohammad, Manaan", "Morris, Lisa",
    "Murphy, Conor", "NAQVI, SYED", "Olson, Austin D", "Pallone, Daniel",
    "Porangada, Suraj", "Quiambao, Generalyn", "Raykova, Silvia",
    "Selvakumar, Sajithan", "Snee, Stefanie J", "Swanson, Patti",
    "Tuazon, Carol", "Zoric, Ivan",
]


def load_sfdc(file):
    if file.name.endswith(".csv"):
        for enc in ["utf-8", "utf-8-sig", "windows-1252", "latin-1", "cp1252"]:
            try:
                file.seek(0)
                df = pd.read_csv(file, encoding=enc)
                break
            except (UnicodeDecodeError, LookupError):
                continue
        else:
            file.seek(0)
            df = pd.read_csv(file, encoding="latin-1")  # last resort
    else:
        df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    rename = {c: SFDC_COL_MAP[c.lower()] for c in df.columns if c.lower() in SFDC_COL_MAP}
    df = df.rename(columns=rename)
    if "close_date" in df.columns:
        df["close_date"] = pd.to_datetime(df["close_date"], errors="coerce")
    # Always build contact_name from first + last (source report has separate columns)
    if "first_name" in df.columns and "last_name" in df.columns:
        df["contact_name"] = (
            df["first_name"].fillna("").astype(str) + " " +
            df["last_name"].fillna("").astype(str)
        ).str.strip()
    # Use opportunity name as the project label
    if "opportunity" not in df.columns and "opportunity_id" in df.columns:
        df["opportunity"] = df["opportunity_id"]
    # Normalise impl_contact_flag — checkbox may come as TRUE/FALSE, Yes/No, 1/0
    if "impl_contact_flag" in df.columns:
        df["impl_contact_flag"] = df["impl_contact_flag"].astype(str).str.strip().str.lower().isin(
            ["true", "yes", "1", "checked", "x"]
        )
    # Normalise is_primary flag — store as readable string to avoid date rendering
    if "is_primary" in df.columns:
        _prim_bool = df["is_primary"].astype(str).str.strip().str.lower().isin(
            ["true", "yes", "1", "checked", "x"]
        )
        df["is_primary"] = _prim_bool.map({True: "✓", False: ""})
    # Normalise contact_roles to lowercase string for matching
    if "contact_roles" in df.columns:
        df["contact_roles"] = df["contact_roles"].fillna("").astype(str).str.strip()
    return df


def load_drs(file):
    """Load SS DRS export and identify stale/unresponsive projects."""
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    rename = {col: SS_COL_MAP_OUT[col.lower()] for col in df.columns if col.lower() in SS_COL_MAP_OUT}
    df = df.rename(columns=rename)

    if "start_date" in df.columns:
        df["start_date"] = pd.to_datetime(df["start_date"], errors="coerce")
    if "go_live_date" in df.columns:
        df["go_live_date"] = pd.to_datetime(df["go_live_date"], errors="coerce")
    if "last_updated" in df.columns:
        df["last_updated"] = pd.to_datetime(df["last_updated"], errors="coerce")
    # Parse all milestone date columns — clip epoch/nonsense dates (e.g. checkbox=1 → 1970-01-01)
    _min_valid = pd.Timestamp("2015-01-01")
    for _ms_col in MILESTONE_COLS_MAP.keys():
        if _ms_col in df.columns:
            _parsed = pd.to_datetime(df[_ms_col], errors="coerce")
            df[_ms_col] = _parsed.where(_parsed >= _min_valid, pd.NaT)

    # Filter to active FF projects only
    today = pd.Timestamp.today().normalize()
    if "phase" in df.columns:
        df = df[~df["phase"].str.strip().str.lower().isin(INACTIVE_PHASES_OUT)]
    if "billing_type" in df.columns:
        df = df[~df["billing_type"].str.strip().str.lower().isin({"t&m","time & material","time and material"})]
    # Calculate remaining sessions from milestone columns
    _session_ms_cols = ["ms_enablement", "ms_session1", "ms_session2"]
    _present_ms = [col for col in _session_ms_cols if col in df.columns]
    if _present_ms:
        def _count_remaining(row):
            return sum(1 for col in _present_ms if pd.isna(row[col]) or str(row[col]).strip() in ("", "nan", "None", "NaT"))
        df["remaining_sessions"] = df.apply(_count_remaining, axis=1).astype(int)
    else:
        df["remaining_sessions"] = None

    # Calculate last milestone (name + date of most recently completed milestone)
    df["last_milestone"] = df.apply(calc_last_milestone, axis=1)

    # Tag On Hold but keep in df — excluded from dropdown, shown in table
    ON_HOLD_VALS = {"on-hold","on hold","onhold","on_hold"}
    if "status" in df.columns:
        df["_on_hold"] = df["status"].str.strip().str.lower().isin(ON_HOLD_VALS)
    else:
        df["_on_hold"] = False

    # ── Inactivity signal hierarchy (DRS-only baseline, NS overrides in calc_days_inactive) ──
    # 1. Last Milestone date — days since most recently completed milestone
    # 2. Phase vs Start Date benchmark — days open minus expected phase duration
    # 3. Unknown (-1) — NS Time Entry will override when uploaded
    _PHASE_BENCH = {
        "00. onboarding": 7, "01. requirements and design": 14,
        "02. configuration": 21, "03. enablement/training": 14,
        "04. uat": 28, "05. prep for go-live": 7,
        "06. go-live (hypercare)": 14, "08. ready for support transition": 5,
        "09. phase 2 scoping": 14,
    }

    def _calc_inactivity_drs(row):
        # 1. Last milestone
        for _ms_col in reversed(list(MILESTONE_COLS_MAP.keys())):
            if _ms_col in row.index and pd.notna(row[_ms_col]):
                try:
                    return int((today - pd.Timestamp(row[_ms_col])).days), "Last Milestone"
                except: pass
        # 2. Phase vs start date
        if "start_date" in row.index and pd.notna(row["start_date"]):
            try:
                days_open    = int((today - pd.Timestamp(row["start_date"])).days)
                bench        = _PHASE_BENCH.get(str(row.get("phase","")).strip().lower(), 14)
                days_stalled = max(0, days_open - bench)
                return days_stalled, "Phase vs Start Date"
            except: pass
        return -1, "Unknown"

    _results = df.apply(_calc_inactivity_drs, axis=1)
    df["days_inactive"]      = _results.apply(lambda x: x[0]).astype(int)
    df["_inactivity_source"] = _results.apply(lambda x: x[1])

    # Normalise legacy flag
    if "legacy" in df.columns:
        df["legacy"] = df["legacy"].fillna("").astype(str).str.strip().str.lower().isin(
            {"yes", "true", "1", "y"}
        )
    else:
        df["legacy"] = False

    # Burn % = actual_hours / (budgeted_hours + change_order)
    if "actual_hours" in df.columns and "budgeted_hours" in df.columns:
        _co = pd.to_numeric(df.get("change_order", 0), errors="coerce").fillna(0)
        _bgt = pd.to_numeric(df["budgeted_hours"], errors="coerce").fillna(0)
        _act = pd.to_numeric(df["actual_hours"], errors="coerce").fillna(0)
        _scope = _bgt + _co
        df["burn_pct"] = (_act / _scope * 100).where(_scope > 0, None).round(1)
    else:
        df["burn_pct"] = None

    # Store unmapped columns for debug
    df.attrs["unmapped_cols"] = [c for c in df.columns if c not in SS_COL_MAP_OUT.values()]

    return df


def load_ns_time(file):
    """Load NS Time Detail — derive last time entry per project for inactivity."""
    if file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)
    df.columns = df.columns.str.strip()
    rename = {col: NS_COL_MAP_OUT[col.strip().lower()] for col in df.columns if col.strip().lower() in NS_COL_MAP_OUT}
    df = df.rename(columns=rename)
    if "date" in df.columns:
        # Handle Unix timestamps from NS exports (seconds or ms) as well as string dates
        # Skip if already datetime
        if not pd.api.types.is_datetime64_any_dtype(df["date"]):
            _date_raw = df["date"]
            try:
                _parsed = pd.to_numeric(_date_raw, errors="coerce")
                _is_ms  = _parsed.notna() & (_parsed > 1e11)
                _is_sec = _parsed.notna() & (_parsed > 1e8) & ~_is_ms
                if _is_ms.any() or _is_sec.any():
                    _dt = pd.Series(pd.NaT, index=df.index)
                    if _is_ms.any():
                        _dt[_is_ms] = pd.to_datetime(_parsed[_is_ms], unit="ms", errors="coerce")
                    if _is_sec.any():
                        _dt[_is_sec] = pd.to_datetime(_parsed[_is_sec], unit="s", errors="coerce")
                    _remaining = ~(_is_ms | _is_sec)
                    if _remaining.any():
                        _dt[_remaining] = pd.to_datetime(_date_raw[_remaining], errors="coerce")
                    df["date"] = _dt
                else:
                    df["date"] = pd.to_datetime(_date_raw, errors="coerce")
            except Exception:
                df["date"] = pd.to_datetime(_date_raw, errors="coerce")
    if "project_id" in df.columns:
        df["project_id"] = df["project_id"].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    if "ns_rate" in df.columns:
        df["ns_rate"] = pd.to_numeric(df["ns_rate"], errors="coerce").fillna(0)
    # Store original cols for debug
    df.attrs["original_cols"] = list(df.columns)
    return df


def calc_days_inactive(df_drs, df_ns):
    """Merge NS last time entry onto DRS projects to get objective inactivity days."""
    today = pd.Timestamp.today().normalize()
    if df_ns is None or "date" not in df_ns.columns:
        return df_drs

    def _clean_id(series):
        """Safely convert project_id series to clean integer string."""
        def _to_id(v):
            try:
                s = str(v).strip()
                if s in ("", "nan", "None", "NaN"): return ""
                # Handle float like 367986.0 → 367986
                return str(int(float(s)))
            except: return str(v).strip()
        return series.apply(_to_id)

    # Try join on project_id first, fall back to project name
    if "project_id" in df_ns.columns and "project_id" in df_drs.columns:
        df_ns  = df_ns.copy()
        df_drs = df_drs.copy()
        df_ns["project_id"]  = _clean_id(df_ns["project_id"])
        df_drs["project_id"] = _clean_id(df_drs["project_id"])
        last_entry = (
            df_ns[df_ns["date"].notna()]
            .groupby("project_id")["date"].max()
            .reset_index()
            .rename(columns={"date": "last_ns_entry"})
        )
        df_drs = df_drs.merge(last_entry, on="project_id", how="left")
    elif "project" in df_ns.columns and "project_name" in df_drs.columns:
        last_entry = (
            df_ns[df_ns["date"].notna()]
            .groupby("project")["date"].max()
            .reset_index()
            .rename(columns={"project": "project_name", "date": "last_ns_entry"})
        )
        df_drs = df_drs.merge(last_entry, on="project_name", how="left")

    if "last_ns_entry" in df_drs.columns:
        # Clip epoch dates (1970-01-01) — these indicate a parse failure, not real dates
        _min_valid_date = pd.Timestamp("2015-01-01")
        df_drs["last_ns_entry"] = df_drs["last_ns_entry"].where(
            df_drs["last_ns_entry"] >= _min_valid_date, pd.NaT
        )
        ns_days = (today - df_drs["last_ns_entry"]).dt.days.clip(lower=0)
        fallback = df_drs["days_inactive"] if "days_inactive" in df_drs.columns else pd.Series(-1, index=df_drs.index)
        df_drs["days_inactive"] = ns_days.where(df_drs["last_ns_entry"].notna(), fallback).fillna(-1).astype(int)
        _existing_src = df_drs["_inactivity_source"].tolist() if "_inactivity_source" in df_drs.columns else ["Unknown"] * len(df_drs)
        df_drs["_inactivity_source"] = [
            "NS Time Entry" if pd.notna(x) else _existing_src[i]
            for i, x in enumerate(df_drs["last_ns_entry"])
        ]
    return df_drs


# ── Outreach Log — session_state + /tmp file for cross-session persistence ────
import os, json as _json
LOG_PATH = "/tmp/ps_outreach_log.json"
_LOG_KEY  = "_outreach_log"


def calc_last_milestone(row):
    """Return 'Milestone Name · YYYY-MM-DD' for the latest completed milestone."""
    best_date = None
    best_name = None
    for col, label in MILESTONE_COLS_MAP.items():
        if col in row.index:
            val = row[col]
            try:
                dt = pd.to_datetime(val, errors="coerce")
                if pd.notna(dt):
                    if best_date is None or dt > best_date:
                        best_date = dt
                        best_name = label
            except: pass
    if best_name and best_date:
        return f"{best_name} · {best_date.strftime('%Y-%m-%d')}"
    return "—"


def normalise_product_name(raw):
    """Convert internal DRS project type labels to customer-facing product names."""
    if not raw or str(raw).strip().lower() in ("nan", "none", ""):
        return ""
    s = str(raw).strip()

    # Explicit mappings — checked first (case-insensitive key match)
    EXPLICIT = {
        "zoneapp: cc":           "CC Import",
        "zoneapp: psp":          "PSP",
        "zoneapp: sftp":         "SFTP",
        "zoneapp: e-invoicing":  "e-Invoicing",
        "zoneapp: einvoicing":   "e-Invoicing",
        "zoneapp: approvals":    "ZoneApprovals",
        "zoneapp: capture":      "ZoneCapture",
        "zoneapp: reconcile":    "ZoneReconcile",
        "zoneapp: payments":     "ZonePayments",
        "zoneapp: reconcile 2.0":"ZoneReconcile",
        "zoneapp: premium":      "Premium",
        "zoneapp: capture and e-invoicing": "ZoneCapture + e-Invoicing",
    }
    if s.lower() in EXPLICIT:
        return EXPLICIT[s.lower()]

    # Generic fallback: strip "ZoneApp: " prefix and add Zone
    import re as _re
    s = _re.sub(r"(?i)^zoneapp:\s*", "Zone", s)
    return s


# Milestone columns in delivery order — mapped internal name : display name
MILESTONE_COLS_MAP = {
    "ms_intro_email":    "Intro. Email Sent",
    "ms_config_start":   "Standard Config Start",
    "ms_enablement":     "Enablement Session",
    "ms_session1":       "Session #1",
    "ms_session2":       "Session #2",
    "ms_uat_signoff":    "UAT Signoff",
    "ms_prod_cutover":   "Prod Cutover",
    "ms_hypercare_start":"Hypercare Start",
    "ms_close_out":      "Close Out Remaining Tasks",
    "ms_transition":     "Transition to Support",
}


def suggest_tier_from_days(days):
    try:
        d = int(days)
        if d < 0:   return None   # unknown inactivity
        if d < 30:  return None   # follow-up flag only, no template
        for name, tmpl in TEMPLATES.items():
            if tmpl["days_min"] <= d <= tmpl["days_max"]:
                return name
    except:
        pass
    return None


NS_COL_MAP_OUT = {
    # ── Employee ──────────────────────────────────────────────────────────────
    "employee":             "employee",
    "name":                 "employee",
    "time by employee":     "employee",       # Finance report header
    # ── Project ──────────────────────────────────────────────────────────────
    "project":              "project",
    "project name":         "project",
    "project id":           "project_id",
    "project id ":          "project_id",
    " project id":          "project_id",
    "project_id":           "project_id",
    "projectid":            "project_id",
    "project internal id":  "project_id",
    "job":                  "project_id",
    # ── Billing ──────────────────────────────────────────────────────────────
    "billing type":         "billing_type",
    "project type":         "project_type",
    "type":                 "project_type",
    "non-billable":         "non_billable",
    "non billable":         "non_billable",
    "nonbillable":          "non_billable",
    "non_billable":         "non_billable",
    "is non billable":      "non_billable",
    # ── Time entry ───────────────────────────────────────────────────────────
    "date":                 "date",
    "transaction date":     "date",
    "time entry date":      "date",           # Finance report header
    "hours":                "hours",
    "quantity":             "hours",
    "hours/quantity":       "hours",
    "time entry duration":  "hours",          # Finance report header
    "hours to date":        "hours_to_date",
    "quantity to date":     "hours_to_date",
    "hours/quantity to date": "hours_to_date",
    # ── Rate / revenue ───────────────────────────────────────────────────────
    "rate":                 "ns_rate",
    "billing rate":         "ns_rate",
    "hourly rate":          "ns_rate",
    "time detail rate":     "ns_rate",
    "time entry rate":      "ns_rate",        # Finance report header
    "blended rate":         "blended_rate",   # Finance report — keep separate from SOW rate
    "amount billed":        "amount_billed",  # Finance report — actual billed amount
    "billing rate (foreign currency)": "ns_rate",
    "rate (foreign currency)":         "ns_rate",
    # ── Project manager / meta ───────────────────────────────────────────────
    "project manager":      "project_manager",
    "territory":            "territory",
    "project phase":        "phase",
    "customer":             "account",        # Finance report account name
    "time entry status":    "entry_status",   # Billed / Unbilled
    "time entry memo":      "memo",
    # ── NS internal row ID (time entry ID — for 1:1 reconciliation) ──────────
    "internal id":          "internal_id",    # NS time entry internal ID
    "id":                   "internal_id",    # short form
    # ── Currency ──────────────────────────────────────────────────────────────
    "currency":             "currency",       # e.g. USD, GBP, EUR, AUD
    "cur":                  "currency",       # NS abbreviation
    "foreign currency":     "currency",
    "cost rate":            "ns_rate",
    "pay rate":             "ns_rate",
}


# ── Product family mapping (SFDC → canonical) ────────────────────────────────
_PRODUCT_FAMILY_MAP = [
    # Specific Zone-prefixed first (longer match wins)
    ("zoneapprovals",  "Approvals"),
    ("zonecapture",    "Capture"),
    ("zonereporting",  "Reporting"),
    ("zonebilling",    "Billing"),
    ("zonereconcile",  "Reconcile"),
    ("zonepayroll",    "Payroll"),
    ("zonepayments",   "Payments"),
    ("zonepsp",        "PSP"),
    ("zonesftp",       "SFTP"),
    # Without prefix
    ("e-invoicing",    "E-Invoicing"),
    ("einvoicing",     "E-Invoicing"),
    ("approvals",      "Approvals"),
    ("capture",        "Capture"),
    ("reporting",      "Reporting"),
    ("billing",        "Billing"),
    ("reconcile",      "Reconcile"),
    ("payroll",        "Payroll"),
    ("payments",       "Payments"),
    ("psp",            "PSP"),
    ("sftp",           "SFTP"),
]

def match_product_family(product_family: str, products_text: str = "") -> str:
    """Map SFDC Product Family → canonical product name.
    Falls back to Products free-text if Product Family is blank or 'Other'.
    """
    pf = str(product_family or "").strip().lower().replace(" ", "")
    if pf and pf != "other":
        for key, val in _PRODUCT_FAMILY_MAP:
            if key in pf:
                return val

    # Fall back to Products free-text
    pt = str(products_text or "").strip().lower()
    for key, val in _PRODUCT_FAMILY_MAP:
        if key in pt:
            return val
    return "Other"


# ── T&M SOW loader ────────────────────────────────────────────────────────────
TM_SOW_COL_MAP = {
    "opportunity owner":                    "opportunity_owner",
    "account name":                         "account_name",
    "opportunity name":                     "opportunity_name",
    "stage":                                "stage",
    "fiscal period":                        "fiscal_period",
    "probability (%)":                      "probability",
    "close date":                           "close_date",
    "ps sow hours":                         "sow_hours",
    "ps sow amount (converted) currency":   "sow_amount_currency",
    "ps sow amount (converted)":            "sow_amount_usd",
    "ps sow rate (converted) currency":     "sow_rate_converted_currency",
    "ps sow rate (converted)":              "sow_rate_usd",
    "ps sow rate currency":                 "sow_rate_local_currency",
    "ps sow rate":                          "sow_rate_local",
    "products":                             "products_text",
    "product family":                       "product_family",
    "region":                               "region",
    "zzz_arr (converted)":                  "arr_usd",
    "created date":                         "created_date",
}

def calc_tm_monthly_actuals(df_ns: pd.DataFrame, df_sow: pd.DataFrame) -> pd.DataFrame:
    """Build monthly T&M actual revenue from NS time entries × matched SOW rates.

    Strategy:
    1. Run join_tm_to_ns to get the best SOW rate per NS project (reuses existing match logic)
    2. Build {ns_project_lower: rate_usd} lookup from matched results
    3. Apply to NS T&M rows grouped by project + month

    Region derived from project_manager → EMPLOYEE_LOCATION → PS_REGION_MAP
    Product derived from project_type keyword match
    """
    from shared.config import EMPLOYEE_LOCATION, PS_REGION_MAP, PS_REGION_OVERRIDE

    if df_ns is None or df_ns.empty:
        return pd.DataFrame()

    # ── T&M row selection ────────────────────────────────────────────────────
    # Primary: explicit T&M billing type
    # Secondary: Fixed Fee projects with BILLABLE time entries — these count as T&M
    #            revenue (e.g. change order overages) but are flagged for Finance review
    _bt = df_ns.get("billing_type", pd.Series("", index=df_ns.index)).fillna("").str.lower()
    _nb_raw = df_ns.get("non_billable", pd.Series("", index=df_ns.index)).fillna("").astype(str).str.strip().str.lower()
    _is_nb_all = _nb_raw.isin(["true","t","yes","1","y"])

    _is_tm = _bt.str.contains("t&m|time", na=False)
    _is_ff_billable = (
        _bt.str.contains("fixed fee|fixed.fee|fixed_fee|\bff\b", na=False, regex=True)
        & ~_is_nb_all
    )

    tm = df_ns[_is_tm | _is_ff_billable].copy()
    if tm.empty:
        return pd.DataFrame()

    # Ensure period column
    if "period" not in tm.columns:
        if "date" not in tm.columns:
            return pd.DataFrame()
        # If already a datetime dtype, use directly
        if pd.api.types.is_datetime64_any_dtype(tm["date"]):
            tm["period"] = tm["date"].dt.strftime("%Y-%m")
        else:
            _d_raw = tm["date"]
            _d_num = pd.to_numeric(_d_raw, errors="coerce")
            _d_ms  = _d_num.notna() & (_d_num > 1e11)
            _d_sec = _d_num.notna() & (_d_num > 1e8) & ~_d_ms
            if _d_ms.any() or _d_sec.any():
                _dt = pd.Series(pd.NaT, index=tm.index)
                if _d_ms.any():
                    _dt[_d_ms] = pd.to_datetime(_d_num[_d_ms], unit="ms", errors="coerce")
                if _d_sec.any():
                    _dt[_d_sec] = pd.to_datetime(_d_num[_d_sec], unit="s", errors="coerce")
                _rem = ~(_d_ms | _d_sec)
                if _rem.any():
                    _dt[_rem] = pd.to_datetime(_d_raw[_rem], errors="coerce")
                tm["date"] = _dt
            else:
                tm["date"] = pd.to_datetime(_d_raw, errors="coerce")
            tm["period"] = tm["date"].dt.strftime("%Y-%m")

    tm["hours"] = pd.to_numeric(tm.get("hours", 0), errors="coerce").fillna(0)

    # Classify non-billable flag (recalculate on filtered set)
    if "non_billable" in tm.columns:
        tm["_is_nb"] = tm["non_billable"].fillna("").astype(str).str.strip().str.lower().isin(
            ["true","t","yes","1","y"])
    else:
        tm["_is_nb"] = False

    # Determine billing_type mask on filtered tm
    _tm_bt = tm.get("billing_type", pd.Series("", index=tm.index)).fillna("").str.lower()
    _tm_is_ff = _tm_bt.str.contains("fixed fee|fixed.fee|fixed_fee|\bff\b", na=False, regex=True)

    # Flags:
    #   ⚠️ T&M / Non-Billable  — T&M row marked non-billable: exclude from revenue, flag for review
    #   ⚠️ FF / Billable hours — FF project with billable time: counts as T&M revenue, flag for Finance
    tm["billing_flag"] = ""
    tm.loc[~_tm_is_ff &  tm["_is_nb"], "billing_flag"] = "⚠️ T&M / Non-Billable"
    tm.loc[ _tm_is_ff & ~tm["_is_nb"], "billing_flag"] = "⚠️ FF / Billable hours"

    # ── Build rate lookup from SOW match results ──────────────────────────────
    # join_tm_to_ns returns ns_project + sow_rate_usd for matched rows
    rate_by_project = {}   # {ns_project_lower: rate_usd}
    rate_by_account = {}   # {account_name_lower: rate_usd} — fallback
    if df_sow is not None and not df_sow.empty:
        # Primary: reuse join_tm_to_ns match results
        try:
            matched = join_tm_to_ns(df_sow.copy(), df_ns)
            matched["sow_rate_usd"] = pd.to_numeric(matched.get("sow_rate_usd", 0), errors="coerce").fillna(0)
            for _, r in matched.iterrows():
                proj = str(r.get("ns_project") or "").strip().lower()
                rate = float(r.get("sow_rate_usd", 0) or 0)
                # Skip nan/empty — these are unmatched SOW rows
                if proj and proj not in ("nan", "none", "") and rate > 0:
                    rate_by_project[proj] = max(rate, rate_by_project.get(proj, 0))
        except Exception:
            pass

        # Fallback: account name substring match for unmatched NS projects
        for _, sr in df_sow.iterrows():
            rate = float(sr.get("sow_rate_usd", 0) or 0)
            if rate <= 0:
                continue
            acc = str(sr.get("account_name", "") or "").strip().lower()
            if acc and len(acc) >= 5:
                rate_by_account[acc] = max(rate, rate_by_account.get(acc, 0))

    # ── Derive region from project_manager ───────────────────────────────────
    def _get_region(pm_name):
        pm = str(pm_name or "").strip()
        if pm in PS_REGION_OVERRIDE: return PS_REGION_OVERRIDE[pm]
        loc = EMPLOYEE_LOCATION.get(pm, "")
        if isinstance(loc, tuple): loc = loc[0]
        if not loc:
            for k, v in EMPLOYEE_LOCATION.items():
                if pm and (pm.lower() in k.lower() or k.lower() in pm.lower()):
                    loc = v[0] if isinstance(v, tuple) else v
                    break
        return PS_REGION_MAP.get(str(loc), "Other")

    # ── Calculate revenue per row BEFORE grouping to avoid rate aggregation errors ──
    # Using max(rate) across a group then multiplying by sum(hours) is wrong when
    # employees on the same project have different rates in the same period.
    from shared.config import get_fx_rate as _get_fx

    if "ns_rate" in tm.columns:
        tm["ns_rate"] = pd.to_numeric(tm["ns_rate"], errors="coerce").fillna(0)

    def _row_revenue(row):
        if row["_is_nb"]:
            return 0.0, 0.0, "Non-Billable"
        rate  = float(row.get("ns_rate", 0) or 0)
        if rate == 0:
            proj_l = str(row.get("project", "")).strip().lower()
            rate   = rate_by_project.get(proj_l, 0.0)
            source = "Project match" if rate > 0 else ""
            if rate == 0:
                for acc, acc_rate in rate_by_account.items():
                    if acc in proj_l:
                        rate, source = acc_rate, "Account match"
                        break
            if rate == 0:
                source = "No SFDC Opp" if df_sow is not None else "No rate"
        else:
            source = "NS rate"
        cur = str(row.get("currency", "USD") or "USD").strip().upper()
        per = str(row.get("period", ""))
        fx  = _get_fx(cur, per) if cur != "USD" else 1.0
        rev = round(float(row["hours"]) * rate * fx, 2)
        return rate, rev, source + (f" (×{fx:.4f} {cur}→USD)" if cur != "USD" and fx != 1.0 else "")

    tm[["_rate_local", "_revenue_usd", "_rate_source"]] = tm.apply(
        lambda r: pd.Series(_row_revenue(r)), axis=1)



    # ── Aggregate by project + period ────────────────────────────────────────
    # Ensure period has no NaT/None — groupby silently drops those rows
    if "period" in tm.columns:
        tm["period"] = tm["period"].fillna("").astype(str)
        tm["period"] = tm["period"].replace({"": "unknown", "NaT": "unknown", "nan": "unknown", "None": "unknown"})
    # Also ensure all groupby key columns have no NaN — groupby silently drops NaN keys
    for _gc in ["billing_flag", "project", "project_type", "currency", "project_manager"]:
        if _gc in tm.columns:
            tm[_gc] = tm[_gc].fillna("").astype(str)
    agg_cols = ["project", "project_type", "period", "billing_flag"] + (["currency"] if "currency" in tm.columns else [])
    if "project_manager" in tm.columns:
        agg_cols.append("project_manager")

    grp = (tm.groupby(agg_cols, as_index=False)
             .agg(hours=("hours", "sum"),
                  revenue_usd=("_revenue_usd", "sum"),
                  rate_local=("_rate_local", "max"),     # max for display only
                  rate_source=("_rate_source", "first")))

    rows = []
    for _, r in grp.iterrows():
        flag    = str(r.get("billing_flag", "") or "")
        pm      = r.get("project_manager", "") if "project_manager" in grp.columns else ""
        region  = _get_region(pm)
        prod    = match_product(str(r.get("project_type", "")))
        cur     = str(r.get("currency", "USD") or "USD").strip().upper() if "currency" in grp.columns else "USD"

        rows.append({
            "period":          str(r["period"]),
            "project":         r["project"],
            "product":         prod,
            "region":          region,
            "project_manager": pm,
            "currency":        cur,
            "hours":           float(r["hours"]),
            "rate_local":      float(r.get("rate_local", 0) or 0),
            "rate_usd":        float(r.get("rate_local", 0) or 0),
            "revenue_usd":     round(float(r["revenue_usd"]), 2),
            "rate_source":     str(r.get("rate_source", "") or ""),
            "billing_flag":    flag,
        })

    result = pd.DataFrame(rows)
    for _nc in ("hours", "rate_local", "rate_usd", "revenue_usd"):
        if _nc in result.columns:
            result[_nc] = pd.to_numeric(result[_nc], errors="coerce").fillna(0)
    return result


def get_billing_mismatches(df_ns: pd.DataFrame) -> pd.DataFrame:
    """Return NS time rows where billing type and non-billable flag conflict.

    Flags:
    - T&M + Non-Billable = Yes  → T&M project marked non-billable — may be misconfigured,
                                   revenue excluded pending review
    - Fixed Fee + Non-Billable = No → FF project with billable time entries — may appear
                                       on an invoice, review for overage
    """
    if df_ns is None or df_ns.empty or "billing_type" not in df_ns.columns:
        return pd.DataFrame()

    df = df_ns.copy()
    df["hours"] = pd.to_numeric(df["hours"] if "hours" in df.columns else pd.Series(0, index=df.index), errors="coerce").fillna(0)

    bt  = df["billing_type"].fillna("").str.strip().str.lower()
    nb_raw = df.get("non_billable", pd.Series("", index=df.index)).fillna("").astype(str).str.strip().str.lower()
    is_nb  = nb_raw.isin(["true","t","yes","1","y"])
    is_tm  = bt.str.contains("t&m|time")
    is_ff  = bt.str.contains("fixed fee|fixed.fee|fixed_fee|\bff\b")

    df["mismatch_flag"] = ""
    df.loc[is_tm &  is_nb, "mismatch_flag"] = "⚠️ T&M / Non-Billable"
    df.loc[is_ff & ~is_nb, "mismatch_flag"] = "⚠️ FF / Billable hours"
    # FF + Non-Billable (is_nb=True) is normal — no flag needed

    mismatches = df[df["mismatch_flag"] != ""].copy()

    keep_cols = ["employee","project","project_type","billing_type",
                 "non_billable","hours","mismatch_flag"]
    if "date" in mismatches.columns:
        keep_cols.insert(2, "date")
    if "project_manager" in mismatches.columns:
        keep_cols.insert(3, "project_manager")

    return mismatches[[c for c in keep_cols if c in mismatches.columns]]


def load_tm_sow(file) -> pd.DataFrame:
    """Load SFDC T&M SOW export.
    Returns one row per opportunity with canonical fields:
      account_name, opportunity_name, opportunity_owner, close_date,
      sow_hours, sow_amount_usd, sow_rate_usd, sow_rate_local,
      sow_rate_local_currency, product, region, fiscal_period
    """
    if hasattr(file, "name") and file.name.endswith(".csv"):
        for _enc in ("utf-8", "utf-8-sig", "windows-1252", "latin-1"):
            try:
                file.seek(0)
                df = pd.read_csv(file, encoding=_enc)
                break
            except (UnicodeDecodeError, AttributeError):
                continue
        else:
            file.seek(0)
            df = pd.read_csv(file, encoding="latin-1")
    else:
        df = pd.read_excel(file)

    df.columns = df.columns.str.strip()
    rename = {c: TM_SOW_COL_MAP[c.lower()] for c in df.columns if c.lower() in TM_SOW_COL_MAP}
    df = df.rename(columns=rename)

    # Parse dates
    for dc in ("close_date", "created_date"):
        if dc in df.columns:
            df[dc] = pd.to_datetime(df[dc], dayfirst=False, errors="coerce")

    # Numeric coercion
    for nc in ("sow_hours", "sow_amount_usd", "sow_rate_usd", "sow_rate_local", "arr_usd", "probability"):
        if nc in df.columns:
            df[nc] = pd.to_numeric(df[nc], errors="coerce").fillna(0)

    # Canonical product
    df["product"] = df.apply(
        lambda r: match_product_family(
            r.get("product_family", ""),
            r.get("products_text", "")
        ), axis=1
    )

    # Region — already NOAM/EMEA/APAC in SFDC report, just clean it
    if "region" in df.columns:
        df["region"] = df["region"].fillna("Other").str.strip()

    # Fiscal period → YYYY-MM for the close month (used for period bucketing)
    if "close_date" in df.columns:
        df["period"] = df["close_date"].dt.strftime("%Y-%m")

    # Opportunity owner name normalisation — "First Last" → keep as-is,
    # matching done at join time via name_matches pattern
    if "opportunity_owner" in df.columns:
        df["opportunity_owner"] = df["opportunity_owner"].fillna("").str.strip()

    return df


def join_tm_to_ns(df_sow: pd.DataFrame, df_ns: pd.DataFrame,
                  df_drs: pd.DataFrame = None) -> pd.DataFrame:
    """Join T&M SOW opportunities to NS time entries via fuzzy match.

    Match strategy:
    1. NS: rapidfuzz partial_ratio(sfdc_account, ns_project_name) >= 85
           + match_product(ns_project_type) == sfdc_product
    2. DRS: same fuzzy logic against DRS project_name (no hours yet)

    Returns df_sow with added columns:
      ns_project, ns_hours_worked, ns_revenue_to_date, ns_rate,
      rate_alignment, match_source
    """
    try:
        from rapidfuzz import fuzz as _fuzz
        def _acc_score(acc_l, target_l):
            """Score account name against a project/DRS name.
            Strips legal suffixes (Ltd, HQ, Inc etc.) and uses best of
            partial_ratio on stripped name vs WRatio on full name.
            """
            import re as _re
            stripped = _re.sub(
                r'\b(ltd|llc|inc|corp|hq|plc|pty|gmbh|srl|bv|ag|sa|nv|co\.?|'
                r'limited|incorporated|holding|holdings|group|international|'
                r'solutions|services|technologies|technology)\b\.?,?',
                '', acc_l, flags=_re.IGNORECASE
            )
            stripped = _re.sub(r'\s+', ' ', stripped).strip().strip(',').strip()
            s1 = _fuzz.partial_ratio(stripped, target_l) if stripped else 0
            s2 = _fuzz.WRatio(acc_l, target_l)
            return max(s1, s2)
    except ImportError:
        def _acc_score(acc_l, target_l):
            return 100.0 if acc_l.split()[0] in target_l else 0.0

    if df_ns is None or df_ns.empty:
        df_sow["ns_project"]         = None
        df_sow["ns_hours_worked"]    = 0.0
        df_sow["ns_revenue_to_date"] = 0.0
        df_sow["ns_rate"]            = 0.0
        df_sow["rate_alignment"]     = ""
        df_sow["match_source"]       = "No NS data"
        return df_sow

    # Build NS T&M project summary: hours + project_type per project
    # Includes: explicit T&M rows + FF projects with billable time (counts as T&M revenue)
    # billing_type absent → include all rows
    if "billing_type" in df_ns.columns:
        _bt_j   = df_ns["billing_type"].fillna("").str.lower()
        _nb_j   = df_ns.get("non_billable", pd.Series("", index=df_ns.index)).fillna("").astype(str).str.strip().str.lower()
        _is_nb_j = _nb_j.isin(["true","t","yes","1","y"])
        _is_tm_j = _bt_j.str.contains("t&m|time", na=False)
        _is_ff_bill_j = _bt_j.str.contains("fixed fee|fixed.fee|fixed_fee|\bff\b", na=False, regex=True) & ~_is_nb_j
        _tm_ns = df_ns[_is_tm_j | _is_ff_bill_j].copy()
    else:
        _tm_ns = df_ns.copy()
    _grp_cols = [c for c in ["project", "project_type"] if c in _tm_ns.columns]
    if not _grp_cols:
        df_sow["ns_project"]         = None
        df_sow["ns_hours_worked"]    = 0.0
        df_sow["ns_revenue_to_date"] = 0.0
        df_sow["ns_rate"]            = 0.0
        df_sow["rate_alignment"]     = ""
        df_sow["match_source"]       = "No NS project column"
        return df_sow
    ns_proj = (_tm_ns.groupby(_grp_cols, as_index=False)
                     .agg(hours=("hours","sum")) if "hours" in _tm_ns.columns
               else _tm_ns[_grp_cols].drop_duplicates().assign(hours=0.0))
    ns_rate_by_proj = {}
    if "ns_rate" in _tm_ns.columns:
        ns_rate_by_proj = (_tm_ns[_tm_ns["ns_rate"] > 0]
                           .groupby("project")["ns_rate"].max().to_dict())

    # DRS project list for tier 2
    drs_proj = []
    if df_drs is not None and not df_drs.empty and "project_name" in df_drs.columns:
        drs_proj = df_drs["project_name"].dropna().str.strip().tolist()

    ACC_THRESHOLD = 88  # tuned: passes real matches, blocks short-name false positives

    def _find_ns_project(opp_name, account_name, product, sow_rate_usd):
        acc_l = str(account_name or "").strip().lower()
        prod  = str(product or "").strip()
        sow_r = float(sow_rate_usd or 0)

        best_proj = None
        best_hrs  = 0.0
        match_src = ""

        # Tier 1: NS projects
        for _, nr in ns_proj.iterrows():
            proj_name = str(nr["project"])
            proj_l    = proj_name.lower()
            proj_type = str(nr.get("project_type", ""))

            # Account fuzzy match (strips legal suffixes, uses WRatio fallback)
            if _acc_score(acc_l, proj_l) < ACC_THRESHOLD:
                continue

            # Product match: NS project_type → canonical must equal SOW product
            ns_prod = match_product(proj_type)
            if prod and ns_prod != "Other" and ns_prod != prod:
                continue

            hrs = float(nr["hours"])
            if best_proj is None or hrs > best_hrs:
                best_proj = proj_name
                best_hrs  = hrs
                match_src = "NS match"

        # Tier 2: DRS (no hours yet)
        drs_match = None
        if best_proj is None and drs_proj:
            for dn in drs_proj:
                dn_l = dn.lower()
                if _acc_score(acc_l, dn_l) < ACC_THRESHOLD:
                    continue
                ns_prod = match_product(dn)
                if prod and ns_prod != "Other" and ns_prod != prod:
                    continue
                drs_match = dn
                match_src = "DRS match (no hours yet)"
                break

        if best_proj is None and drs_match is None:
            return None, 0.0, 0.0, 0.0, "", "Unmatched"

        ns_rate = float(ns_rate_by_proj.get(best_proj, 0) or 0) if best_proj else 0.0
        rev     = round(best_hrs * sow_r, 2) if best_proj else 0.0

        if best_proj and ns_rate > 0 and sow_r > 0:
            diff_pct = abs(ns_rate - sow_r) / sow_r * 100
            flag = (f"⚠️ NS ${ns_rate:,.0f} vs SOW ${sow_r:,.0f} ({diff_pct:.0f}% diff)"
                    if diff_pct > 5 else "✓ Rates aligned")
        elif best_proj and ns_rate > 0:
            flag = f"NS rate ${ns_rate:,.0f} — no SOW rate"
        elif drs_match:
            flag = "No hours in NS yet"
        else:
            flag = ""

        result_proj = best_proj if best_proj else drs_match
        return result_proj, round(best_hrs, 2), rev, ns_rate, flag, match_src

    results = df_sow.apply(
        lambda r: _find_ns_project(
            r.get("opportunity_name", ""),
            r.get("account_name", ""),
            r.get("product", ""),
            r.get("sow_rate_usd", 0)
        ), axis=1, result_type="expand"
    )
    results.columns = ["ns_project","ns_hours_worked","ns_revenue_to_date",
                        "ns_rate","rate_alignment","match_source"]
    out = pd.concat([df_sow.reset_index(drop=True), results], axis=1)
    return out.loc[:, ~out.columns.duplicated()]


def get_unmatched_sow(df_tm: pd.DataFrame) -> pd.DataFrame:
    """Return SOW opportunities that weren't matched to NS or DRS.
    Useful for vetting fuzzy match quality.
    """
    if df_tm is None or df_tm.empty or "match_source" not in df_tm.columns:
        return pd.DataFrame()
    unmatched = df_tm[df_tm["match_source"] == "Unmatched"].copy()
    keep = ["account_name","opportunity_name","opportunity_owner","product",
            "region","close_date","sow_hours","sow_amount_usd","sow_rate_usd","match_source"]
    return unmatched[[c for c in keep if c in unmatched.columns]]

# ── Product keyword matcher ───────────────────────────────────────────────────
# Maps subscription item / project type text → canonical product name
# Order matters — more specific entries first
_PRODUCT_KEYWORDS = [
    ("Capture and E-Invoicing", ["capture and e-invoic"]),
    ("Reconcile 2.0",           ["reconcile 2"]),
    # NS project_type prefixes: "ZoneBill: ZB_Premium", "ZonePay: Implementation" etc.
    # Match on Zone* prefix first (more specific)
    ("Billing",                 ["zonebill", "zbilling", "zab partner",
                                  "billing", "zb_standard", "zb_premium",
                                  "subscription services"]),
    ("Payroll",                 ["zonepay", "zpayroll", "payroll", "zep:"]),
    ("Reporting",               ["zonerpt", "zonerepor", "reporting",
                                  "install, dwh", "dwh"]),
    ("E-Invoicing",             ["e-invoic", "einvoic"]),
    ("Capture",                 ["zonecapture", "zcapture", "capture"]),
    ("Approvals",               ["zoneapprovals", "zapprovals", "approval",
                                  "zoneapps: consulting"]),
    ("Reconcile",               ["zonereconcile", "zreconcile", "reconcile"]),
    ("Payments",                ["zonepayments", "zpayment", "payment"]),
    ("PSP",                     ["psp"]),
    ("SFTP",                    ["sftp"]),
    ("CC",                      ["cc statement", "credit card"]),
    ("Premium",                 ["zb_premium", "premium"]),
]

def match_product(text: str) -> str:
    """Return canonical product name from subscription item / project type text.
    Returns 'Other' if no match found."""
    tl = str(text).strip().lower()
    for product, keywords in _PRODUCT_KEYWORDS:
        if any(kw in tl for kw in keywords):
            return product
    return "Other"


# ── Revenue charge loader ─────────────────────────────────────────────────────
REV_COL_MAP = {
    # Charge / SKU
    "charge item":              "charge_item",
    "charge":                   "charge_item",
    "item":                     "charge_item",
    # Subscription item
    "subscription item":        "subscription_item",
    "subscription":             "subscription_item",
    "item name":                "subscription_item",
    "service item":             "subscription_item",
    "subscription id":          "subscription_id",
    "subscription_id":          "subscription_id",
    # Project
    "project id":               "project_id",
    "project":                  "project_id",
    "job":                      "project_id",
    "project name":             "project_name",
    "name":                     "project_name",
    "customer":                 "project_name",
    # Revenue dates
    "rev rec start":            "rev_start",
    "rev rec start date":       "rev_start",
    "revenue recognition start":"rev_start",
    "rev start":                "rev_start",
    "rev start date":           "rev_start",
    "start date":               "rev_start",
    "service start date":       "rev_start",   # Finance report header
    "rev rec end":              "rev_end",
    "rev rec end date":         "rev_end",
    "revenue recognition end":  "rev_end",
    "rev end":                  "rev_end",
    "rev end date":             "rev_end",
    "end date":                 "rev_end",
    "service end date":         "service_end",
    "service end":              "service_end",
    # Currency + amounts
    "currency":                 "currency",
    "quantity":                 "quantity",
    "gross amount":             "gross_amount",
    "gross":                    "gross_amount",
    "rate":                     "rate",
    "discount":                 "discount",
    "amount":                   "amount",
    "net amount":               "amount",
    "rev carve amount":         "rev_carve_amount",
    "rev carve":                "rev_carve_amount",
    "carve out amount":         "rev_carve_amount",
    "carve amount":             "rev_carve_amount",
    "carve out":                "rev_carve_amount",
    # Status / transaction
    "status":                   "status",
    "transaction":              "transaction",
    "invoice":                  "transaction",
}

def load_revenue(file) -> pd.DataFrame:
    """Load NS FF revenue charge export.
    Derives:
      - recognizable_amount: Amount if > 0, else Rev Carve Amount
      - product: from subscription_item keyword match
      - region: from currency
      - monthly_slices: dict of {YYYY-MM: usd_amount} for straight-line recognition
    """
    from shared.config import CURRENCY_REGION_MAP, get_fx_rate

    if hasattr(file, "name") and file.name.endswith(".csv"):
        df = pd.read_csv(file)
    else:
        df = pd.read_excel(file)

    df.columns = df.columns.str.strip()
    # Build rename map — if multiple columns map to the same target, prefer more specific names
    _col_lower = {c: c.strip().lower() for c in df.columns}
    _seen_targets = set()
    _priority_order = ["project name", "project id", "subscription id", "service start date",
                       "service end date", "gross amount", "charge item", "subscription item"]
    # First pass: priority columns
    rename = {}
    for c, cl in _col_lower.items():
        if cl in _priority_order and cl in REV_COL_MAP:
            target = REV_COL_MAP[cl]
            rename[c] = target
            _seen_targets.add(target)
    # Second pass: remaining columns (skip if target already claimed)
    for c, cl in _col_lower.items():
        if c in rename: continue
        if cl in REV_COL_MAP:
            target = REV_COL_MAP[cl]
            if target not in _seen_targets:
                rename[c] = target
                _seen_targets.add(target)
    df = df.rename(columns=rename)

    # Parse dates — handle short year format (1/1/26) and standard formats
    def _parse_dates(series):
        try:
            return pd.to_datetime(series, dayfirst=False, errors="coerce")
        except Exception:
            return pd.to_datetime(series, errors="coerce")

    for dc in ("rev_start", "rev_end", "service_end"):
        if dc in df.columns:
            df[dc] = _parse_dates(df[dc])

    # Numeric coercion
    for nc in ("amount", "rev_carve_amount", "gross_amount", "rate"):
        if nc in df.columns:
            df[nc] = pd.to_numeric(df[nc], errors="coerce").fillna(0)

    # Recognizable amount:
    # 1. If Amount > 0 → use Amount (fully billed, recognize in full)
    # 2. If Amount = 0 → look up carve-out table by SKU + currency
    # 3. If not in table → fall back to Rev Carve Amount column from NS
    from shared.config import get_carve_out_amount
    from shared.config import (CAPTURE_IMPL_SKU, APPROVALS_IMPL_SKU,
                                RECONCILE_IMPL_SKU as _RIMPL)
    _CARVE_HANDLED_SKUS = {CAPTURE_IMPL_SKU, APPROVALS_IMPL_SKU, _RIMPL}

    def _strip_ci(s):
        s = str(s).strip()
        return s.split(" : ")[-1].strip() if " : " in s else s

    def _rec_amount(row):
        amt = float(row.get("amount", 0) or 0)
        if amt > 0:
            return amt
        # Skip carve table for SKUs handled by dedicated carve-out functions
        ci = _strip_ci(row.get("charge_item", ""))
        if ci in _CARVE_HANDLED_SKUS:
            # Will be set by product-specific carve function; use gross_amount as placeholder
            return float(row.get("gross_amount", 0) or 0)
        # Try SKU carve-out table for other SKUs
        tbl = get_carve_out_amount(row.get("charge_item", ""), row.get("currency", "USD"))
        if tbl is not None:
            return tbl
        # Fall back to rev_carve_amount column
        carve = float(row.get("rev_carve_amount", 0) or 0)
        if carve > 0:
            return carve
        # Last resort: gross_amount
        return float(row.get("gross_amount", 0) or 0)

    df["recognizable_amount"] = df.apply(_rec_amount, axis=1)

    # ── Apply product-specific carve-out logic ────────────────────────────────
    # Reconcile: license-based carve-out (must run before slicing)
    df = calc_reconcile_carveout(df)
    # Capture + Approvals: table-based carve-out with license validation
    df = calc_capture_approvals_carveout(df)
    # ── Deduplicate impl rows — keep earliest rev_start per project_id + SKU ────
    # Multiple charge rows can exist for multi-year licenses but only 1 impl project
    from shared.config import (CAPTURE_IMPL_SKU as _CIS, APPROVALS_IMPL_SKU as _AIS,
                                RECONCILE_IMPL_SKU as _RIS2)
    _impl_skus_dedup = {_CIS, _AIS, _RIS2}

    def _strip_for_dedup(s):
        s = str(s).strip()
        return s.split(" : ")[-1].strip() if " : " in s else s

    _is_impl = df["charge_item"].apply(_strip_for_dedup).isin(_impl_skus_dedup)
    if _is_impl.any() and "project_id" in df.columns:
        impl_df   = df[_is_impl].copy()
        other_df  = df[~_is_impl].copy()
        # Sort by rev_start ascending so earliest comes first
        impl_df["_rev_start_sort"] = pd.to_datetime(impl_df.get("rev_start", pd.Series()), errors="coerce")
        impl_df = impl_df.sort_values("_rev_start_sort", na_position="last")
        # Keep first (earliest) per project_id + charge_item
        impl_deduped = impl_df.drop_duplicates(subset=["project_id","charge_item"], keep="first")
        dupes = impl_df[~impl_df.index.isin(impl_deduped.index)]
        if not dupes.empty:
            # Flag dupes so they appear in Carve Flags tab
            dupes = dupes.copy()
            dupes["notes"] = dupes["notes"].fillna("") + " | ⚠️ Duplicate impl charge — excluded (same project_id, earlier row retained)"
            dupes["_exclude_from_slices"] = True
            impl_deduped["_rev_start_sort"] = None
            dupes["_rev_start_sort"] = None
            df = pd.concat([other_df, impl_deduped, dupes], ignore_index=True)
        else:
            impl_df.drop(columns=["_rev_start_sort"], inplace=True)

    # ── Derive rev rec window for ALL rows that don't have one yet ───────────
    import calendar as _cal_lr
    if "rev_rec_start" not in df.columns:
        df["rev_rec_start"] = None
    if "rev_rec_end" not in df.columns:
        df["rev_rec_end"] = None

    def _derive_window_lr(rev_start_val):
        try:
            ts = pd.Timestamp(rev_start_val)
            if ts.day == 1:
                tgt_mo = ts.month + 1; tgt_yr = ts.year
                if tgt_mo > 12: tgt_mo -= 12; tgt_yr += 1
                rre = pd.Timestamp(tgt_yr, tgt_mo, _cal_lr.monthrange(tgt_yr, tgt_mo)[1])
            else:
                tgt_mo = ts.month + 2; tgt_yr = ts.year
                while tgt_mo > 12: tgt_mo -= 12; tgt_yr += 1
                rre = pd.Timestamp(tgt_yr, tgt_mo, _cal_lr.monthrange(tgt_yr, tgt_mo)[1])
            return ts.strftime("%Y-%m-%d"), rre.strftime("%Y-%m-%d")
        except Exception:
            return None, None

    _needs_window = (
        df["rev_rec_start"].isna() | (df["rev_rec_start"].astype(str).str.strip().isin(["", "None", "NaT"]))
    )
    for _idx in df[_needs_window].index:
        _rs = df.at[_idx, "rev_start"]
        if pd.notna(_rs) and str(_rs).strip() not in ("", "NaT"):
            _rrs, _rre = _derive_window_lr(_rs)
            if _rrs:
                df.at[_idx, "rev_rec_start"] = _rrs
                df.at[_idx, "rev_rec_end"]   = _rre

    # Rename reconcile_flag → notes for unified audit trail
    if "reconcile_flag" in df.columns and "notes" not in df.columns:
        df = df.rename(columns={"reconcile_flag": "notes"})
    elif "reconcile_flag" in df.columns and "notes" in df.columns:
        # Merge: prefer notes content, fall back to reconcile_flag
        df["notes"] = df["notes"].where(df["notes"] != "", df["reconcile_flag"])
        df.drop(columns=["reconcile_flag"], inplace=True)

    # Product from subscription item
    df["product"] = df.get("subscription_item", pd.Series("", index=df.index)).apply(match_product)

    # Region from currency
    df["region"] = df.get("currency", pd.Series("USD", index=df.index)).apply(
        lambda c: CURRENCY_REGION_MAP.get(str(c).strip().upper(), "Other")
    )

    # project_id as string for joining
    if "project_id" in df.columns:
        df["project_id"] = df["project_id"].astype(str).str.strip().str.split(".").str[0]

    return df


def calc_reconcile_carveout(df: pd.DataFrame) -> pd.DataFrame:
    """Apply Reconcile carve-out logic to FF revenue charge DataFrame.

    For each SERV-APP-ZR2-STD_IMPL (Implementation) row:
      1. Derive rev_rec_start and rev_rec_end from Service Start Date (rev_start):
           - If rev_start is 1st of month → rev_rec_end = last day of following month
           - If rev_start is after 1st     → rev_rec_end = last day of 3rd month from start
      2. Find all Reconcile License rows with same subscription_id
      3. Calculate annual license cost: sum of each license row's gross_amount
         pro-rated to the 12-month window starting from impl rev_start
      4. Convert annual license cost to USD if non-USD
      5. carve_amount = min(table_max[license_sku], annual_license_cost_usd)
         If no license rows: carve = 0.00, flag
      6. Override recognizable_amount on impl row with carve_amount
      7. Set rev_rec_start / rev_rec_end on impl row
    """
    import calendar as _cal
    from shared.config import (RECONCILE_LICENSE_CARVE, RECONCILE_IMPL_SKU,
                                RECONCILE_LICENSE_SKUS, get_fx_rate)

    if df is None or df.empty:
        return df

    df = df.copy()
    df["reconcile_flag"]      = ""
    df["license_sku"]         = ""
    df["license_cost_usd"]    = None
    df["carve_max"]           = None
    df["impl_gross"]          = None
    df["rev_rec_start"]       = None
    df["rev_rec_end"]         = None
    df["_exclude_from_slices"] = False

    def _strip_sku(s):
        s = str(s).strip()
        # Split on all " : " separators and take the last segment
        # Handles "SERVICES : SERV-APP-ZR2-STD_IMPL" and "ZRR : BUNDLES : PROD-APP-ZR2_START15"
        return s.split(" : ")[-1].strip() if " : " in s else s

    df["_sku"] = df["charge_item"].apply(_strip_sku)

    impl_mask    = df["_sku"] == RECONCILE_IMPL_SKU
    license_mask = df["_sku"].isin(RECONCILE_LICENSE_SKUS)

    if not impl_mask.any():
        df["_exclude_from_slices"] = license_mask
        df.drop(columns=["_sku"], inplace=True)
        return df

    license_rows = df[license_mask].copy()
    impl_rows    = df[impl_mask].copy()

    def _derive_rev_rec_window(rev_start):
        """Derive rev_rec_start and rev_rec_end from service start date."""
        try:
            ts = pd.Timestamp(rev_start)
            rrs = ts  # rev_rec_start = service start
            if ts.day == 1:
                # 1st of month → end of following month
                next_mo = ts + pd.DateOffset(months=1)
                rre = pd.Timestamp(next_mo.year, next_mo.month,
                                   _cal.monthrange(next_mo.year, next_mo.month)[1])
            else:
                # After 1st → end of 3rd month from start
                third_mo = ts + pd.DateOffset(months=2)
                rre = pd.Timestamp(third_mo.year, third_mo.month,
                                   _cal.monthrange(third_mo.year, third_mo.month)[1])
            return rrs, rre
        except Exception:
            return None, None

    def _annual_license_cost(license_rows_sub, impl_start):
        """Sum license gross amounts pro-rated to the 12-month window from impl_start."""
        window_start = pd.Timestamp(impl_start)
        window_end   = window_start + pd.DateOffset(years=1) - pd.Timedelta(days=1)
        total = 0.0
        for _, lr in license_rows_sub.iterrows():
            lic_start = lr.get("rev_start")
            lic_end   = lr.get("service_end") or lr.get("rev_end")
            gross     = float(lr.get("gross_amount", 0) or 0)
            if pd.isna(lic_start) or lic_end is None or pd.isna(lic_end):
                total += gross  # can't pro-rate, include in full
                continue
            ls = pd.Timestamp(lic_start)
            le = pd.Timestamp(lic_end)
            lic_days    = max((le - ls).days, 1)
            overlap_s   = max(ls, window_start)
            overlap_e   = min(le, window_end)
            overlap_days = max((overlap_e - overlap_s).days, 0)
            total += gross * (overlap_days / lic_days)
        return total

    # Flag orphan licenses (no matching impl)
    if "subscription_id" in df.columns:
        impl_subs    = set(impl_rows["subscription_id"].dropna().astype(str))
        license_subs = set(license_rows["subscription_id"].dropna().astype(str))
        orphan       = license_subs - impl_subs
        if orphan:
            df.loc[license_mask & df["subscription_id"].astype(str).isin(orphan),
                   "reconcile_flag"] = "⚠️ License: no matching Implementation SKU"

    # Process each implementation row
    for idx, impl_row in impl_rows.iterrows():
        impl_start = impl_row.get("rev_start")
        impl_gross = float(impl_row.get("gross_amount", 0) or 0)

        # Store original impl gross for reviewer reference
        df.at[idx, "impl_gross"]        = impl_gross

        # Step 1: derive rev rec window
        rrs, rre = _derive_rev_rec_window(impl_start) if pd.notna(impl_start) else (None, None)
        df.at[idx, "rev_rec_start"] = rrs.strftime("%Y-%m-%d") if rrs else ""
        df.at[idx, "rev_rec_end"]   = rre.strftime("%Y-%m-%d") if rre else ""

        # Step 2: find license rows
        sub_id = str(impl_row.get("subscription_id", "")) if "subscription_id" in impl_row.index else ""
        if sub_id and "subscription_id" in df.columns:
            matched = license_rows[license_rows["subscription_id"].astype(str) == sub_id]
        else:
            matched = pd.DataFrame()

        if matched.empty:
            df.at[idx, "recognizable_amount"] = 0.0
            df.at[idx, "carve_max"]           = 0.0
            df.at[idx, "reconcile_flag"]      = "⚠️ No License SKU found — carve set to $0"
            continue

        # Identify license SKU — use earliest rev_start row; flag if mixed SKUs
        matched_sorted = matched.sort_values("rev_start") if "rev_start" in matched.columns else matched
        year1_lic      = matched_sorted.iloc[0]
        lic_sku        = str(year1_lic["_sku"])

        unique_skus = matched_sorted["_sku"].unique()
        mixed_flag  = len(unique_skus) > 1

        max_carve = RECONCILE_LICENSE_CARVE.get(lic_sku, 0.0)

        # Step 3: annual license cost (pro-rated 12-month window)
        if pd.notna(impl_start):
            annual_gross = _annual_license_cost(matched, impl_start)
        else:
            annual_gross = float(year1_lic.get("gross_amount", 0) or 0)

        # Step 4: convert to USD
        lic_cur = str(year1_lic.get("currency", "USD") or "USD").strip().upper()
        if lic_cur != "USD" and annual_gross > 0:
            try:
                period_str = pd.Timestamp(year1_lic.get("rev_start")).strftime("%Y-%m")
            except Exception:
                period_str = ""
            annual_gross_usd = annual_gross * get_fx_rate(lic_cur, period_str)
        else:
            annual_gross_usd = annual_gross

        # Step 5: carve = min(table max, annual license cost USD)
        carve_amount  = min(max_carve, annual_gross_usd) if annual_gross_usd > 0 else max_carve
        discount_flag = annual_gross_usd > 0 and annual_gross_usd < max_carve

        # Step 6: apply
        df.at[idx, "recognizable_amount"] = round(carve_amount, 2)
        df.at[idx, "license_sku"]         = lic_sku
        df.at[idx, "license_cost_local"]  = round(annual_gross, 2)
        df.at[idx, "license_currency"]    = lic_cur
        df.at[idx, "license_cost_usd"]    = round(annual_gross_usd, 2)
        df.at[idx, "carve_max"]           = max_carve

        flags = []
        if mixed_flag:
            flags.append(f"⚠️ Mixed License SKUs: {', '.join(unique_skus)} — using earliest ({lic_sku})")
        if discount_flag:
            flags.append(
                f"License: {lic_sku} · License Cost USD ${annual_gross_usd:,.2f} · "
                f"⚠️ Discount — Carve ${carve_amount:,.2f} (table max ${max_carve:,.2f})"
            )
        if not flags:
            flags.append(
                f"License: {lic_sku} · License Cost USD ${annual_gross_usd:,.2f} · Carve ${carve_amount:,.2f}"
            )
        df.at[idx, "reconcile_flag"] = " | ".join(flags)

    # Mark license rows for exclusion from slices
    df["_exclude_from_slices"] = df["_sku"].isin(RECONCILE_LICENSE_SKUS)
    df.drop(columns=["_sku"], inplace=True)
    return df

def calc_capture_approvals_carveout(df: pd.DataFrame) -> pd.DataFrame:
    """Apply Capture and Approvals carve-out logic to FF revenue charge DataFrame.

    For each SERV-APP-ZC_STD-IMPL or SERV-APP-ZA_STD-IMPL row:
      1. If gross_amount > 0 → recognizable_amount = gross_amount, notes = "Billed: $X"
      2. If gross_amount = 0 → recognizable_amount = carve table lookup (by SKU + currency)
         Capture non-CAD non-USD: convert $3,000 USD via FX
      3. Find matching license rows by subscription_id, annualise 12-month window
      4. If year1_license_cost_usd < carve_table_max → note the cap
      5. Flag orphan licenses (license with no matching impl)

    Notes column (renamed from reconcile_flag) accumulates all observations.
    """
    from shared.config import (CAPTURE_IMPL_SKU, CAPTURE_LICENSE_SKUS,
                                APPROVALS_IMPL_SKU, APPROVALS_LICENSE_SKUS,
                                FF_CARVE_OUT_TABLE, get_carve_out_amount, get_fx_rate)

    if df is None or df.empty:
        return df

    df = df.copy()
    if "notes" not in df.columns:
        df["notes"] = ""
    if "rev_rec_start" not in df.columns:
        df["rev_rec_start"] = None
    if "rev_rec_end" not in df.columns:
        df["rev_rec_end"] = None

    def _strip_sku(s):
        s = str(s).strip()
        return s.split(" : ")[-1].strip() if " : " in s else s

    df["_sku"] = df["charge_item"].apply(_strip_sku)

    all_impl_skus    = {CAPTURE_IMPL_SKU, APPROVALS_IMPL_SKU}
    all_license_skus = CAPTURE_LICENSE_SKUS | APPROVALS_LICENSE_SKUS

    impl_mask    = df["_sku"].isin(all_impl_skus)
    license_mask = df["_sku"].isin(all_license_skus)

    if not impl_mask.any():
        df["_exclude_from_slices_ca"] = license_mask.copy()
        df.drop(columns=["_sku"], inplace=True)
        return df

    license_rows = df[license_mask].copy()
    impl_rows    = df[impl_mask].copy()

    # Flag orphan licenses
    if "subscription_id" in df.columns:
        impl_subs    = set(impl_rows["subscription_id"].dropna().astype(str))
        license_subs = set(license_rows["subscription_id"].dropna().astype(str))
        orphans      = license_subs - impl_subs
        if orphans:
            df.loc[license_mask & df["subscription_id"].astype(str).isin(orphans),
                   "notes"] = "⚠️ Orphan License: no matching Implementation SKU"

    # Annual license cost helper (same as Reconcile)
    def _annual_license_cost(matched_lic, impl_start_val):
        try:
            window_start = pd.Timestamp(impl_start_val)
            window_end   = window_start + pd.DateOffset(years=1) - pd.Timedelta(days=1)
        except Exception:
            return sum(float(r.get("gross_amount", 0) or 0) for _, r in matched_lic.iterrows())
        total = 0.0
        for _, lr in matched_lic.iterrows():
            gross = float(lr.get("gross_amount", 0) or 0)
            ls    = lr.get("rev_start")
            le    = lr.get("service_end") or lr.get("rev_end")
            if not pd.notna(ls) or le is None or not pd.notna(le):
                total += gross
                continue
            try:
                ls = pd.Timestamp(ls); le = pd.Timestamp(le)
                lic_days     = max((le - ls).days, 1)
                overlap_s    = max(ls, window_start)
                overlap_e    = min(le, window_end)
                overlap_days = max((overlap_e - overlap_s).days, 0)
                total += gross * (overlap_days / lic_days)
            except Exception:
                total += gross
        return total

    for idx, impl_row in impl_rows.iterrows():
        sku        = str(impl_row["_sku"])
        curr       = str(impl_row.get("currency", "USD") or "USD").strip().upper()
        gross      = float(impl_row.get("gross_amount", 0) or 0)
        impl_start = impl_row.get("rev_start")
        notes_parts = []

        # Derive rev rec window for all impl rows (same rule as all FF)
        if pd.notna(impl_start):
            try:
                import calendar as _cal2
                ts = pd.Timestamp(impl_start)
                if ts.day == 1:
                    tgt_mo = ts.month + 1; tgt_yr = ts.year
                    if tgt_mo > 12: tgt_mo -= 12; tgt_yr += 1
                    rre = pd.Timestamp(tgt_yr, tgt_mo, _cal2.monthrange(tgt_yr, tgt_mo)[1])
                else:
                    tgt_mo = ts.month + 2; tgt_yr = ts.year
                    while tgt_mo > 12: tgt_mo -= 12; tgt_yr += 1
                    rre = pd.Timestamp(tgt_yr, tgt_mo, _cal2.monthrange(tgt_yr, tgt_mo)[1])
                df.at[idx, "rev_rec_start"] = ts.strftime("%Y-%m-%d")
                df.at[idx, "rev_rec_end"]   = rre.strftime("%Y-%m-%d")
            except Exception:
                pass

        # Step 1: gross_amount > 0 → use directly
        if gross > 0:
            df.at[idx, "recognizable_amount"] = round(gross, 2)
            notes_parts.append(f"Billed: {curr} {gross:,.2f}")
            df.at[idx, "notes"] = " | ".join(notes_parts)
            continue

        # Step 2: gross = 0 → table lookup
        period_str = ""
        if pd.notna(impl_start):
            try: period_str = pd.Timestamp(impl_start).strftime("%Y-%m")
            except Exception: pass

        carve_local = get_carve_out_amount(sku, curr, period_str)
        if carve_local is None:
            carve_local = 0.0

        # Get USD equivalent of carve for comparison
        if curr == "USD":
            carve_usd = carve_local
            fx_used   = 1.0
        else:
            fx_used   = get_fx_rate(curr, period_str) if curr != "USD" else 1.0
            # For Capture non-CAD non-USD: carve_local already is USD-converted amount
            if sku == CAPTURE_IMPL_SKU and curr not in ("USD", "CAD"):
                carve_usd = carve_local  # already USD
                carve_local_display = round(3000.00, 2)
                notes_parts.append(
                    f"Carve USD 3,000.00 → {curr} {carve_local:,.2f} @ FX {fx_used:.4f}"
                )
            else:
                carve_usd = carve_local * fx_used
                notes_parts.append(
                    f"Carve {curr} {carve_local:,.2f} → USD {carve_usd:,.2f} @ FX {fx_used:.4f}"
                )

        table_max_usd = FF_CARVE_OUT_TABLE.get((sku, "USD"), carve_usd)

        # Step 3: Year 1 license cost validation
        sub_id = str(impl_row.get("subscription_id", "")) if "subscription_id" in impl_row.index else ""
        if sub_id and "subscription_id" in df.columns:
            matched_lic = license_rows[license_rows["subscription_id"].astype(str) == sub_id]
        else:
            matched_lic = pd.DataFrame()

        if not matched_lic.empty:
            lic_currency = str(matched_lic.iloc[0].get("currency", "USD") or "USD").strip().upper()
            year1_local  = _annual_license_cost(matched_lic, impl_start)
            if lic_currency != "USD":
                lic_period = ""
                lic_start_val = matched_lic.iloc[0].get("rev_start")
                if pd.notna(lic_start_val):
                    try: lic_period = pd.Timestamp(lic_start_val).strftime("%Y-%m")
                    except Exception: pass
                lic_fx       = get_fx_rate(lic_currency, lic_period)
                year1_usd    = year1_local * lic_fx
            else:
                year1_usd = year1_local

            if year1_usd < table_max_usd:
                notes_parts.append(
                    f"⚠️ License Year 1 cost USD {year1_usd:,.2f} < table max USD {table_max_usd:,.2f} — carve capped"
                )
                # Cap carve at year1 cost
                if curr == "USD":
                    carve_local = min(carve_local, year1_usd)
                    carve_usd   = carve_local
                else:
                    carve_usd   = min(carve_usd, year1_usd)
                    carve_local = carve_usd * fx_used

        df.at[idx, "recognizable_amount"] = round(carve_usd, 2)
        if not notes_parts:
            notes_parts.append(f"Carve {curr} {carve_local:,.2f}")
        df.at[idx, "notes"] = " | ".join(notes_parts)

    # Exclude license rows from slices
    if "_sku" in df.columns:
        df["_exclude_from_slices_ca"] = df["_sku"].isin(all_license_skus)
        df.drop(columns=["_sku"], inplace=True)
    else:
        df["_exclude_from_slices_ca"] = False

    return df


def calc_monthly_slices(df: pd.DataFrame) -> pd.DataFrame:
    """Expand each charge row into monthly recognition slices.
    Returns a long-format DataFrame with columns:
      project_id, period (YYYY-MM), local_amount, usd_amount,
      currency, region, product, charge_item, subscription_item
    """
    from shared.config import get_fx_rate
    import calendar

    import calendar as _cal

    def _derive_ff_window(rev_start_val):
        """Derive rev rec start/end from service start date.
        1st of month → end of following month (2-month window)
        After 1st    → end of 3rd month from start (3-month window)
        Returns (rev_rec_start, rev_rec_end) as Timestamps, or (None, None) if invalid.
        """
        try:
            ts = pd.Timestamp(rev_start_val)
            if ts.day == 1:
                # End of following month
                if ts.month == 12:
                    rre = pd.Timestamp(ts.year + 1, 2, 1) - pd.Timedelta(days=1)
                else:
                    next_mo = ts.month + 1
                    yr = ts.year
                    last = _cal.monthrange(yr, next_mo)[1]
                    rre = pd.Timestamp(yr, next_mo, last)
            else:
                # End of 3rd month from start
                target_mo = ts.month + 2
                target_yr = ts.year
                while target_mo > 12:
                    target_mo -= 12
                    target_yr += 1
                last = _cal.monthrange(target_yr, target_mo)[1]
                rre = pd.Timestamp(target_yr, target_mo, last)
            return ts, rre
        except Exception:
            return None, None

    rows = []
    for _, r in df.iterrows():
        # Skip license rows — reference only, revenue recognised on impl row
        if r.get("_exclude_from_slices", False) or r.get("_exclude_from_slices_ca", False):
            continue

        total = float(r.get("recognizable_amount", 0) or 0)
        curr  = str(r.get("currency", "USD")).strip().upper()
        rev_start_raw = r.get("rev_start")

        # If rev_start is missing — flag, skip rev calc
        if pd.isna(rev_start_raw) or str(rev_start_raw).strip() == "":
            rows.append({
                "project_id":    str(r.get("project_id", "")),
                "project_name":  r.get("project_name", ""),
                "charge_item":   r.get("charge_item", ""),
                "subscription_item": r.get("subscription_item", ""),
                "subscription_id":   r.get("subscription_id", ""),
                "product":       r.get("product", "Other"),
                "region":        r.get("region", "Other"),
                "currency":      curr,
                "rev_start":     "",
                "rev_end":       "",
                "period":        "missing_start",
                "local_amount":  0.0,
                "usd_amount":    0.0,
                "status":        r.get("status", ""),
                "transaction":   r.get("transaction", ""),
                "notes": "⚠️ Missing rev_start — no rev calc until resolved",
                "license_sku":   r.get("license_sku", ""),
                "license_cost_local": r.get("license_cost_local", ""),
                "license_currency":   r.get("license_currency", ""),
                "license_cost_usd":   r.get("license_cost_usd", ""),
                "carve_max":     r.get("carve_max", ""),
                "impl_gross":    r.get("impl_gross", ""),
                "rev_rec_start": "",
                "rev_rec_end":   "",
            })
            continue

        # Use Reconcile-derived window if already set, otherwise derive for all FF rows
        _has_rr = (pd.notna(r.get("rev_rec_start")) and
                   str(r.get("rev_rec_start", "")).strip() not in ("", "NaT"))
        if _has_rr:
            start = pd.Timestamp(r.get("rev_rec_start"))
            end   = pd.Timestamp(r.get("rev_rec_end"))
        else:
            start, end = _derive_ff_window(rev_start_raw)
            if start is None:
                continue

        if total == 0:
            continue

        start = pd.Timestamp(start)
        end   = pd.Timestamp(end)

        # Build list of months in the recognition window
        months = []
        cur = start.to_period("M")
        end_p = end.to_period("M")
        while cur <= end_p:
            months.append(cur)
            cur += 1

        if not months:
            continue

        n_months = len(months)

        for mp in months:
            period_str = mp.strftime("%Y-%m")
            mo_start   = mp.to_timestamp()
            mo_end     = mp.to_timestamp("M")
            days_in_mo = calendar.monthrange(mp.year, mp.month)[1]

            # All FF rows: equal split across derived window, last month absorbs rounding
            # Reconcile rows already in USD (fx=1.0); others converted per period
            _is_reconcile_row = (pd.notna(r.get("rev_rec_start")) and
                                  str(r.get("rev_rec_start", "")).strip() not in ("", "NaT"))
            _mo_idx  = months.index(mp)
            _is_last = (_mo_idx == n_months - 1)
            if _is_last:
                _already = round(total / n_months, 2) * (n_months - 1)
                prorated = round(total - _already, 2)
            else:
                prorated = round(total / n_months, 2)
            fx = 1.0 if _is_reconcile_row else get_fx_rate(curr, period_str)

            rows.append({
                "project_id":        str(r.get("project_id", "")),
                "project_name":      r.get("project_name", ""),
                "charge_item":       r.get("charge_item", ""),
                "subscription_item": r.get("subscription_item", ""),
                "subscription_id":   r.get("subscription_id", ""),
                "product":           r.get("product", "Other"),
                "region":            r.get("region", "Other"),
                "currency":          curr,
                "service_start":     str(r.get("rev_start", ""))[:10] if pd.notna(r.get("rev_start")) else "",
                "service_end_orig":  str(r.get("service_end", "") or r.get("rev_end", ""))[:10],
                "rev_start":         str(r.get("rev_start", ""))[:10] if pd.notna(r.get("rev_start")) else "",
                "rev_end":           str(r.get("rev_end", "") or r.get("service_end", ""))[:10],
                "period":            period_str,
                "local_amount":      round(prorated, 2),
                "usd_amount":        round(prorated * fx, 2),
                "status":            r.get("status", ""),
                "transaction":       r.get("transaction", ""),
                "notes":              r.get("notes", r.get("reconcile_flag", "")),
                "license_sku":         r.get("license_sku", ""),
                "license_cost_local":  r.get("license_cost_local", ""),
                "license_currency":    r.get("license_currency", ""),
                "license_cost_usd":    r.get("license_cost_usd", ""),
                "carve_max":           r.get("carve_max", ""),
                "impl_gross":          r.get("impl_gross", ""),
                "rev_rec_start":       str(r.get("rev_rec_start", ""))[:10] if pd.notna(r.get("rev_rec_start")) else "",
                "rev_rec_end":         str(r.get("rev_rec_end", ""))[:10]   if pd.notna(r.get("rev_rec_end"))   else "",
            })

    result = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["project_id","charge_item","subscription_item","subscription_id",
                 "product","region","currency","period","local_amount","usd_amount",
                 "status","transaction","notes","license_sku",
                 "license_cost_usd","carve_max"]
    )
    # Ensure numeric columns are correct dtype
    for _nc in ("local_amount", "usd_amount"):
        if _nc in result.columns:
            result[_nc] = pd.to_numeric(result[_nc], errors="coerce").fillna(0)
    return result



