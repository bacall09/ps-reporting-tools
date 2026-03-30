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
        df["contact_name"] = (df["first_name"].fillna("") + " " + df["last_name"].fillna("")).str.strip()
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
    # Parse all milestone date columns
    for _ms_col in MILESTONE_COLS_MAP.keys():
        if _ms_col in df.columns:
            df[_ms_col] = pd.to_datetime(df[_ms_col], errors="coerce")

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
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
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
    "employee":             "employee",
    "name":                 "employee",
    "project":              "project",
    "project name":         "project",
    "project id":           "project_id",
    "project id ":          "project_id",
    " project id":          "project_id",
    "project_id":           "project_id",
    "projectid":            "project_id",
    "project internal id":  "project_id",
    "job":                  "project_id",
    "billing type":         "billing_type",
    "project type":         "project_type",
    "type":                 "project_type",
    "project manager":      "project_manager",
    "date":                 "date",
    "transaction date":     "date",
    "hours":                "hours",
    "quantity":             "hours",
    "hours/quantity":       "hours",
    "hours to date":        "hours_to_date",
    "quantity to date":     "hours_to_date",
    "hours/quantity to date": "hours_to_date",
    "non-billable":         "non_billable",
    "non billable":         "non_billable",
    "nonbillable":          "non_billable",
    "non_billable":         "non_billable",
    "is non billable":      "non_billable",
    # Billing rate — add to NS Time Detail export to enable direct T&M revenue calc
    "rate":                 "ns_rate",
    "billing rate":         "ns_rate",
    "hourly rate":          "ns_rate",
    "time detail rate":     "ns_rate",
    "billing rate (foreign currency)": "ns_rate",
    "rate (foreign currency)":         "ns_rate",
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

    # Filter T&M rows
    tm = df_ns[df_ns.get("billing_type", pd.Series(dtype=str))
               .fillna("").str.lower().str.contains("t&m|time")].copy()
    if tm.empty:
        return pd.DataFrame()

    # Ensure period column
    if "period" not in tm.columns:
        if "date" in tm.columns:
            tm["date"]   = pd.to_datetime(tm["date"], errors="coerce")
            tm["period"] = tm["date"].dt.strftime("%Y-%m")
        else:
            return pd.DataFrame()

    tm["hours"] = pd.to_numeric(tm.get("hours", 0), errors="coerce").fillna(0)

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

    # ── Aggregate by project + period ────────────────────────────────────────
    agg_cols = ["project", "project_type", "period"]
    if "project_manager" in tm.columns:
        agg_cols.append("project_manager")
    if "ns_rate" in tm.columns:
        tm["ns_rate"] = pd.to_numeric(tm["ns_rate"], errors="coerce").fillna(0)
        # Use max rate per project-period (rates should be consistent per project)
        grp_hrs  = tm.groupby(agg_cols, as_index=False)["hours"].sum()
        grp_rate = tm.groupby(agg_cols, as_index=False)["ns_rate"].max()
        grp = grp_hrs.merge(grp_rate, on=agg_cols, how="left")
    else:
        grp = tm.groupby(agg_cols, as_index=False)["hours"].sum()

    rows = []
    for _, r in grp.iterrows():
        proj_l   = str(r["project"]).strip().lower()
        ns_rate  = float(r.get("ns_rate", 0) or 0) if "ns_rate" in grp.columns else 0.0

        if ns_rate > 0:
            # Best source: rate directly from NS Time Detail export
            rate   = ns_rate
            source = "NS rate"
        else:
            # Fall back to SOW project match
            rate   = rate_by_project.get(proj_l, 0.0)
            source = "Project match" if rate > 0 else ""

            # Fall back to account name substring
            if rate == 0:
                for acc, acc_rate in rate_by_account.items():
                    if acc in proj_l:
                        rate   = acc_rate
                        source = "Account match"
                        break

            if rate == 0:
                source = "No SFDC Opp" if df_sow is not None else "No rate"
        pm      = r.get("project_manager", "") if "project_manager" in grp.columns else ""
        region  = _get_region(pm)
        prod    = match_product(str(r.get("project_type", "")))
        hrs     = float(r["hours"])
        rows.append({
            "period":          r["period"],
            "project":         r["project"],
            "product":         prod,
            "region":          region,
            "project_manager": pm,
            "hours":           hrs,
            "rate_usd":        rate,
            "revenue_usd":     round(hrs * rate, 2),
            "rate_source":     source,
        })

    result = pd.DataFrame(rows)
    for _nc in ("hours", "rate_usd", "revenue_usd"):
        result[_nc] = pd.to_numeric(result[_nc], errors="coerce").fillna(0)
    return result


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


def join_tm_to_ns(df_sow: pd.DataFrame, df_ns: pd.DataFrame) -> pd.DataFrame:
    """Join T&M SOW opportunities to NS time entries via fuzzy account+product match.
    
    Match strategy (in order):
    1. Exact project name match (opp_name ≈ NS project name)
    2. Account name fuzzy match + product match
    
    Returns df_sow with added columns:
      ns_project, ns_hours_worked, ns_revenue_to_date
    """
    if df_ns is None or df_ns.empty:
        df_sow["ns_project"]          = None
        df_sow["ns_hours_worked"]     = 0.0
        df_sow["ns_revenue_to_date"]  = 0.0
        return df_sow

    # Build NS project summary: total hours + avg ns_rate per project
    _tm_ns = df_ns[df_ns["billing_type"].fillna("").str.lower().str.contains("t&m|time")]
    ns_proj = (_tm_ns.groupby("project")
               .agg(hours=("hours", "sum"))
               .reset_index())

    # Build NS rate per project (max rate — should be consistent per project)
    ns_rate_by_proj = {}
    if "ns_rate" in _tm_ns.columns:
        _rates = _tm_ns[_tm_ns["ns_rate"] > 0].groupby("project")["ns_rate"].max()
        ns_rate_by_proj = _rates.to_dict()

    def _find_ns_project(opp_name, account_name, product, sow_rate_usd):
        """Return (ns_project, hours_worked, revenue, ns_rate, rate_flag) for best NS match."""
        on_l   = str(opp_name or "").strip().lower()
        acc_l  = str(account_name or "").strip().lower()
        prod_l = str(product or "").strip().lower()

        best_proj = None
        best_hrs  = 0.0

        for _, nr in ns_proj.iterrows():
            proj_l = str(nr["project"]).strip().lower()
            acc_score  = 1 if acc_l and acc_l[:8] in proj_l else 0
            prod_score = 1 if prod_l and prod_l in proj_l else 0
            opp_score  = 1 if on_l and on_l[:10] in proj_l else 0

            if acc_score + prod_score + opp_score >= 2:
                if best_proj is None or nr["hours"] > best_hrs:
                    best_proj = nr["project"]
                    best_hrs  = float(nr["hours"])

        if not best_proj:
            return None, 0.0, 0.0, 0.0, ""

        ns_rate = float(ns_rate_by_proj.get(best_proj, 0) or 0)
        sow_r   = float(sow_rate_usd or 0)
        rev     = round(best_hrs * sow_r, 2)

        # Rate alignment flag
        if ns_rate > 0 and sow_r > 0:
            diff_pct = abs(ns_rate - sow_r) / sow_r * 100
            if diff_pct > 5:
                flag = f"⚠️ NS ${ns_rate:,.0f} vs SOW ${sow_r:,.0f} ({diff_pct:.0f}% diff)"
            else:
                flag = "✓ Rates aligned"
        elif ns_rate > 0 and sow_r == 0:
            flag = f"NS rate ${ns_rate:,.0f} — no SOW rate"
        else:
            flag = ""

        return best_proj, round(best_hrs, 2), rev, ns_rate, flag

    results = df_sow.apply(
        lambda r: _find_ns_project(
            r.get("opportunity_name",""),
            r.get("account_name",""),
            r.get("product",""),
            r.get("sow_rate_usd", 0)
        ), axis=1, result_type="expand"
    )
    results.columns = ["ns_project", "ns_hours_worked", "ns_revenue_to_date",
                        "ns_rate", "rate_alignment"]
    return pd.concat([df_sow.reset_index(drop=True), results], axis=1)

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
    # Project
    "project id":               "project_id",
    "project":                  "project_id",
    "job":                      "project_id",
    # Revenue dates
    "rev rec start":            "rev_start",
    "rev rec start date":       "rev_start",
    "revenue recognition start":"rev_start",
    "rev start":                "rev_start",
    "rev start date":           "rev_start",
    "start date":               "rev_start",
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
    rename = {c: REV_COL_MAP[c.lower()] for c in df.columns if c.lower() in REV_COL_MAP}
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
    def _rec_amount(row):
        amt = float(row.get("amount", 0) or 0)
        if amt > 0:
            return amt
        # Try SKU table first
        tbl = get_carve_out_amount(row.get("charge_item", ""), row.get("currency", "USD"))
        if tbl is not None:
            return tbl
        # Fall back to column value
        return float(row.get("rev_carve_amount", 0) or 0)

    df["recognizable_amount"] = df.apply(_rec_amount, axis=1)

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


def calc_monthly_slices(df: pd.DataFrame) -> pd.DataFrame:
    """Expand each charge row into monthly recognition slices.
    Returns a long-format DataFrame with columns:
      project_id, period (YYYY-MM), local_amount, usd_amount,
      currency, region, product, charge_item, subscription_item
    """
    from shared.config import get_fx_rate
    import calendar

    rows = []
    for _, r in df.iterrows():
        start = r.get("rev_start")
        end   = r.get("rev_end")
        total = float(r.get("recognizable_amount", 0) or 0)
        curr  = str(r.get("currency", "USD")).strip().upper()

        if pd.isna(start) or pd.isna(end) or total == 0:
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

            # Pro-rate partial first / last months
            actual_start = max(start, mo_start)
            actual_end   = min(end, mo_end)
            days_active  = (actual_end - actual_start).days + 1

            # Monthly slice = total / n_months, scaled by days active
            full_slice    = total / n_months
            prorated      = full_slice * (days_active / days_in_mo)
            fx            = get_fx_rate(curr, period_str)

            rows.append({
                "project_id":        str(r.get("project_id", "")),
                "charge_item":       r.get("charge_item", ""),
                "subscription_item": r.get("subscription_item", ""),
                "product":           r.get("product", "Other"),
                "region":            r.get("region", "Other"),
                "currency":          curr,
                "period":            period_str,
                "local_amount":      round(prorated, 2),
                "usd_amount":        round(prorated * fx, 2),
                "status":            r.get("status", ""),
                "transaction":       r.get("transaction", ""),
            })

    result = pd.DataFrame(rows) if rows else pd.DataFrame(
        columns=["project_id","charge_item","subscription_item","product",
                 "region","currency","period","local_amount","usd_amount",
                 "status","transaction"]
    )
    # Ensure numeric columns are correct dtype
    for _nc in ("local_amount", "usd_amount"):
        if _nc in result.columns:
            result[_nc] = pd.to_numeric(result[_nc], errors="coerce").fillna(0)
    return result



