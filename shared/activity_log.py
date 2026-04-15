"""
Zone PS — Session Activity Logger
Tracks project-scoped interactions within the app and builds draft NS time entries.
"""
import streamlit as st
from datetime import datetime, date
import pandas as pd

# ── Activity type catalogue ───────────────────────────────────────────────────
ACTIVITY_TYPES = {
    "Customer Email":        {"default_hrs": 0.25, "ns_memo": "Customer communication - email"},
    "Project Review":        {"default_hrs": 0.50, "ns_memo": "Project review and status check"},
    "Re-engagement Outreach":{"default_hrs": 0.25, "ns_memo": "Re-engagement outreach"},
    "DRS Health Check":      {"default_hrs": 0.25, "ns_memo": "Project data review"},
    "Customer Engagement":   {"default_hrs": 0.50, "ns_memo": "Customer engagement activity"},
    "Internal Admin":        {"default_hrs": 0.25, "ns_memo": "Internal admin"},
    "Other":                 {"default_hrs": 0.25, "ns_memo": "PS activity"},
}

ACTIVITY_TYPE_LIST = list(ACTIVITY_TYPES.keys())

NS_HOUR_INCREMENTS = [0.25, 0.50, 0.75, 1.00, 1.25, 1.50, 1.75, 2.00,
                      2.50, 3.00, 3.50, 4.00, 4.50, 5.00, 6.00, 7.00, 8.00]

def _get_log() -> list:
    if "activity_log" not in st.session_state:
        st.session_state["activity_log"] = []
    return st.session_state["activity_log"]

def log_activity(project_id: str, project_name: str,
                 activity_type: str, employee: str = "",
                 notes: str = "", entry_date: date = None):
    """
    Append an activity entry to the session log.
    Called from any page where a project-scoped action occurs.
    """
    if not project_id or str(project_id).strip() in ("", "nan", "None"):
        return  # Don't log without a project_id

    log = _get_log()
    default_hrs = ACTIVITY_TYPES.get(activity_type, {}).get("default_hrs", 0.25)
    ns_memo     = ACTIVITY_TYPES.get(activity_type, {}).get("ns_memo", activity_type)

    entry = {
        "id":           len(log),
        "date":         (entry_date or date.today()).isoformat(),
        "project_id":   str(project_id).strip(),
        "project_name": str(project_name or "").strip(),
        "activity_type":activity_type,
        "hours":        default_hrs,
        "employee":     str(employee or "").strip(),
        "memo":         ns_memo,
        "notes":        str(notes or "").strip(),
        "logged_at":    datetime.now().isoformat(),
        "approved":     False,
    }
    log.append(entry)
    st.session_state["activity_log"] = log

def get_log_df() -> pd.DataFrame:
    """Return session activity log as a DataFrame."""
    log = _get_log()
    if not log:
        return pd.DataFrame(columns=["id","date","project_id","project_name",
                                      "activity_type","hours","employee","memo",
                                      "notes","logged_at","approved"])
    return pd.DataFrame(log)

def clear_log():
    st.session_state["activity_log"] = []

def log_count() -> int:
    return len(_get_log())

def to_ns_export(df: pd.DataFrame) -> pd.DataFrame:
    """
    Format approved entries for NS Time Entry import.
    Columns: Date, Project ID, Hours, Memo, Employee
    """
    if df.empty:
        return pd.DataFrame(columns=["Date","Project ID","Hours","Memo","Employee"])
    return df[["date","project_id","hours","memo","employee"]].rename(columns={
        "date":       "Date",
        "project_id": "Project ID",
        "hours":      "Hours",
        "memo":       "Memo",
        "employee":   "Employee",
    })
