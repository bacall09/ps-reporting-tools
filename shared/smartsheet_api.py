"""
PS Tools — Smartsheet API Integration
Read DRS sheet → DataFrame, Write row updates back to Smartsheet.

Secrets required (in .streamlit/secrets.toml):
    smartsheet_token   = "your-token-here"
    smartsheet_sheet_id = "your-sheet-id-here"

Column name → internal key mapping reuses SS_COL_MAP_OUT from template_utils.
Writeback maps internal keys back to original SS column names for the API.

Row identity: when loading via API, each row gets a hidden `_ss_row_id` column
containing the Smartsheet rowId. Writeback uses this to target the correct row.
"""

import requests
import pandas as pd
import streamlit as st
from datetime import date

from shared.template_utils import SS_COL_MAP_OUT

# ── Smartsheet API base ────────────────────────────────────────────────────────
_SS_BASE = "https://api.smartsheet.com/2.0"

# Inverse map: internal column name → Smartsheet display column name
# Used to look up the correct columnId when writing back.
# Where multiple SS headers map to the same internal key, we use the canonical one.
_INTERNAL_TO_SS = {}
for _ss_header, _internal in SS_COL_MAP_OUT.items():
    # Prefer the longer/more-specific header as canonical
    if _internal not in _INTERNAL_TO_SS or len(_ss_header) > len(_INTERNAL_TO_SS[_internal]):
        _INTERNAL_TO_SS[_internal] = _ss_header

# ── Fields that are writable back to Smartsheet ───────────────────────────────
# Internal key → display column name shown in My Projects editor
# The display name is used for editor column headers only.
# The actual SS column title is resolved via _INTERNAL_TO_SS → col_map.
WRITEBACK_FIELDS = {
    # ── Core project fields ───────────────────────────────────────────────────
    "phase":                    "Phase",
    "status":                   "Status",
    # rag intentionally excluded — calculated field, not writable

    # ── Weekly update / health fields ─────────────────────────────────────────
    "overall_summary":          "Overall Summary",
    "schedule_health":          "Schedule Health",
    "resource_health":          "Resource Health",
    "scope_health":             "Scope Health",
    "risk_level":               "Risk Level",
    "risk_detail":              "Risk Detail",
    "client_sentiment":         "Client Sentiment",
    "client_responsiveness":    "Client Responsiveness",

    # ── Project dates ─────────────────────────────────────────────────────────
    "go_live_date":             "Go-Live Date",
    "start_date":               "Start Date",
    "finish_date":              "Finish Date",

    # ── On hold fields ────────────────────────────────────────────────────────
    "on_hold_reason":           "On Hold Reason",
    "responsible_for_delay":    "Responsible for Delay",
    "responsible_for_delay":    "Responsible for Delay",
    "on_hold_response":         "On Hold Response",
    "support_transition_notes": "Support Transition Notes",
    "resume_date":              "Resume Date",
    "delay_summary":            "Delay Summary",

    # ── Other ─────────────────────────────────────────────────────────────────
    "jira_links":               "Jira Project",

    # ── Milestone dates ───────────────────────────────────────────────────────
    "ms_intro_email":           "Intro Email Sent",
    "ms_config_start":          "Config Start",
    "ms_enablement":            "Enablement Session",
    "ms_session1":              "Session #1",
    "ms_session2":              "Session #2",
    "ms_uat_signoff":           "UAT Signoff",
    "ms_prod_cutover":          "Prod Cutover",
    "ms_hypercare_start":       "Hypercare Start",
    "ms_close_out":             "Close Out Tasks",
    "ms_transition":            "Transition to Support",
}

# ── Explicit SS column title overrides ────────────────────────────────────────
# For fields where _INTERNAL_TO_SS (built from SS_COL_MAP_OUT) may not have
# the correct SS column title, we override here with the exact title from the
# DRS Blueprint data sheet.
_SS_TITLE_OVERRIDE = {
    "phase":                    "Project Phase",
    "overall_summary":          "Overall Summary",
    "schedule_health":          "Schedule Health",
    "resource_health":          "Resource Health",
    "scope_health":             "Scope Health",
    "risk_level":               "Risk Level",
    "risk_detail":              "Risk Detail",
    "client_sentiment":         "Client Sentiment",
    "client_responsiveness":    "Client Responsiveness",
    "go_live_date":             "Go-Live Date",
    "start_date":               "Start Date",
    "finish_date":              "Finish Date",
    "on_hold_reason":           "On Hold Reason",
    "on_hold_response":         "On Hold Response",
    "support_transition_notes": "Support Transition Notes",
    "resume_date":              "Resume Date",
    "delay_summary":            "Delay Summary",
    "jira_links":               "Jira Project",
    # Milestone SS column titles (exact from Blueprint)
    "ms_intro_email":           "Intro. Email Sent",
    "ms_config_start":          "Standard Configuration Set Up",
    "ms_enablement":            "Configuration Enablement Session",
    "ms_session1":              "Working Session 1 - Application Walkthrough",
    "ms_session2":              "Working Session 2 - Workshop / Q&A",
    "ms_uat_signoff":           "UAT Signoff",
    "ms_prod_cutover":          "Prod Cutover",
    "ms_hypercare_start":       "Hypercare Start",
    "ms_close_out":             "Close Out Remaining Tasks",
    "ms_transition":            "Project Closure / Transition to Support",
}


def _get_headers() -> dict:
    """Build auth headers from Streamlit secrets."""
    token = st.secrets.get("SMARTSHEET_TOKEN", "")
    if not token:
        raise ValueError("SMARTSHEET_TOKEN not found in Streamlit secrets.")
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/json",
    }


def _get_sheet_id() -> str:
    """Get sheet ID from Streamlit secrets."""
    sheet_id = st.secrets.get("SMARTSHEET_DRS_ID", "")
    if not sheet_id:
        raise ValueError("SMARTSHEET_DRS_ID not found in Streamlit secrets.")
    return str(sheet_id).strip()


def ss_available() -> bool:
    """Return True if Smartsheet secrets are configured."""
    try:
        return bool(st.secrets.get("SMARTSHEET_TOKEN")) and bool(st.secrets.get("SMARTSHEET_DRS_ID"))
    except Exception:
        return False


# ── READ ──────────────────────────────────────────────────────────────────────

def list_accessible_sheets() -> list[dict]:
    """Return a list of sheets the token can access."""
    headers = _get_headers()
    resp = requests.get(f"{_SS_BASE}/sheets", headers=headers, timeout=15)
    resp.raise_for_status()
    return [
        {"id": str(s["id"]), "name": s["name"], "permalink": s.get("permalink", "")}
        for s in resp.json().get("data", [])
    ]


def fetch_sheet_as_df() -> pd.DataFrame:
    """
    Fetch the full DRS Smartsheet and return as a raw DataFrame.
    Each row gets a `_ss_row_id` column with the Smartsheet internal rowId.
    """
    sheet_id = _get_sheet_id()
    headers  = _get_headers()

    resp = requests.get(
        f"{_SS_BASE}/sheets/{sheet_id}",
        headers=headers,
        params={"include": "rowPermalink"},
        timeout=30,
    )
    if resp.status_code == 401:
        raise PermissionError("Smartsheet token is invalid or expired.")
    if resp.status_code == 403:
        raise PermissionError("Smartsheet token does not have access to this sheet.")
    if resp.status_code == 404:
        raise ValueError(f"Sheet ID {sheet_id} not found.")
    resp.raise_for_status()

    data = resp.json()
    columns_def = data.get("columns", [])
    rows        = data.get("rows", [])

    col_id_to_title = {c["id"]: c["title"] for c in columns_def}

    records = []
    for row in rows:
        record = {"_ss_row_id": row["id"]}
        for cell in row.get("cells", []):
            col_title = col_id_to_title.get(cell.get("columnId"), "")
            if col_title:
                record[col_title] = cell.get("displayValue", cell.get("value", None))
        records.append(record)

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df.columns = [str(c) for c in df.columns]
    return df


def load_sheet_as_df() -> pd.DataFrame:
    """Fetch DRS and normalise through identical pipeline as load_drs()."""
    from shared.loaders import _normalise_drs_df
    raw = fetch_sheet_as_df()
    if raw.empty:
        return raw
    return _normalise_drs_df(raw)


# ── COLUMN ID CACHE ────────────────────────────────────────────────────────────

@st.cache_data(ttl=3600, show_spinner=False)
def _get_column_map(sheet_id: str, token: str) -> dict:
    """Return dict: lowercase column title → columnId. Cached 1 hour."""
    resp = requests.get(
        f"{_SS_BASE}/sheets/{sheet_id}/columns",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        timeout=15,
    )
    resp.raise_for_status()
    return {c["title"].lower().strip(): c["id"] for c in resp.json().get("data", [])}


# ── WRITE ──────────────────────────────────────────────────────────────────────

def _format_cell_value(internal_key: str, value) -> object:
    """Convert Python/pandas value to Smartsheet API format.
    Ensures all values are JSON-serialisable (no datetime objects, no NaN/NaT).
    """
    import datetime as _dt_mod
    if value is None:
        return None
    # Catch pandas NA/NaT/NaN
    try:
        if not isinstance(value, (str, bool, int, float)) and pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    # date / datetime / Timestamp → ISO string
    if isinstance(value, (_dt_mod.date, _dt_mod.datetime, pd.Timestamp)):
        try:
            return pd.Timestamp(value).strftime("%Y-%m-%d")
        except Exception:
            return None
    # Empty / sentinel strings
    if isinstance(value, str) and value.strip() in ("", "—", "nan", "None", "NaT"):
        return None
    # Float NaN
    if isinstance(value, float) and (value != value):
        return None
    # Ensure the result is JSON-serialisable — convert anything unexpected to str
    try:
        import json as _json
        _json.dumps(value, allow_nan=False)
        return value
    except (TypeError, ValueError):
        return str(value)


def write_row_updates(updates: list[dict]) -> tuple[int, list[str]]:
    """
    Write edited rows back to Smartsheet.

    `updates` is a list of dicts:
        {
            "_ss_row_id":   int,
            "project_name": str,
            "changes": {
                "phase": "03. Configuration",
                "schedule_health": "Green",
                ...
            }
        }

    Returns (success_count, list_of_error_messages).
    """
    if not updates:
        return 0, []

    sheet_id = _get_sheet_id()
    token    = st.secrets.get("SMARTSHEET_TOKEN", "")
    headers  = _get_headers()
    col_map  = _get_column_map(sheet_id, token)  # lowercase title → columnId

    def _col_id_for(internal_key: str) -> int | None:
        # 1. Check explicit override first (exact SS column title from Blueprint)
        override = _SS_TITLE_OVERRIDE.get(internal_key, "").lower().strip()
        if override and override in col_map:
            return col_map[override]
        # 2. Fall back to _INTERNAL_TO_SS (built from SS_COL_MAP_OUT)
        canonical = _INTERNAL_TO_SS.get(internal_key, "").lower().strip()
        if canonical and canonical in col_map:
            return col_map[canonical]
        # 3. Partial match on either
        for search in (override, canonical):
            if search:
                for title, cid in col_map.items():
                    if search in title:
                        return cid
        return None

    success_count = 0
    errors        = []

    BATCH_SIZE = 100
    for batch_start in range(0, len(updates), BATCH_SIZE):
        batch        = updates[batch_start : batch_start + BATCH_SIZE]
        rows_payload = []

        for upd in batch:
            row_id  = upd.get("_ss_row_id")
            changes = upd.get("changes", {})
            proj    = upd.get("project_name", str(row_id))

            if not row_id:
                errors.append(f"{proj}: missing _ss_row_id — skipped")
                continue

            cells = []
            for internal_key, new_val in changes.items():
                if internal_key not in WRITEBACK_FIELDS:
                    continue  # not a writable field — skip silently
                col_id = _col_id_for(internal_key)
                if col_id is None:
                    errors.append(f"{proj} / {internal_key}: column not found in sheet — skipped")
                    continue
                formatted = _format_cell_value(internal_key, new_val)
                cell_payload = {"columnId": int(col_id), "value": formatted}
                cells.append(cell_payload)

            if cells:
                rows_payload.append({"id": int(row_id), "cells": cells})

        if not rows_payload:
            continue

        # Final serialisation safety — convert any numpy/pandas types to plain Python
        import json as _json
        def _to_plain(obj):
            if isinstance(obj, dict):
                return {k: _to_plain(v) for k, v in obj.items()}
            if isinstance(obj, (list, tuple)):
                return [_to_plain(i) for i in obj]
            # numpy int/float (int64, float64, etc.)
            try:
                import numpy as _np
                if isinstance(obj, _np.integer): return int(obj)
                if isinstance(obj, _np.floating): return None if _np.isnan(obj) else float(obj)
                if isinstance(obj, _np.ndarray): return obj.tolist()
            except ImportError:
                pass
            # pandas Timestamp / NaT
            try:
                import pandas as _pd2
                if isinstance(obj, _pd2.Timestamp): return obj.strftime("%Y-%m-%d") if not _pd2.isna(obj) else None
                if obj is _pd2.NaT: return None
            except ImportError:
                pass
            # datetime.date / datetime.datetime
            import datetime as _dt2
            if isinstance(obj, (_dt2.date, _dt2.datetime)): return obj.isoformat()
            # float NaN
            if isinstance(obj, float) and (obj != obj): return None
            return obj

        rows_safe = _to_plain(rows_payload)

        resp = requests.put(
            f"{_SS_BASE}/sheets/{sheet_id}/rows",
            headers=headers,
            json=rows_safe,
            timeout=30,
        )

        if resp.status_code == 200:
            result = resp.json()
            success_count += len(result.get("result", rows_payload))
        else:
            try:
                err_detail = resp.json().get("message", resp.text)
            except Exception:
                err_detail = resp.text
            errors.append(f"Batch write failed ({resp.status_code}): {err_detail}")

    return success_count, errors
