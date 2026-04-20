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
    # Prefer the longer/more-specific header as canonical (e.g. "project name" over "name")
    if _internal not in _INTERNAL_TO_SS or len(_ss_header) > len(_INTERNAL_TO_SS[_internal]):
        _INTERNAL_TO_SS[_internal] = _ss_header

# ── Fields that are writable back to Smartsheet ───────────────────────────────
# Internal key → display column name shown in My Projects editor
WRITEBACK_FIELDS = {
    "phase":               "Phase",
    "status":              "Status",
    # rag is intentionally excluded — calculated field in Smartsheet, not writable
    "ms_intro_email":      "Intro Email Sent",
    "ms_config_start":     "Config Start",
    "ms_enablement":       "Enablement Session",
    "ms_session1":         "Session #1",
    "ms_session2":         "Session #2",
    "ms_uat_signoff":      "UAT Signoff",
    "ms_prod_cutover":     "Prod Cutover",
    "ms_hypercare_start":  "Hypercare Start",
    "ms_close_out":        "Close Out Tasks",
    "ms_transition":       "Transition to Support",
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

def fetch_sheet_as_df() -> pd.DataFrame:
    """
    Fetch the full DRS Smartsheet and return as a raw DataFrame.
    Column headers come from the sheet's column definitions.
    Each row gets a `_ss_row_id` column with the Smartsheet internal rowId.
    Raises on HTTP error or auth failure.
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

    # Build column index: columnId → column title
    col_id_to_title = {c["id"]: c["title"] for c in columns_def}

    # Build rows as dicts
    records = []
    for row in rows:
        record = {"_ss_row_id": row["id"]}
        for cell in row.get("cells", []):
            col_title = col_id_to_title.get(cell.get("columnId"), "")
            if col_title:
                # displayValue preferred (formatted text); fall back to value
                record[col_title] = cell.get("displayValue", cell.get("value", None))
        records.append(record)

    if not records:
        return pd.DataFrame()

    df = pd.DataFrame(records)
    df.columns = [str(c) for c in df.columns]
    return df


def load_sheet_as_df() -> pd.DataFrame:
    """
    Fetch DRS from Smartsheet API and run through identical normalisation
    as load_drs() in loaders.py. Returns ready-to-use df with _ss_row_id intact.
    """
    from shared.loaders import _normalise_drs_df
    raw = fetch_sheet_as_df()
    if raw.empty:
        return raw
    return _normalise_drs_df(raw)


# ── COLUMN ID CACHE ────────────────────────────────────────────────────────────

@st.cache_data(ttl=3600, show_spinner=False)
def _get_column_map(sheet_id: str, token: str) -> dict:
    """
    Return dict: lowercase column title → columnId.
    Cached for 1 hour — column structure rarely changes.
    """
    resp = requests.get(
        f"{_SS_BASE}/sheets/{sheet_id}/columns",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        timeout=15,
    )
    resp.raise_for_status()
    return {c["title"].lower().strip(): c["id"] for c in resp.json().get("data", [])}


# ── WRITE ──────────────────────────────────────────────────────────────────────

def _format_cell_value(internal_key: str, value) -> object:
    """
    Convert a Python/pandas value to the format Smartsheet expects.
    Dates → ISO string. None/NaT → None (clears cell). Everything else → str/num.
    """
    if value is None:
        return None
    # Pandas NaT
    if pd.isna(value) if not isinstance(value, (str, bool)) else False:
        return None
    # date / datetime → ISO 8601 string
    if isinstance(value, (pd.Timestamp, date)):
        try:
            return pd.Timestamp(value).strftime("%Y-%m-%d")
        except Exception:
            return None
    # Empty string → clear cell
    if isinstance(value, str) and value.strip() in ("", "—", "nan", "None", "NaT"):
        return None
    return value


def write_row_updates(updates: list[dict]) -> tuple[int, list[str]]:
    """
    Write edited rows back to Smartsheet.

    `updates` is a list of dicts, one per changed project row:
        {
            "_ss_row_id":   int,          # Smartsheet internal rowId
            "project_name": str,          # for logging/error messages
            "changes": {
                "phase": "03. Configuration",
                "ms_session1": pd.Timestamp("2026-05-01"),
                ...                       # any WRITEBACK_FIELDS keys
            }
        }

    Returns (success_count, list_of_error_messages).
    """
    if not updates:
        return 0, []

    sheet_id   = _get_sheet_id()
    token      = st.secrets.get("SMARTSHEET_TOKEN", "")
    headers    = _get_headers()
    col_map    = _get_column_map(sheet_id, token)  # lowercase title → columnId

    # Build reverse lookup: internal key → SS column title (lowercase)
    # We match against the col_map using the canonical SS header name.
    def _col_id_for(internal_key: str) -> int | None:
        canonical = _INTERNAL_TO_SS.get(internal_key, "").lower()
        if canonical in col_map:
            return col_map[canonical]
        # Fallback: try partial match
        for title, cid in col_map.items():
            if canonical and canonical in title:
                return cid
        return None

    success_count = 0
    errors        = []

    # Smartsheet accepts up to 100 rows per request — batch if needed
    BATCH_SIZE = 100
    for batch_start in range(0, len(updates), BATCH_SIZE):
        batch       = updates[batch_start : batch_start + BATCH_SIZE]
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
                    continue  # not a writeback field — skip silently
                col_id = _col_id_for(internal_key)
                if col_id is None:
                    errors.append(f"{proj} / {internal_key}: column not found in sheet — skipped")
                    continue
                formatted = _format_cell_value(internal_key, new_val)
                cell_payload = {"columnId": col_id, "value": formatted}
                cells.append(cell_payload)

            if cells:
                rows_payload.append({"id": row_id, "cells": cells})

        if not rows_payload:
            continue

        resp = requests.put(
            f"{_SS_BASE}/sheets/{sheet_id}/rows",
            headers=headers,
            json=rows_payload,
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
