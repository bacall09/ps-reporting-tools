"""
shared/whs.py
─────────────
Workload Health Score — shared scoring logic.
Used by pages/4_Workload_Health_Score.py and pages/1_Daily_Briefing.py.
"""
import pandas as pd
from shared.config import EMPLOYEE_LOCATION, PS_REGION_OVERRIDE, PS_REGION_MAP
from shared.constants import EMPLOYEE_ROLES

# ── Colours ───────────────────────────────────────────────────────────────────
GREEN = "#27AE60"
AMBER = "#F39C12"
RED   = "#E74C3C"

# ── Phase weights ─────────────────────────────────────────────────────────────
PHASE_WEIGHTS = {
    "00. onboarding":                    1.0,
    "01. requirements and design":       1.0,
    "02. configuration":                 2.0,
    "03. enablement/training":           2.5,
    "04. uat":                           3.0,
    "05. prep for go-live":              1.0,
    "06. go-live":                       3.0,
    "07. data migration":                1.0,
    "08. ready for support transition":  0.5,
    "09. phase 2 scoping":               1.0,
    "10. complete/pending final billing": 0.0,
    "11. on hold":                       0.25,
    "12. ps review":                     0.25,
}

# Phases excluded from active workload score
INACTIVE_PHASES = {"10. complete/pending final billing", "12. ps review"}

# Workload thresholds
WHS_LOW    = 25
WHS_MEDIUM = 60


def workload_level(score):
    """Return (label, colour) for a given workload score."""
    if score <= WHS_LOW:    return "Low",    GREEN
    if score <= WHS_MEDIUM: return "Medium", AMBER
    return "High", RED


def client_health_multiplier(responsiveness, sentiment):
    r = str(responsiveness).strip().lower() if pd.notna(responsiveness) else ""
    s = str(sentiment).strip().lower()      if pd.notna(sentiment)      else ""
    if ("negative" in s or "unresponsive" in r) and ("negative" in s and "unresponsive" in r):
        return 1.3
    if "negative" in s or "unresponsive" in r or "not responding" in r:
        return 1.15
    return 1.0


def risk_multiplier(risk_level):
    r = str(risk_level).strip().lower() if pd.notna(risk_level) else ""
    if "high"   in r: return 1.2
    if "medium" in r: return 1.1
    return 1.0


def get_phase_weight(phase):
    """Return (weight, normalised_phase_str)."""
    if not phase or str(phase).strip().lower() in ("", "nan", "none"):
        return 1.0, "Undefined"
    p = str(phase).strip().lower()
    if p in PHASE_WEIGHTS:
        return PHASE_WEIGHTS[p], str(phase).strip()
    return 1.0, str(phase).strip()


def get_ps_region(name):
    if not name or str(name).strip().lower() in ("", "nan", "none"):
        return "Unknown"
    name = str(name).strip()
    if name in PS_REGION_OVERRIDE:
        return PS_REGION_OVERRIDE[name]
    loc = EMPLOYEE_LOCATION.get(name, "")
    if isinstance(loc, tuple):
        loc = loc[0]
    return PS_REGION_MAP.get(loc, "Unknown")


def score_projects(ss_df, ns_df=None):
    """
    Compute per-project weighted score (FF only, excl inactive phases).
    ns_df is accepted for API compatibility but PM is read from SS directly.
    Returns scored DataFrame.
    """
    df = ss_df.copy()

    def is_active_phase(phase):
        if not phase or str(phase).strip().lower() in ("", "nan", "none"):
            return True
        return str(phase).strip().lower() not in INACTIVE_PHASES

    # PM from SS
    _ss_pm_col = next((col for col in ["project_manager", "consultant"] if col in df.columns), None)
    if _ss_pm_col and _ss_pm_col != "project_manager":
        df["project_manager"] = df[_ss_pm_col]
    elif _ss_pm_col is None:
        df["project_manager"] = None
    df["pm_flag"] = df["project_manager"].isna()

    # Phase weight
    df["phase_weight"] = df["phase"].apply(lambda p: get_phase_weight(p)[0])

    # Multipliers
    resp_col = "client_responsiveness" if "client_responsiveness" in df.columns else None
    sent_col = "client_sentiment"      if "client_sentiment"      in df.columns else None
    risk_col = "risk_level"            if "risk_level"            in df.columns else None

    df["client_health_mult"] = df.apply(
        lambda r: client_health_multiplier(
            r[resp_col] if resp_col else None,
            r[sent_col] if sent_col else None,
        ), axis=1
    )
    df["risk_mult"] = df[risk_col].apply(risk_multiplier) if risk_col else 1.0

    def is_active_row(row):
        phase_inactive = not is_active_phase(row.get("phase", ""))
        status_val     = str(row.get("status", "")).strip().lower()
        status_onhold  = status_val in ("on-hold", "on hold", "onhold", "on_hold")
        return not phase_inactive and not status_onhold

    df["active"] = df.apply(is_active_row, axis=1)

    complete_phases = {"10. complete/pending final billing"}
    df["total_project"] = df["phase"].apply(
        lambda p: str(p).strip().lower() not in complete_phases if pd.notna(p) else True
    )

    _tm_mask = (
        df["project_type"].str.lower().str.contains("t&m|time.*material", na=False, regex=True)
        if "project_type" in df.columns
        else pd.Series(False, index=df.index)
    )
    df["weighted_score"] = df.apply(
        lambda r: round(r["phase_weight"] * r["client_health_mult"] * r["risk_mult"], 2)
        if (r["active"] and not _tm_mask.loc[r.name]) else 0.0,
        axis=1,
    )

    df["ps_region"] = df["project_manager"].apply(get_ps_region)
    df["role"] = df["project_manager"].apply(
        lambda n: EMPLOYEE_ROLES.get(str(n).strip(), {}).get("role", "Consultant")
        if pd.notna(n) else ""
    )

    return df


def build_consultant_summary(scored_df, ss_df=None):
    """Aggregate scored projects by consultant. Returns (summary_df, missing_pm_count)."""
    active = scored_df[scored_df["active"]].copy()

    grp = active.groupby("project_manager").agg(
        total_score=("weighted_score", "sum"),
        ps_region=("ps_region", "first"),
        role=("role", "first"),
    ).reset_index()

    grp["total_score"]    = grp["total_score"].round(1)
    grp["workload_level"] = grp["total_score"].apply(lambda s: workload_level(s)[0])
    grp = grp.sort_values("total_score", ascending=False).reset_index(drop=True)

    _pm_col = next((col for col in ["project_manager", "consultant"] if col in ss_df.columns), None) if ss_df is not None else None
    if ss_df is not None and _pm_col is not None:
        _complete_ph = {"10. complete/pending final billing", "12. ps review"}
        _onhold_st   = {"on-hold", "on hold", "onhold", "on_hold"}
        _tm_types    = {"t&m", "time & material", "time and material"}
        _id_col      = "project_id" if "project_id" in ss_df.columns else "project_name"
        _total_map, _active_map = {}, {}
        for _cons, _g in ss_df.groupby(_pm_col):
            _billing      = (_g["billing_type"].astype(str).str.strip().str.lower()
                             if "billing_type" in _g.columns else pd.Series("", index=_g.index))
            _phase        = (_g["phase"].astype(str).str.strip().str.lower()
                             if "phase" in _g.columns else pd.Series("", index=_g.index))
            _stat         = (_g["status"].astype(str).str.strip().str.lower()
                             if "status" in _g.columns else pd.Series("", index=_g.index))
            _ff           = ~_billing.isin(_tm_types)
            _not_complete = ~_phase.isin(_complete_ph)
            _not_onhold   = ~_stat.isin(_onhold_st)
            _total_map[str(_cons).strip()]  = int(_g[_ff & _not_complete][_id_col].nunique())
            _active_map[str(_cons).strip()] = int(_g[_ff & _not_complete & _not_onhold][_id_col].nunique())
        grp["total_project_count"]  = grp["project_manager"].map(_total_map).fillna(0).astype(int)
        grp["active_project_count"] = grp["project_manager"].map(_active_map).fillna(0).astype(int)
    else:
        total = scored_df[scored_df["total_project"]].copy() if "total_project" in scored_df.columns else scored_df.copy()
        active_counts = active.groupby("project_manager")["project_id"].nunique()
        total_counts  = total.groupby("project_manager")["project_id"].nunique()
        grp["active_project_count"] = grp["project_manager"].map(active_counts).fillna(0).astype(int)
        grp["total_project_count"]  = grp["project_manager"].map(total_counts).fillna(0).astype(int)

    missing = int(scored_df[scored_df["pm_flag"] == True]["project_id"].nunique()) if "project_id" in scored_df.columns else 0
    return grp, missing


def consultant_whs(name, ss_df):
    """
    Return (score, label, colour) for a single consultant given a DRS DataFrame.
    Convenience wrapper for Daily Briefing metric card.
    """
    if ss_df is None or ss_df.empty:
        return None, "—", "#718096"
    try:
        scored   = score_projects(ss_df)
        summary, _ = build_consultant_summary(scored, ss_df)
        row = summary[summary["project_manager"].astype(str).str.strip() == name.strip()]
        if row.empty:
            return 0.0, "Low", GREEN
        score = float(row.iloc[0]["total_score"])
        label, colour = workload_level(score)
        return score, label, colour
    except Exception:
        return None, "—", "#718096"
