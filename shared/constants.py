"""
PS Tools — Shared Constants
Employee roster, role assignments, column maps, templates.
Updated: 2026-03-22 v2
"""
import re as _re_constants

# ── Streamlit permission roles ─────────────────────────────────────────────────
# Manager-only: see all pages, do NOT appear in consultant dropdowns
MANAGERS_ONLY = []  # Login handles identity — no manager-only access tier needed

# Manager + Consultant: see all pages AND appear in consultant-scoped views
MANAGER_CONSULTANTS = [
    "Barrio, Nairobi",     # Project Manager
    "Cadelina, Macoy",     # Project Manager
    "Hopkins, Chris",      # Team Lead — Capture & Approvals (NOAM)
    "Hughes, Madalyn",     # Project Manager
    "Ickler, Georganne",   # Consultant & Manager of PS
    "Lappin, Thomas",      # Manager-level Consultant
    "Longi, Sameer",               # Director of PS
    "Murphy, Conor",       # Solution Architect — manager tier
    "Porangada, Suraj",    # Project Manager
    "Prince, Trevor",              # VP of PMO
    "Rusnak, Connor",              # VP of PS
    "Snee, Stefanie J",    # Manager-level Consultant
    "Stone, Matt",         # Manager-level Consultant
    "Swanson, Patti",      # Consultant & Director of PS
]

# Reporting only: access to management reports, do NOT appear in consultant dropdowns
REPORTING_ONLY = [
    "Soares, Erica",    # Finance — reporting access only
    "Lindahl, Jenni",   # Reporting access only
]

# No access (leavers or no Streamlit access)
NO_ACCESS = [
    "Alam, Laisa",          # No longer employed
    "Chan, Joven",          # No longer employed
    "Eyong, Eyong",         # No longer employed
    "Hernandez, Camila",    # No longer employed
    "Centinaje, Rhodechild",# Left March 16 2026
    "Cloete, Bronwyn",      # Left Feb 23 2026 — not part of PS org
]

def get_role(name: str) -> str:
    """Return 'manager_only' | 'manager' | 'consultant' | 'reporting_only' | 'no_access'."""
    if any(name.startswith(m) for m in MANAGERS_ONLY):
        return "manager_only"
    if name in NO_ACCESS:
        return "no_access"
    if name in REPORTING_ONLY:
        return "reporting_only"
    if name in MANAGER_CONSULTANTS:
        return "manager"
    return "consultant"

def is_manager(name: str) -> bool:
    return get_role(name) in ("manager", "manager_only")

def is_consultant(name: str) -> bool:
    return get_role(name) in ("consultant", "manager")

# ── Employee roster ────────────────────────────────────────────────────────────
# util_exempt: True = no utilisation targets (managers, solution architects etc.)
EMPLOYEE_ROLES = {
    # ── Reporting Only (Finance / external reporting access) ─────────────────
    "Soares, Erica":          {"role": "Reporting",  "products": [], "learning": [], "util_exempt": True},
    "Lindahl, Jenni":         {"role": "Reporting",  "products": [], "learning": [], "util_exempt": True},
    # ── Project Managers (no product delivery) ────────────────────────────────
    "Barrio, Nairobi":        {"role": "Project Manager",    "products": [], "learning": []},
    "Cadelina, Macoy":        {"role": "Project Manager",    "products": [], "learning": [], "note": "New as of March 3 2026"},
    "Hughes, Madalyn":        {"role": "Project Manager",    "products": [], "learning": []},
    "Porangada, Suraj":       {"role": "Project Manager",    "products": [], "learning": []},
    # ── Solution Architects ───────────────────────────────────────────────────
    "Bell, Stuart":           {"role": "Solution Architect", "products": ["Billing"],     "learning": []},
    "DiMarco, Nicole R":      {"role": "Solution Architect", "products": ["Billing"],     "learning": []},
    "Murphy, Conor":          {"role": "Solution Architect", "products": ["Billing"],     "learning": [], "util_exempt": True},
    "Finalle-Newton, Jesse":  {"role": "Solution Architect", "products": ["Reporting"],   "learning": []},
    # ── Developers ────────────────────────────────────────────────────────────
    "Church, Jason G":        {"role": "Developer",          "products": ["All"],         "learning": [], "util_exempt": True},
    "Dunn, Steven":           {"role": "Developer",          "products": ["All"],         "learning": []},
    "Law, Brandon":           {"role": "Developer",          "products": ["Reporting"],   "learning": []},
    "Quiambao, Generalyn":    {"role": "Developer",          "products": ["All"],         "learning": []},
    # ── Consultants ───────────────────────────────────────────────────────────
    "Arestarkhov, Yaroslav":  {"role": "Consultant", "products": ["Billing", "Capture"],                                                                              "learning": []},
    "Carpen, Anamaria":       {"role": "Consultant", "products": ["Capture", "Approvals", "e-Invoicing"],                                                             "learning": []},
    "Cooke, Ellen":           {"role": "Consultant", "products": ["Billing", "Payroll"],                                                                              "learning": []},
    "Cruz, Daniel":           {"role": "Consultant", "products": ["Capture", "Approvals", "Reconcile", "SFTP Connector", "CC Statement Import"], "learning": []},
    "Dolha, Madalina":        {"role": "Consultant", "products": ["Capture", "Reconcile", "CC Statement Import", "PSP", "e-Invoicing"],                               "learning": []},
    "Gardner, Cheryll L":     {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
    "Hopkins, Chris":         {"role": "Team Lead",   "products": ["Capture", "Approvals", "Payments"],                                                                "learning": []},
    "Ickler, Georganne":      {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
    "Isberg, Eric":           {"role": "Consultant", "products": ["Reporting"],                                                                                       "learning": []},
    "Jordanova, Marija":      {"role": "Consultant", "products": ["Approvals", "Reconcile", "CC Statement Import", "PSP", "SFTP Connector"],                          "learning": []},
    "Lappin, Thomas":         {"role": "Consultant", "products": ["Payroll"],                                                                                         "learning": ["Capture", "Reconcile"]},
    "Longalong, Santiago":    {"role": "Consultant", "products": ["Capture", "Approvals", "Reconcile"],                                                               "learning": ["Billing"]},
    "Mohammad, Manaan":       {"role": "Consultant", "products": ["Capture", "Approvals", "Reconcile"],                                                               "learning": []},
    "Morris, Lisa":           {"role": "Consultant", "products": ["Payroll"],                                                                                         "learning": []},
    "NAQVI, SYED":            {"role": "Consultant", "products": ["Payroll"],                                                                                         "learning": []},
    "Olson, Austin D":        {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
    "Pallone, Daniel":        {"role": "Consultant", "products": ["Payroll"],                                                                                         "learning": []},
    "Raykova, Silvia":        {"role": "Consultant", "products": ["Capture", "Approvals", "e-Invoicing"],                                                             "learning": []},
    "Selvakumar, Sajithan":   {"role": "Consultant", "products": ["Capture", "Approvals", "Reconcile"],                                                               "learning": []},
    "Snee, Stefanie J":       {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
    "Swanson, Patti":         {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": [], "util_exempt": True},
    "Tuazon, Carol":          {"role": "Consultant", "products": ["Payroll", "Reconcile", "CC Statement Import", "PSP", "SFTP Connector"],                            "learning": []},
    "Zoric, Ivan":            {"role": "Consultant", "products": ["Capture", "Approvals", "Reconcile", "CC Statement Import", "PSP", "SFTP Connector", "Payments"],   "learning": []},
    # ── Leadership (managers only — no product delivery) ─────────────────────
    "Longi, Sameer":                  {"role": "Manager", "products": [], "learning": []},
    "Prince, Trevor":                 {"role": "Manager", "products": [], "learning": []},
    "Rusnak, Connor":                 {"role": "Manager", "products": [], "learning": []},
    # ── Leavers (historical data only — do not remove) ────────────────────────
    "Alam, Laisa":            {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
    "Centinaje, Rhodechild":  {"role": "Consultant", "products": ["Capture", "Approvals", "Reconcile", "CC Statement Import", "PSP", "SFTP Connector"],               "learning": []},
    "Chan, Joven":            {"role": "Consultant", "products": ["Capture"],                                                                                         "learning": []},
    "Cloete, Bronwyn":        {"role": "Consultant", "products": ["Capture", "Approvals"],                                                                            "learning": []},
    "Eyong, Eyong":           {"role": "Consultant", "products": ["Capture"],                                                                                         "learning": []},
    "Hamilton, Julie C":      {"role": "Consultant", "products": ["Reporting"],                                                                                       "learning": []},
    "Hernandez, Camila":      {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
    "Rushbrook, Emma C":      {"role": "Consultant", "products": ["Payroll"],                                                                                         "learning": []},
    "Strauss, John W":        {"role": "Consultant", "products": ["Billing"],                                                                                         "learning": []},
}

# Active employees (excludes no-access and leavers)
_LEAVERS = {
    "Alam, Laisa", "Chan, Joven", "Centinaje, Rhodechild", "Cloete, Bronwyn",
    "Eyong, Eyong", "Hamilton, Julie C", "Hernandez, Camila",
    "Rushbrook, Emma C", "Strauss, John W",
}

# Leaver exit dates — used for prorated available hours in team breakdown
# Format: "Name": "YYYY-MM-DD"
LEAVER_EXIT_DATES = {
    "Centinaje, Rhodechild": "2026-03-16",
    "Cloete, Bronwyn":       "2026-02-23",
    "Alam, Laisa":           None,   # date unknown
    "Chan, Joven":           None,
    "Eyong, Eyong":          None,
    "Hamilton, Julie C":     None,
    "Hernandez, Camila":     None,
    "Rushbrook, Emma C":     None,
    "Strauss, John W":       None,
}
ACTIVE_EMPLOYEES = [k for k in EMPLOYEE_ROLES if k not in NO_ACCESS and k not in _LEAVERS]

# Dropdown: consultants + manager-consultants (alphabetical)
CONSULTANT_DROPDOWN = sorted([
    e for e in ACTIVE_EMPLOYEES
    if get_role(e) in ("consultant", "manager")
])

# Dropdown: managers (all tiers)
MANAGER_DROPDOWN = sorted(
    MANAGERS_ONLY +
    [e for e in ACTIVE_EMPLOYEES if e in MANAGER_CONSULTANTS]
)

# ── Name aliases — maps DRS/NS name variants to canonical roster names ───────
# Use when source data drops middle initials or uses alternate spelling
NAME_ALIASES = {
    "church, jason":   "Church, Jason G",
    "church, jason g": "Church, Jason G",
}

def resolve_name(raw: str) -> str:
    """Return canonical roster name for a raw DRS/NS name, or the original if no alias."""
    return NAME_ALIASES.get(str(raw).strip().lower(), str(raw).strip())

def name_matches(drs_value: str, roster_name: str) -> bool:
    """Match a DRS/NS name value against a canonical roster name.
    Handles Last, First vs First Last format differences and aliases.
    """
    if not drs_value or not roster_name: return False
    v = resolve_name(str(drs_value)).strip().lower()
    r = str(roster_name).strip().lower()
    # Build match variants from roster name (Last, First format)
    parts = [p.strip() for p in r.split(",")]
    variants = {r, parts[0]}  # "last, first" and "last"
    if len(parts) == 2:
        variants.add(f"{parts[1].strip()} {parts[0]}")  # "first last"
    # Exact / boundary match
    if v in variants or any(v == nv or v.startswith(nv + " ") or v.endswith(" " + nv) for nv in variants):
        return True
    # Substring fallback for format mismatches (min 5 chars to avoid false positives)
    return any(nv in v for nv in variants if len(nv) >= 5)

# ── Column maps ───────────────────────────────────────────────────────────────
SS_COL_MAP = {
    "project name":          "project_name",
    "project id":            "project_id",
    "overall rag":           "rag",
    "start date":            "start_date",
    "go live date":          "go_live_date",
    "est. go-live date":      "go_live_date",
    "estimated go-live":      "go_live_date",
    "original go-live date":  "original_go_live_date",
    "original go live date":  "original_go_live_date",
    "forecast go-live date":  "forecast_go_live_date",
    "forecast go live date":  "forecast_go_live_date",
    "actual go-live date":    "actual_go_live_date",
    "actual go live date":    "actual_go_live_date",
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
    "legacy":                "legacy",
    "risk owner":            "risk_owner",
    "risk detail":           "risk_detail",
    "responsible for delay": "responsible_for_delay",
    "delay summary":         "delay_summary",
    "jira links":            "jira_links",
    "jira":                  "jira_links",
    "project manager":       "project_manager",
    "billing type":          "billing_type",
    "billing":               "billing_type",
}

NS_COL_MAP = {
    "employee":        "employee",
    "name":            "employee",
    "project":         "project",
    "project name":    "project",
    "project id":      "project_id",
    "billing type":    "billing_type",
    "project manager": "project_manager",
    "date":            "date",
    "hours":           "hours",
    "quantity":        "hours",
}

SFDC_COL_MAP = {
    # Exact headers from SFDC contacts export template
    "18 digit opportunity id":   "opportunity_id",
    "first name":                "first_name",
    "last name":                 "last_name",
    "primary":                   "is_primary",
    "title":                     "title",
    "email":                     "email",
    "opportunity owner":         "account_manager",
    "implementation contact exists": "impl_contact_flag",
    "contact roles":             "contact_roles",
    "opp contact role count":    "role_count",
    "partner contact":           "partner_contact",
    "opportunity owner email":   "account_manager_email",
    "account name":              "account",
    "opportunity name":          "opportunity",
    "close date":                "close_date",
    # Fallback aliases
    "opportunity":               "opportunity",
    "account":                   "account",
    "stage":                     "stage",
    "amount":                    "amount",
    "product":                   "product",
    "owner":                     "account_manager",
    "closed date":               "close_date",
    "territory":                 "territory",
    "primary title":             "title",
    "contact name":              "contact_name",
    "contact email":             "email",
    "owner email":               "account_manager_email",
}

# ── WHS phase benchmarks (days) ───────────────────────────────────────────────
PHASE_BENCHMARKS = {
    "Discovery":    14,
    "Planning":     21,
    "Build":        45,
    "UAT":          21,
    "Go-Live":      14,
    "Hypercare":    30,
    "Close-Out":    14,
}

# ── Milestone column map ──────────────────────────────────────────────────────
MILESTONE_COLS_MAP = {
    "ms_intro_email":     "Intro. Email Sent",
    "ms_config_start":    "Standard Config Start",
    "ms_enablement":      "Enablement Session",
    "ms_session1":        "Session #1",
    "ms_session2":        "Session #2",
    "ms_uat_signoff":     "UAT Signoff",
    "ms_prod_cutover":    "Prod Cutover",
    "ms_hypercare_start": "Hypercare Start",
    "ms_close_out":       "Close Out Remaining Tasks",
    "ms_transition":      "Transition to Support",
}

# ── Product keywords ──────────────────────────────────────────────────────────
PRODUCT_KEYWORDS = [
    "Capture", "Approvals", "Reconcile", "PSP", "Payments", "SFTP",
    "E-Invoicing", "eInvoicing", "CC", "Premium", "ZoneCapture",
    "ZoneApprovals", "ZoneReconcile", "ZonePayments", "ZCapture",
    "ZApprovals", "ZReconcile",
]

# ── Fixed Fee scope hours by project type ─────────────────────────────────────
# Source of truth — same table used by Utilization Report engine
DEFAULT_SCOPE = {
    "Capture":                 20,
    "Approvals":               17,
    "Reconcile":               17,
    "PSP":                     18,
    "Payments":                30,
    "AP Payment":               4,
    "Reconcile 2.0":           20,
    "CC":                       6,
    "SFTP":                    12,
    "Premium - 10":            10,
    "Premium - 20":            20,
    "E-Invoicing":             15,
    "Capture and E-Invoicing": 30,
    "Additional Subsidiary":    2,
}

def get_ff_scope(project_type: str, project_name: str = ""):
    """Return scoped hours for a project type, or None if not found / T&M.

    For ZoneApp: Premium projects, extracts hours from the project name
    e.g. 'Acme - ZA - 20 Premium Implementation' → 20
    Falls back to DEFAULT_SCOPE lookup if no number found in name.
    """
    if not project_type:
        return None
    pt = str(project_type).strip().lower()

    # Premium project type — extract hours from project name
    if "premium" in pt:
        if project_name:
            # Try IMPL10/IMPL20 SKU pattern first (from Time Item SKU in NS)
            _sku_nums = _re_constants.findall(r"IMPL(\d+)", str(project_name).upper())
            if _sku_nums:
                return float(_sku_nums[0])
            # Look for standalone 10 or 20 in the project name
            _nums = _re_constants.findall(r"(?<!\d)(10|20)(?!\d)", str(project_name))
            if _nums:
                return float(_nums[0])
        # Fall back to DEFAULT_SCOPE premium entries if no match in name
        _prem_matches = [(k, float(v)) for k, v in DEFAULT_SCOPE.items()
                         if "premium" in k.strip().lower() and k.strip().lower() in pt]
        if _prem_matches:
            return max(_prem_matches, key=lambda x: len(x[0]))[1]
        return None  # Can't determine — surface as NO SCOPE DEFINED

    matches = [(k, float(v)) for k, v in DEFAULT_SCOPE.items() if k.strip().lower() in pt]
    if not matches:
        return None
    return max(matches, key=lambda x: len(x[0]))[1]


# ── View As resolver — reads home_browse and returns (name_or_none, region_or_none, is_manager_view) ──
def resolve_view_as(consultant_name: str, home_browse: str, employee_roles: dict, employee_location: dict,
                    ps_region_map: dict, ps_region_override: dict, consultant_dropdown: list):
    """
    Returns (selected_name, region, is_group_view):
      - selected_name: single consultant name to filter to (or None for region/all)
      - region: region string if a region is selected (or None)
      - is_group_view: True if manager viewing a region or all
    """
    logged_in_role = get_role(consultant_name) if consultant_name else "consultant"
    is_mgr = logged_in_role in ("manager", "manager_only")

    if not is_mgr:
        return consultant_name, None, False

    b = str(home_browse or "").strip()
    if not b or b in ("— My own view —", "— Select —", ""):
        return consultant_name, None, False
    if b in ("👥 All team",):
        return None, None, True
    if b.startswith("── ") and b.endswith(" ──"):
        region = b[3:-3].strip()
        return None, region, True
    # Individual consultant selected
    return b, None, False


def get_region_consultants(region: str, employee_location: dict, ps_region_map: dict,
                           ps_region_override: dict, consultant_dropdown: list) -> set:
    """Return set of normalised name variants for all consultants in a region."""
    names = set()
    for n in consultant_dropdown:
        loc = employee_location.get(n, "")
        if isinstance(loc, tuple): loc = loc[0]
        nr = ps_region_override.get(n, ps_region_map.get(loc, "Other"))
        if nr == region:
            names.add(n.lower())
            parts = [p.strip() for p in n.split(",")]
            names.add(parts[0].lower())
            if len(parts) == 2:
                names.add(f"{parts[1]} {parts[0]}".lower())
    return names
