"""
PS Tools — Shared Configuration
Constants, lookup tables, and scope maps used across all reports.
"""
import streamlit as st

# ── Constants ─────────────────────────────────────────────────────────────────
NAVY     = "1e2c63"
TEAL     = "4472C4"
WHITE    = "FFFFFF"
LTGRAY   = "F2F2F2"
MID_GRAY = "BDC3C7"

TAG_COLORS = {
    "CREDITED":     "EAF9F1",
    "OVERRUN":      "FDECED",
    "PARTIAL":      "FEF9E7",
    "NON-BILLABLE": "EBEDEE",
    "FF: NO SCOPE DEFINED": "F2F2F2",
}
TAG_BADGE = {
    "CREDITED":     "🟢",
    "OVERRUN":      "🔴",
    "PARTIAL":      "🟡",
    "NON-BILLABLE": "⚫",
    "FF: NO SCOPE DEFINED": "⚪",
}

# ── Stored scope map ──────────────────────────────────────────────────────────
# ── PTO / Vacation keywords (matched against task/case column) ───────────────
PTO_KEYWORDS = ["vacation", "pto", "sick", "vacation/pto"]

# ── Employees excluded from utilization targets ──────────────────────────────
UTIL_EXEMPT_EMPLOYEES = ["swanson"]  # case-insensitive match

# ── Employee → Location lookup (drives avail hours + PS region) ──────────────
EMPLOYEE_LOCATION = {
    # ── Consultants ───────────────────────────────────────────────────────────
    "Arestarkhov, Yaroslav":  "Czech Republic",
    "Carpen, Anamaria":       "Spain",
    "Cooke, Ellen":           "Northern Ireland",
    "Cruz, Daniel":           "Manila (PH)",
    "DiMarco, Nicole R":      "USA",
    "Dolha, Madalina":        "Faroe Islands",
    "Finalle-Newton, Jesse":  "USA",
    "Gardner, Cheryll L":     "USA",
    "Hopkins, Chris":         "USA",
    "Ickler, Georganne":      "USA",
    "Isberg, Eric":           "USA",
    "Jordanova, Marija":      "North Macedonia",
    "Lappin, Thomas":         "Northern Ireland",
    "Longalong, Santiago":    "Manila (PH)",
    "Mohammad, Manaan":       "Canada",
    "Morris, Lisa":           "Sydney (AU)",
    "NAQVI, SYED":            "Canada",
    "Olson, Austin D":        "USA",
    "Pallone, Daniel":        "Sydney (AU)",
    "Raykova, Silvia":        "Netherlands",
    "Selvakumar, Sajithan":   "Canada",
    "Snee, Stefanie J":       "USA",
    "Swanson, Patti":         "UK",
    "Tuazon, Carol":          "Manila (PH)",
    "Zoric, Ivan":            "Serbia",
    # ── Solution Architects ───────────────────────────────────────────────────
    "Bell, Stuart":           "USA",
    "Murphy, Conor":          "USA",
    # ── Developers ────────────────────────────────────────────────────────────
    "Church, Jason G":        "USA",
    "Dunn, Steven":           "USA",
    "Law, Brandon":           "USA",
    "Quiambao, Generalyn":    "Manila (PH)",
    # ── Project Managers ──────────────────────────────────────────────────────
    "Barrio, Nairobi":        "USA",
    "Cadelina, Macoy":        "Manila (PH)",
    "Hughes, Madalyn":        "USA",
    "Porangada, Suraj":       "USA",
    # ── Managers only ─────────────────────────────────────────────────────────
    "Longi":                  "Sydney (AU)",
    "Prince":                 "USA",
    "Rusnak":                 "USA",
    # ── Leavers (kept for historical avail hours lookups) ─────────────────────
    "Centinaje, Rhodechild":  "Manila (PH)",
    "Chan, Joven":            "Manila (PH)",
    "Cloete, Bronwyn":        "Netherlands",  # was Capture/Approvals — left Feb 2026
    "Hamilton, Julie C":      "USA",
    "Hernandez, Camila":      "USA",
    "Rushbrook, Emma C":      "Sydney (AU)",
    "Stone, Matt":            "USA",
    "Strauss, John W":        "USA",
}

# ── PS Region overrides (employee name → region, bypasses location mapping) ──
PS_REGION_OVERRIDE = {
    "NAQVI, SYED":         "EMEA",  # Canada-based but reports into EMEA
    "Cruz, Daniel":        "NOAM",  # Manila-based but reports into NOAM
    "Quiambao, Generalyn": "NOAM",  # Manila-based but reports into NOAM
    "Cadelina, Macoy":     "NOAM",  # Manila-based but reports into NOAM
}

PS_REGION_MAP = {
    "Sydney (AU)":      "APAC",
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

DEFAULT_SCOPE = {
    "Capture":                 20,
    "Approvals":               17,
    "Reconcile":               17,
    "PSP":                     18,
    "Payments":                30,
    "Reconcile 2.0":           20,
    "CC":                       6,
    "SFTP":                    12,
    "Premium - 10":            10,
    "Premium - 20":            20,
    "E-Invoicing":             15,
    "Capture and E-Invoicing": 30,
    "Additional Subsidiary":    2,
}

# ── Available hours by region/month (2026) ────────────────────────────────────
AVAIL_HOURS = {
    "Spain":            {"2026-01":155.00,"2026-02":155.00,"2026-03":170.50,"2026-04":162.75,"2026-05":155.00,"2026-06":170.50,"2026-07":178.25,"2026-08":162.75,"2026-09":170.50,"2026-10":162.75,"2026-11":155.00,"2026-12":155.00},
    "UK":               {"2026-01":157.50,"2026-02":150.00,"2026-03":165.00,"2026-04":150.00,"2026-05":142.50,"2026-06":165.00,"2026-07":172.50,"2026-08":150.00,"2026-09":165.00,"2026-10":165.00,"2026-11":157.50,"2026-12":157.50},
    "Northern Ireland": {"2026-01":157.50,"2026-02":150.00,"2026-03":157.50,"2026-04":150.00,"2026-05":142.50,"2026-06":165.00,"2026-07":165.00,"2026-08":150.00,"2026-09":165.00,"2026-10":165.00,"2026-11":157.50,"2026-12":157.50},
    "Netherlands":      {"2026-01":168.00,"2026-02":160.00,"2026-03":176.00,"2026-04":152.00,"2026-05":144.00,"2026-06":176.00,"2026-07":184.00,"2026-08":168.00,"2026-09":176.00,"2026-10":176.00,"2026-11":168.00,"2026-12":176.00},
    "Faroe Islands":    {"2026-01":168.00,"2026-02":160.00,"2026-03":176.00,"2026-04":144.00,"2026-05":144.00,"2026-06":176.00,"2026-07":168.00,"2026-08":168.00,"2026-09":176.00,"2026-10":176.00,"2026-11":168.00,"2026-12":168.00},
    "North Macedonia":  {"2026-01":160.00,"2026-02":160.00,"2026-03":168.00,"2026-04":168.00,"2026-05":160.00,"2026-06":176.00,"2026-07":184.00,"2026-08":168.00,"2026-09":168.00,"2026-10":168.00,"2026-11":168.00,"2026-12":176.00},
    "Czech Republic":   {"2026-01":168.00,"2026-02":160.00,"2026-03":176.00,"2026-04":160.00,"2026-05":152.00,"2026-06":176.00,"2026-07":176.00,"2026-08":168.00,"2026-09":168.00,"2026-10":168.00,"2026-11":160.00,"2026-12":168.00},
    "Serbia":           {"2026-01":152.00,"2026-02":152.00,"2026-03":176.00,"2026-04":160.00,"2026-05":160.00,"2026-06":176.00,"2026-07":184.00,"2026-08":168.00,"2026-09":176.00,"2026-10":176.00,"2026-11":160.00,"2026-12":184.00},
    "Canada":           {"2026-01":168.00,"2026-02":160.00,"2026-03":176.00,"2026-04":168.00,"2026-05":160.00,"2026-06":176.00,"2026-07":176.00,"2026-08":160.00,"2026-09":160.00,"2026-10":168.00,"2026-11":160.00,"2026-12":168.00},
    "USA":              {"2026-01":160.00,"2026-02":152.00,"2026-03":176.00,"2026-04":176.00,"2026-05":160.00,"2026-06":168.00,"2026-07":176.00,"2026-08":168.00,"2026-09":168.00,"2026-10":168.00,"2026-11":152.00,"2026-12":176.00},
    "Sydney (AU)":     {"2026-01":152.00,"2026-02":152.00,"2026-03":167.20,"2026-04":144.40,"2026-05":159.60,"2026-06":159.60,"2026-07":174.80,"2026-08":152.00,"2026-09":167.20,"2026-10":159.60,"2026-11":159.60,"2026-12":159.60},
    "Manila (PH)":      {"2026-01":168.00,"2026-02":152.00,"2026-03":176.00,"2026-04":152.00,"2026-05":160.00,"2026-06":168.00,"2026-07":184.00,"2026-08":152.00,"2026-09":176.00,"2026-10":176.00,"2026-11":152.00,"2026-12":144.00},
}

# Fixed fee task keywords (Case/Task/Event column)
FF_TASKS = ["Configuration", "Enablement", "Training", "Post Go-live", "Project Management"]

# ── Currency → PS Region ───────────────────────────────────────────────────────
# Used for revenue reporting where region is derived from billing currency
CURRENCY_REGION_MAP = {
    "USD": "NOAM",
    "CAD": "NOAM",
    "AUD": "APAC",
    "NZD": "APAC",
    "GBP": "EMEA",
    "EUR": "EMEA",
}

# ── FX Rates → USD (monthly average) ─────────────────────────────────────────
# Update these monthly. Format: {"YYYY-MM": rate_to_usd}
# Rate = 1 unit of currency → USD  (e.g. 1 AUD = 0.63 USD)
# Source: use monthly average from xe.com or your finance team's published rates
FX_RATES_TO_USD = {
    "AUD": {"2026-01": 0.620, "2026-02": 0.630, "2026-03": 0.628,
            "2026-04": 0.645, "2026-05": 0.645, "2026-06": 0.645,
            "2026-07": 0.645, "2026-08": 0.645, "2026-09": 0.645,
            "2026-10": 0.645, "2026-11": 0.645, "2026-12": 0.645},
    "CAD": {"2026-01": 0.695, "2026-02": 0.700, "2026-03": 0.718,
            "2026-04": 0.720, "2026-05": 0.720, "2026-06": 0.720,
            "2026-07": 0.720, "2026-08": 0.720, "2026-09": 0.720,
            "2026-10": 0.720, "2026-11": 0.720, "2026-12": 0.720},
    "NZD": {"2026-01": 0.565, "2026-02": 0.572, "2026-03": 0.570,
            "2026-04": 0.580, "2026-05": 0.580, "2026-06": 0.580,
            "2026-07": 0.580, "2026-08": 0.580, "2026-09": 0.580,
            "2026-10": 0.580, "2026-11": 0.580, "2026-12": 0.580},
    "GBP": {"2026-01": 1.240, "2026-02": 1.252, "2026-03": 1.290,
            "2026-04": 1.295, "2026-05": 1.295, "2026-06": 1.295,
            "2026-07": 1.295, "2026-08": 1.295, "2026-09": 1.295,
            "2026-10": 1.295, "2026-11": 1.295, "2026-12": 1.295},
    "EUR": {"2026-01": 1.030, "2026-02": 1.038, "2026-03": 1.082,
            "2026-04": 1.090, "2026-05": 1.090, "2026-06": 1.090,
            "2026-07": 1.090, "2026-08": 1.090, "2026-09": 1.090,
            "2026-10": 1.090, "2026-11": 1.090, "2026-12": 1.090},
    "USD": {"2026-01": 1.000, "2026-02": 1.000, "2026-03": 1.000,
            "2026-04": 1.000, "2026-05": 1.000, "2026-06": 1.000,
            "2026-07": 1.000, "2026-08": 1.000, "2026-09": 1.000,
            "2026-10": 1.000, "2026-11": 1.000, "2026-12": 1.000},
}

def get_fx_rate(currency: str, period: str) -> float:
    """Return USD conversion rate for a currency and YYYY-MM period.
    Falls back to most recent known rate if period not found.
    """
    rates = FX_RATES_TO_USD.get(currency.upper(), {})
    if period in rates:
        return rates[period]
    # Fall back to nearest available period
    available = sorted(rates.keys())
    if not available:
        return 1.0
    return rates[available[-1]]  # use most recent

def get_avail_hours(region, period):
    """Look up available hours for a region/period. Returns None if not found."""
    region_clean = str(region).strip()
    # Try exact match first, then case-insensitive
    for r, months in AVAIL_HOURS.items():
        if r.lower() == region_clean.lower():
            return months.get(str(period), None)
    return None
