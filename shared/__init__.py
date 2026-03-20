"""
PS Tools — Shared Package
Import everything from one place.
"""
from shared.config import (
    NAVY, TEAL, WHITE, LTGRAY, MID_GRAY,
    TAG_COLORS, TAG_BADGE,
    PTO_KEYWORDS, UTIL_EXEMPT_EMPLOYEES,
    EMPLOYEE_LOCATION, PS_REGION_OVERRIDE, PS_REGION_MAP,
    AVAIL_HOURS, DEFAULT_SCOPE, FF_TASKS,
    get_avail_hours,
)
from shared.constants import (
    MANAGERS_ONLY, MANAGER_CONSULTANTS, NO_ACCESS,
    EMPLOYEE_ROLES, ACTIVE_EMPLOYEES,
    CONSULTANT_DROPDOWN, MANAGER_DROPDOWN,
    SS_COL_MAP, NS_COL_MAP, SFDC_COL_MAP,
    PHASE_BENCHMARKS, MILESTONE_COLS_MAP,
    PRODUCT_KEYWORDS,
    get_role, is_manager, is_consultant,
)
from shared.loaders import (
    fuzzy_match_sfdc,
    load_sfdc, load_drs, load_ns_time,
    calc_days_inactive, calc_last_milestone,
    normalise_product_name, suggest_tier_from_days,
)
from shared.template_utils import (
    TEMPLATES,
    suggest_tier, fill_template,
    highlight_placeholders, extract_placeholders,
)
