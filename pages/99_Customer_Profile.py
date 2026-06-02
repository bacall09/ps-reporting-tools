"""
PS Tools — Customer Profile
Gong handover intelligence fetched live from the RevOps PS Handover site.
Supports multiple opportunities per customer (multi-select), fuzzy customer
name matching, merged intelligence view, and an AI Q&A stub (Phase 2).
"""
import streamlit as st
import pandas as pd
import re
import io
from datetime import date

st.session_state["current_page"] = "Customer Profile"

from shared.constants import (
    EMPLOYEE_ROLES, get_role, is_manager, name_matches, get_ff_scope,
)

# ── Auth ──────────────────────────────────────────────────────────────────────
_session_name = st.session_state.get("consultant_name", "")
if not _session_name:
    st.warning("Sign in on the Home page to use Customer Profile.")
    st.stop()

role = get_role(_session_name)

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;700&display=swap" rel="stylesheet">
<style>
    html,body,[class*="css"]{font-family:'Manrope',sans-serif!important}
    h1,h2,h3,h4,p,label,button{font-family:'Manrope',sans-serif!important}
    .section-label{font-size:13px;font-weight:700;text-transform:uppercase;
                   letter-spacing:.8px;color:#4472C4;margin-bottom:8px}
    .divider{border:none;border-top:1px solid rgba(128,128,128,.15);margin:20px 0}
    .pill{display:inline-block;font-size:11px;font-weight:700;padding:2px 9px;
          border-radius:10px;letter-spacing:.5px;margin:2px 3px 2px 0}
    .pill-teal{background:rgba(8,169,183,.15);color:#08A9B7}
    .pill-blue{background:rgba(59,158,255,.12);color:#3B9EFF}
    .pill-amber{background:rgba(243,156,18,.15);color:#D68910}
    .pill-red{background:rgba(192,57,43,.15);color:#C0392B}
    .pill-green{background:rgba(39,174,96,.12);color:#1A7A4A}
    .pill-grey{background:rgba(128,128,128,.12);color:inherit;opacity:.75;
               border:0.5px solid rgba(128,128,128,.2)}
    .cp-card{border:1px solid rgba(128,128,128,.18);border-radius:10px;
             padding:16px 18px;margin-bottom:12px}
    .cp-bullet{font-size:13px;padding:5px 0 5px 16px;position:relative;
               line-height:1.6;border-bottom:0.5px solid rgba(128,128,128,.1)}
    .cp-bullet:last-child{border-bottom:none}
    .cp-bullet::before{content:'·';position:absolute;left:4px;
                       color:rgba(128,128,128,.5);font-size:16px;top:3px}
    .cp-flag{font-size:11px;font-weight:700;padding:1px 7px;border-radius:8px;
             background:rgba(243,156,18,.15);color:#D68910;margin-left:6px;
             vertical-align:middle}
    .risk-badge{font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px;
                white-space:nowrap;flex-shrink:0;margin-top:3px}
    .risk-tech{background:rgba(59,158,255,.12);color:#3B9EFF}
    .risk-exp{background:rgba(243,156,18,.15);color:#D68910}
    .risk-org{background:rgba(192,57,43,.15);color:#C0392B}
    .risk-timeline{background:rgba(128,128,128,.15);color:inherit;opacity:.8}
    .risk-high{background:rgba(192,57,43,.15);color:#C0392B}
    .risk-med{background:rgba(243,156,18,.12);color:#D68910}
    .commit-icon{font-size:14px;flex-shrink:0;margin-top:2px;line-height:1}
    .stakeholder-row{display:flex;align-items:center;gap:10px;
                     padding:8px 0;border-bottom:0.5px solid rgba(128,128,128,.1)}
    .stakeholder-row:last-child{border-bottom:none}
    .avatar{width:30px;height:30px;border-radius:50%;
            background:rgba(59,158,255,.12);color:#3B9EFF;
            display:inline-flex;align-items:center;justify-content:center;
            font-size:10px;font-weight:700;flex-shrink:0}
    .avatar-int{background:rgba(8,169,183,.15);color:#08A9B7}
    .info-gap-row{font-size:13px;padding:5px 0 5px 18px;position:relative;
                  line-height:1.6;border-bottom:0.5px solid rgba(128,128,128,.1);
                  color:rgba(128,128,128,.9)}
    .info-gap-row:last-child{border-bottom:none}
    .info-gap-row::before{content:'?';position:absolute;left:3px;top:5px;
                          font-size:10px;font-weight:700;color:#D68910}
    .ai-panel{border:1px solid rgba(59,158,255,.25);border-radius:10px;
              padding:16px 18px;margin-bottom:16px;
              background:rgba(59,158,255,.03)}
    .ai-stub{border:1px dashed rgba(128,128,128,.25);border-radius:10px;
             padding:16px 18px;margin-top:8px;opacity:.65}
    .usage-stub{border:1px dashed rgba(128,128,128,.2);border-radius:10px;
                padding:24px;text-align:center;color:rgba(128,128,128,.6);
                font-size:13px}
    .req-row{font-size:13px;padding:5px 0 5px 16px;position:relative;
             line-height:1.6;border-bottom:0.5px solid rgba(128,128,128,.1)}
    .req-row:last-child{border-bottom:none}
    .req-row::before{content:'✓';position:absolute;left:2px;top:5px;
                     font-size:10px;color:rgba(128,128,128,.4)}
    .req-nice::before{content:'○';color:rgba(128,128,128,.35)}
    .no-data-msg{text-align:center;padding:36px;opacity:.4;font-size:14px}
    .opp-source-badge{font-size:9px;font-weight:700;padding:1px 6px;border-radius:8px;
                      background:rgba(59,158,255,.12);color:#3B9EFF;margin-left:6px;
                      vertical-align:middle;letter-spacing:.3px}
    .opp-chip{display:inline-flex;align-items:center;gap:5px;font-size:11px;
              font-weight:600;padding:3px 10px;border-radius:12px;margin:2px 4px 2px 0;
              border:1px solid rgba(68,114,196,.3);background:rgba(68,114,196,.08);
              color:#4472C4}
</style>
""", unsafe_allow_html=True)

# ── Zone SVG watermark ────────────────────────────────────────────────────────
_zone_svg = """"""

# ── Hero banner ───────────────────────────────────────────────────────────────
_hero = st.empty()
_hero.markdown(
    f"<div style='background:linear-gradient(135deg,#1a56db 0%,#050D1F 55%,#050D1F 100%);"
    f"padding:32px 40px 28px;border-radius:10px;margin-bottom:24px;font-family:Manrope,sans-serif;"
    f"position:relative;overflow:hidden'>{_zone_svg}"
    f"<div style='font-size:13px;font-weight:700;letter-spacing:2.5px;text-transform:uppercase;"
    f"color:#3B9EFF;margin-bottom:10px'>Professional Services · Tools</div>"
    f"<h1 style='color:white;margin:0;font-size:28px;font-family:Manrope,sans-serif;font-weight:800'>"
    f"Customer Profile</h1>"
    f"<p style='color:rgba(255,255,255,0.45);margin:6px 0 0;font-size:14px;"
    f"font-family:Manrope,sans-serif;max-width:560px'>"
    f"Gong handover intelligence — loaded live from RevOps. Select a customer to begin.</p>"
    f"</div>",
    unsafe_allow_html=True
)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — REVOPS HTML PARSER
# ══════════════════════════════════════════════════════════════════════════════

HANDOVER_BASE_URL = "https://ps-handover-82020413660.us-central1.run.app"
_FUZZY_THRESHOLD = 82  # minimum score to auto-match customer name


def _initials(name: str) -> str:
    parts = str(name).strip().split()
    if not parts: return "?"
    if len(parts) == 1: return parts[0][:2].upper()
    return (parts[0][0] + parts[-1][0]).upper()


def _strip_flag(text: str) -> tuple:
    flagged = "⚠️" in text or "🔴" in text
    clean = text.replace("⚠️", "").replace("🔴", "").strip().rstrip(".")
    return clean, flagged


def _get_section_body(soup, title_keyword: str):
    """Find a card-section whose section-title contains title_keyword."""
    from bs4 import BeautifulSoup as _BS
    for sec in soup.find_all(class_='card-section'):
        title_el = sec.find(class_='section-title')
        if not title_el:
            continue
        for sn in title_el.find_all(class_='sec-num'):
            sn.decompose()
        t = title_el.get_text(strip=True).lower()
        if title_keyword.lower() in t:
            return sec.find(class_='section-body')
    return None


def parse_handover_html(html_content: str, customer_name: str) -> dict:
    """Parse RevOps PS Handover HTML into the standard intelligence dict."""
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html_content, 'html.parser')

    # ── Hero ──────────────────────────────────────────────────────────────────
    account_el  = soup.find(class_='hero-account')
    account_name = account_el.get_text(strip=True) if account_el else customer_name
    opp_el      = soup.find(class_='hero-opp')
    opp_name    = opp_el.get_text(strip=True) if opp_el else ""
    sf_link_el  = soup.find(class_='hero-sf-link')
    sf_link     = sf_link_el['href'] if sf_link_el and sf_link_el.get('href') else ""
    opp_id_m    = re.search(r'/([A-Za-z0-9]{15,18})(?:/|$)', sf_link)
    sf_opp_id   = opp_id_m.group(1) if opp_id_m else ""
    meta        = {}
    for item in soup.find_all(class_='hero-meta-item'):
        lbl = item.find(class_='hero-meta-label')
        val = item.find(class_='hero-meta-value')
        if lbl and val:
            meta[lbl.get_text(strip=True)] = val.get_text(strip=True)

    # ── Summary ───────────────────────────────────────────────────────────────
    summary_el = soup.find(class_='summary-text')
    summary    = summary_el.get_text(strip=True) if summary_el else ""

    # ── Pain Points ───────────────────────────────────────────────────────────
    pain_points = []
    pp_body = _get_section_body(soup, 'pain point')
    if pp_body:
        for li in pp_body.find_all('li'):
            t = li.get_text(strip=True)
            if t: pain_points.append(_strip_flag(t))
        if not pain_points:
            for p in pp_body.find_all('p'):
                t = p.get_text(strip=True)
                if t: pain_points.append(_strip_flag(t))

    # ── Requirements ──────────────────────────────────────────────────────────
    must_haves, nice_haves = [], []
    req_body = _get_section_body(soup, 'requirement')
    if req_body:
        current_list = must_haves
        for el in req_body.children:
            if not hasattr(el, 'get_text'):
                continue
            t = el.get_text(strip=True)
            if re.search(r'nice.to.have', t, re.IGNORECASE):
                current_list = nice_haves
                continue
            if el.name == 'ul':
                for li in el.find_all('li'):
                    lt = li.get_text(strip=True)
                    if lt and not re.search(r'must.have', lt, re.IGNORECASE):
                        current_list.append(lt)
            elif el.name == 'p' and t and not re.search(r'must.have', t, re.IGNORECASE):
                current_list.append(t)

    # ── Use Cases ─────────────────────────────────────────────────────────────
    use_cases = []
    uc_body = _get_section_body(soup, 'use case')
    if uc_body:
        for uc in uc_body.find_all(class_='use-case'):
            title_el = uc.find(class_='uc-title')
            meta_el  = uc.find(class_='uc-meta')
            if title_el:
                use_cases.append({
                    'title': title_el.get_text(strip=True),
                    'meta':  meta_el.get_text(' | ', strip=True) if meta_el else '',
                })

    # ── Technical Environment ─────────────────────────────────────────────────
    tech_env = []
    te_body = _get_section_body(soup, 'technical')
    if te_body:
        for li in te_body.find_all('li'):
            t = li.get_text(strip=True)
            if t: tech_env.append(_strip_flag(t))
        if not tech_env:
            for p in te_body.find_all('p'):
                t = p.get_text(strip=True)
                if t: tech_env.append(_strip_flag(t))

    # ── Timeline ──────────────────────────────────────────────────────────────
    timeline = []
    tl_body = _get_section_body(soup, 'delivery')
    if tl_body:
        for li in tl_body.find_all('li'):
            t = li.get_text(strip=True)
            if t: timeline.append(_strip_flag(t))

    # ── Stakeholders ──────────────────────────────────────────────────────────
    stakeholders = []
    sk_body = _get_section_body(soup, 'stakeholder')
    if sk_body:
        for tr in sk_body.find_all('tr'):
            tds = tr.find_all('td')
            if len(tds) >= 2:
                name = tds[0].get_text(strip=True)
                role_t = tds[1].get_text(strip=True)
                relevance = tds[2].get_text(strip=True) if len(tds) > 2 else ''
                influence = tds[3].get_text(strip=True) if len(tds) > 3 else ''
                sentiment = tds[4].get_text(strip=True) if len(tds) > 4 else ''
                if name and name.lower() not in ('name', 'role', 'relevance', 'influence', 'sentiment'):
                    stakeholders.append({
                        'name':      name,
                        'title':     role_t,
                        'relevance': relevance,
                        'influence': influence,
                        'sentiment': sentiment,
                        'internal':  False,
                        'email':     '',
                        'role_note': role_t,
                    })

    # ── Sales Commitment Flags ────────────────────────────────────────────────
    commitments = []
    cf_body = _get_section_body(soup, 'commitment')
    if cf_body:
        for tr in cf_body.find_all('tr'):
            tds = tr.find_all('td')
            if len(tds) >= 2:
                pills = tds[0].find_all(class_='pill') if tds else []
                status_text = pills[0].get_text(strip=True) if pills else tds[0].get_text(strip=True)
                commit_text = tds[1].get_text(strip=True) if len(tds) > 1 else ''
                if commit_text and commit_text.lower() not in ('commitment', 'flag', 'context'):
                    if '🔴' in status_text or 'risk' in status_text.lower():
                        status = 'risk'
                    elif '⚠️' in status_text or 'review' in status_text.lower():
                        status = 'review'
                    else:
                        status = 'aligned'
                    commitments.append({'status': status, 'text': commit_text})

    # ── Risk Signals ──────────────────────────────────────────────────────────
    risks = []
    rs_body = _get_section_body(soup, 'risk signal')
    if rs_body:
        for ri in rs_body.find_all(class_='risk-item'):
            title_el = ri.find(class_='risk-title')
            body_el  = ri.find(class_='risk-body')
            pills    = ri.find_all(class_='pill')
            severity = pills[0].get_text(strip=True) if pills else ''
            title_t  = title_el.get_text(strip=True) if title_el else ''
            body_t   = body_el.get_text(strip=True) if body_el else ''
            if title_t or body_t:
                cat = title_t.lower()
                if 'technical' in cat or 'tech' in cat:
                    badge = 'tech'
                elif 'expectation' in cat or 'misalign' in cat:
                    badge = 'exp'
                elif 'org' in cat or 'readiness' in cat or 'timeline' in cat:
                    badge = 'org'
                else:
                    badge = 'tech'
                risks.append({
                    'badge':    badge,
                    'category': title_t,
                    'text':     body_t,
                    'severity': severity,
                    'flagged':  severity.upper() == 'HIGH',
                })

    # ── Assumptions & Information Gaps ───────────────────────────────────────
    info_gaps = []
    ig_body = _get_section_body(soup, 'assumption')
    if ig_body:
        for gr in ig_body.find_all(class_='gap-row'):
            t = gr.get_text(' ', strip=True)
            if t: info_gaps.append(t)

    # ── Competitors ───────────────────────────────────────────────────────────
    competitors = []
    co_body = _get_section_body(soup, 'competitor')
    if co_body:
        for ci in co_body.find_all(class_='comp-item'):
            t = ci.get_text(' | ', strip=True)
            if t: competitors.append(t)

    # ── Products ──────────────────────────────────────────────────────────────
    products_raw = meta.get('Products', '')
    products = [p.strip() for p in products_raw.split(',') if p.strip()] if products_raw else []

    return {
        'customer':     account_name,
        'opp_name':     opp_name,
        'opp_link':     sf_link,
        'sf_opp_id':    sf_opp_id,
        'arr':          meta.get('ARR', ''),
        'close_date':   meta.get('Closed', ''),
        'ae':           meta.get('AE', ''),
        'region':       meta.get('Region', ''),
        'products':     products,
        'summary':      summary,
        'pain_points':  pain_points,
        'requirements': {'must': must_haves, 'nice': nice_haves},
        'use_cases':    use_cases,
        'tech_env':     tech_env,
        'timeline':     timeline,
        'stakeholders': stakeholders,
        'commitments':  commitments,
        'risks':        risks,
        'info_gaps':    info_gaps,
        'competitors':  competitors,
        'data_used':    {'AE': meta.get('AE', ''), 'ARR': meta.get('ARR', '')},
        # raw text truncated for AI context use — enough for Q&A, not the full DOM
        'raw':          soup.get_text(separator='\n', strip=True)[:12000],
        'source':       'revops_fetch',
    }


def _fetch_handover(opp_id: str) -> tuple:
    """
    Fetch HTML from the RevOps PS Handover site for a given opp_id.
    Returns (html_content, error_message).
    """
    try:
        import requests
        url = f"{HANDOVER_BASE_URL}/?opp_id={opp_id}"
        resp = requests.get(url, timeout=15)
        if resp.status_code == 200:
            return resp.text, ""
        return "", f"RevOps site returned HTTP {resp.status_code} for opp {opp_id}."
    except Exception as e:
        return "", f"Could not reach RevOps site: {e}"


def _extract_opp_ids_from_sfdc(df_sfdc: pd.DataFrame, customer_name: str) -> list:
    """
    Find all Salesforce opportunities in SFDC contacts data matching customer_name.
    Uses fuzzy matching. Returns list of dicts: {opp_id, opp_name, account, score}.
    """
    from rapidfuzz import fuzz as _fuzz, process as _proc

    if df_sfdc is None or df_sfdc.empty:
        return []

    # Identify the account name column
    acct_col = next((c for c in ('account', 'account_name', 'company') if c in df_sfdc.columns), None)
    opp_id_col = next((c for c in ('opportunity_id', 'opp_id', 'sf_opp_id') if c in df_sfdc.columns), None)
    opp_name_col = next((c for c in ('opportunity_name', 'opp_name', 'opportunity') if c in df_sfdc.columns), None)

    if not acct_col or not opp_id_col:
        return []

    # Deduplicate to unique (opp_id, account) pairs
    dedup_cols = [opp_id_col, acct_col] + ([opp_name_col] if opp_name_col else [])
    unique_opps = df_sfdc[dedup_cols].drop_duplicates(subset=[opp_id_col]).dropna(subset=[opp_id_col])

    results = []
    for _, row in unique_opps.iterrows():
        acct = str(row[acct_col] or '').strip()
        if not acct:
            continue
        score = _fuzz.token_set_ratio(customer_name.lower(), acct.lower())
        if score >= _FUZZY_THRESHOLD:
            results.append({
                'opp_id':   str(row[opp_id_col]).strip(),
                'opp_name': str(row[opp_name_col]).strip() if opp_name_col else '',
                'account':  acct,
                'score':    score,
            })

    # Sort: exact / near-exact first, then by opp_name
    results.sort(key=lambda x: (-x['score'], x['opp_name']))
    return results


def _merge_docs(docs: list) -> dict:
    """
    Merge multiple parsed handover dicts into a single combined view.
    Lists are unioned (pain points, risks, etc.). Stakeholders are deduplicated by name.
    Each item tagged with its source opp label where relevant.
    """
    if not docs:
        return {}
    if len(docs) == 1:
        return docs[0]

    def _merge_list(key):
        seen, merged = set(), []
        for doc in docs:
            opp_label = doc.get('opp_name', '')[:30]
            for item in doc.get(key, []):
                if isinstance(item, tuple):
                    text, flag = item
                else:
                    text, flag = str(item), False
                if text not in seen:
                    seen.add(text)
                    merged.append((f"[{opp_label}] {text}" if opp_label else text, flag))
        return merged

    def _merge_stakeholders():
        from rapidfuzz import fuzz as _fuzz
        seen_names, merged = [], []
        for doc in docs:
            for s in doc.get('stakeholders', []):
                match = any(_fuzz.token_set_ratio(s['name'].lower(), n.lower()) >= 88 for n in seen_names)
                if not match:
                    seen_names.append(s['name'])
                    merged.append(s)
        return merged

    def _merge_requirements():
        must_seen, nice_seen = set(), set()
        must_all, nice_all = [], []
        for doc in docs:
            req = doc.get('requirements', {})
            for t in (req.get('must', []) if isinstance(req, dict) else req):
                if t not in must_seen:
                    must_seen.add(t)
                    must_all.append(t)
            for t in (req.get('nice', []) if isinstance(req, dict) else []):
                if t not in nice_seen:
                    nice_seen.add(t)
                    nice_all.append(t)
        return {'must': must_all, 'nice': nice_all}

    # Products: union all
    all_products = []
    for doc in docs:
        for p in doc.get('products', []):
            if p not in all_products:
                all_products.append(p)

    # Use the first doc's summary/meta as the primary; append others
    primary = docs[0]
    summaries = [d['summary'] for d in docs if d.get('summary')]
    combined_summary = "\n\n".join(summaries) if len(summaries) > 1 else (summaries[0] if summaries else "")

    return {
        'customer':     primary['customer'],
        'opp_name':     ' + '.join(d['opp_name'] for d in docs if d.get('opp_name')),
        'opp_link':     primary.get('opp_link', ''),
        'sf_opp_id':    primary.get('sf_opp_id', ''),
        'arr':          primary.get('arr', ''),
        'close_date':   primary.get('close_date', ''),
        'ae':           primary.get('ae', ''),
        'region':       primary.get('region', ''),
        'products':     all_products,
        'summary':      combined_summary,
        'pain_points':  _merge_list('pain_points'),
        'requirements': _merge_requirements(),
        'use_cases':    [uc for doc in docs for uc in doc.get('use_cases', [])],
        'tech_env':     _merge_list('tech_env'),
        'timeline':     _merge_list('timeline'),
        'stakeholders': _merge_stakeholders(),
        'commitments':  [c for doc in docs for c in doc.get('commitments', [])],
        'risks':        [r for doc in docs for r in doc.get('risks', [])],
        'info_gaps':    [g for doc in docs for g in doc.get('info_gaps', [])],
        'competitors':  list({c for doc in docs for c in doc.get('competitors', [])}),
        'data_used':    primary.get('data_used', {}),
        'raw':          '\n\n---\n\n'.join(d.get('raw', '') for d in docs)[:16000],
        'source':       'revops_fetch',
        '_source_docs': docs,  # keep originals for per-opp views
    }


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — DATA + SESSION INIT
# ══════════════════════════════════════════════════════════════════════════════

df_drs  = st.session_state.get("df_drs")
df_sfdc = st.session_state.get("df_sfdc")
df_ns   = st.session_state.get("df_ns")

# Session cache: {customer_name: {opp_id: parsed_doc}}
if "cp_handover_cache" not in st.session_state:
    st.session_state["cp_handover_cache"] = {}

# ── Customer name extraction from DRS project_name ───────────────────────────
_PC = ["ZEP","ZBilling","ZPayroll","ZoneCapture","ZoneApprovals","ZoneReconcile",
       "ZA","ZC","ZR","ZB","ZP"]
_PW = ["Payroll","Billing","Capture","Approvals","Reconcile","Implementation",
       "Optimization","Migration","Integration","Training","Support"]

def _extract_customer_name(project_name):
    import re as _re
    n = str(project_name).strip()
    m = _re.match(r'^(.+?)\s*-\s*[A-Z]{1,4}\s*-\s*.+$', n)
    if m: return m.group(1).strip()
    _pc_pat = '|'.join(_PC)
    m = _re.match(r'^(.+?)\s*-\s*(?:' + _pc_pat + r')(?:\s|$|-)', n, _re.IGNORECASE)
    if m: return m.group(1).strip()
    _pw_pat = '|'.join(_PW)
    m = _re.match(r'^(.+?)\s*-\s*(?:' + _pw_pat + r').+$', n, _re.IGNORECASE)
    if m: return m.group(1).strip()
    for code in sorted(_PC, key=len, reverse=True):
        m = _re.search(r'\s+' + _re.escape(code) + r'(?:\s|$|-)', n, _re.IGNORECASE)
        if m and m.start() > 2: return n[:m.start()].strip()
    for word in _PW:
        m = _re.search(r'\s+' + _re.escape(word) + r'(?:\s|$)', n, _re.IGNORECASE)
        if m and m.start() > 3: return n[:m.start()].strip().rstrip('-').strip()
    return n

drs_customers = []
if df_drs is not None and not df_drs.empty:
    if "account" in df_drs.columns:
        _raw = df_drs["account"].dropna()
        drs_customers = sorted(set(str(r).strip() for r in _raw if str(r).strip()))
    elif "project_name" in df_drs.columns:
        _raw_extracted = set(
            _extract_customer_name(r)
            for r in df_drs["project_name"].dropna() if str(r).strip()
        )
        import re as _re2
        _LEGAL = r'\b(limited|ltd|group|inc|gmbh|llc|corp|corporation|pty|plc|bv|ag)\b'
        def _is_legal_ext(short, long):
            extra = long[len(short):].strip()
            return bool(_re2.search(_LEGAL, extra, _re2.IGNORECASE))
        _consolidated = set()
        for name in sorted(_raw_extracted, key=len, reverse=True):
            _skip = False
            for ex in _consolidated:
                if ex.lower().startswith(name.lower()) and ex != name and _is_legal_ext(name, ex):
                    _skip = True
                    break
            if not _skip:
                _consolidated = {e for e in _consolidated if not (
                    e.lower().startswith(name.lower()) and e != name and not _is_legal_ext(name, e)
                )}
                _consolidated.add(name)
        drs_customers = sorted(_consolidated)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — CUSTOMER SELECTOR
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div style='background:var(--color-background-secondary,rgba(59,158,255,.05));
            border-left:4px solid #3B9EFF;border-radius:6px;
            padding:16px 20px;margin:0 0 20px;font-family:Manrope,sans-serif'>
    <div style='font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;
                color:#3B9EFF;margin-bottom:10px'>How it works</div>
    <div style='display:flex;gap:32px;flex-wrap:wrap'>
        <div style='flex:1;min-width:200px'>
            <span style='background:#1E2C63;color:#fff;font-size:10px;font-weight:700;
                         padding:2px 8px;border-radius:10px;letter-spacing:1px'>LIVE · AUTO-FETCH</span>
            <p style='margin:8px 0 0;font-size:13px;color:inherit;line-height:1.6'>
                Select a customer — PSPT matches their Salesforce opportunities and loads
                handover intelligence directly from the RevOps PS Handover site. No file uploads needed.
            </p>
        </div>
        <div style='flex:1;min-width:200px'>
            <span style='background:rgba(59,158,255,.15);color:#3B9EFF;font-size:10px;font-weight:700;
                         padding:2px 8px;border-radius:10px;letter-spacing:1px;
                         border:1px solid rgba(59,158,255,.4)'>PHASE 2 · COMING</span>
            <p style='margin:8px 0 0;font-size:13px;color:inherit;opacity:.65;line-height:1.6'>
                <strong>AI Q&amp;A powered by Claude</strong> — ask natural language questions
                about any customer using their Gong intelligence, requirements, and NetSuite usage data.
            </p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-label">Customer</div>', unsafe_allow_html=True)

_autofill = st.session_state.pop("_cp_autofill_customer", "")

if drs_customers:
    customer_options = ["— Select a customer —"] + drs_customers + ["+ Enter manually"]
    _default_idx = 0
    if _autofill and _autofill in drs_customers:
        _default_idx = drs_customers.index(_autofill) + 1
    _sel = st.selectbox("Customer", customer_options, index=_default_idx,
                        label_visibility="collapsed")
    if _sel == "+ Enter manually":
        selected_customer = st.text_input("Customer name", value=_autofill,
                                          placeholder="e.g. FastMarkets Global Limited",
                                          label_visibility="collapsed")
    elif _sel == "— Select a customer —":
        selected_customer = _autofill
    else:
        selected_customer = _sel
else:
    st.markdown('<div style="font-size:12px;opacity:.5;margin-bottom:6px">DRS not loaded — enter name manually</div>',
                unsafe_allow_html=True)
    selected_customer = st.text_input("Customer name", value=_autofill,
                                      placeholder="e.g. FastMarkets Global Limited",
                                      label_visibility="collapsed")

if not selected_customer:
    st.markdown("""
    <div class="no-data-msg">
        Select a customer above to view their profile.
    </div>""", unsafe_allow_html=True)
    st.stop()

st.markdown("<hr class='divider'>", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — OPPORTUNITY DISCOVERY + MULTI-SELECT
# ══════════════════════════════════════════════════════════════════════════════

_cache_key = selected_customer
if _cache_key not in st.session_state["cp_handover_cache"]:
    st.session_state["cp_handover_cache"][_cache_key] = {}

_customer_cache = st.session_state["cp_handover_cache"][_cache_key]

# Find matching opps from SFDC
_sfdc_opps = _extract_opp_ids_from_sfdc(df_sfdc, selected_customer)

# ── Fallback: allow manual opp ID entry if SFDC not loaded ───────────────────
_sfdc_loaded = df_sfdc is not None and not df_sfdc.empty
_sfdc_has_opps = bool(_sfdc_opps)

if not _sfdc_loaded:
    st.info("SFDC Contacts not loaded — enter a Salesforce opportunity ID manually to load handover data.")
    _manual_opp_id = st.text_input("Salesforce opportunity ID",
                                   placeholder="e.g. 006Uh00000ic95lIAA",
                                   label_visibility="collapsed")
    if _manual_opp_id and _manual_opp_id.strip():
        _sfdc_opps = [{'opp_id': _manual_opp_id.strip(),
                       'opp_name': 'Manual entry',
                       'account': selected_customer,
                       'score': 100}]
    else:
        _sfdc_opps = []

elif not _sfdc_has_opps:
    from rapidfuzz import fuzz as _rfuzz
    # Show best near-miss if score close but below threshold
    acct_col = next((c for c in ('account', 'account_name', 'company') if c in df_sfdc.columns), None)
    if acct_col:
        _candidates = df_sfdc[acct_col].dropna().unique()
        _scores = [(str(c), _rfuzz.token_set_ratio(selected_customer.lower(), str(c).lower()))
                   for c in _candidates]
        _scores.sort(key=lambda x: -x[1])
        if _scores and _scores[0][1] >= 60:
            st.warning(
                f"No SFDC opportunities found for **{selected_customer}**. "
                f"Closest match in SFDC: **{_scores[0][0]}** (similarity {_scores[0][1]}%). "
                f"If that's the same customer, select them from the dropdown above, or enter an opp ID manually."
            )
        else:
            st.warning(f"No SFDC opportunities found for **{selected_customer}**. "
                       f"Check the customer name or enter an opp ID manually.")
    _manual_opp_id = st.text_input("Or enter opp ID manually",
                                   placeholder="e.g. 006Uh00000ic95lIAA",
                                   label_visibility="collapsed")
    if _manual_opp_id and _manual_opp_id.strip():
        _sfdc_opps = [{'opp_id': _manual_opp_id.strip(),
                       'opp_name': 'Manual entry',
                       'account': selected_customer,
                       'score': 100}]

# ── Opp picker ────────────────────────────────────────────────────────────────
if not _sfdc_opps:
    st.markdown(f"""
    <div style='margin-bottom:16px'>
        <span style='font-size:20px;font-weight:700'>{selected_customer}</span>
    </div>
    <div class='no-data-msg'>
        No Salesforce opportunities found.<br>
        <span style='font-size:12px;opacity:.6'>
            Load SFDC Contacts on the Home page to enable auto-discovery.
        </span>
    </div>""", unsafe_allow_html=True)
    st.stop()

# Build opp labels for display
def _opp_display_label(opp: dict) -> str:
    name = opp.get('opp_name', '') or opp['opp_id']
    # If already fetched, enrich with products
    cached = _customer_cache.get(opp['opp_id'])
    if cached and cached.get('products'):
        products_str = ' · '.join(cached['products'][:3])
        return f"{name}  [{products_str}]"
    return name

_opp_labels = [_opp_display_label(o) for o in _sfdc_opps]

st.markdown('<div class="section-label">Opportunities</div>', unsafe_allow_html=True)

if len(_sfdc_opps) == 1:
    # Single opp — show it but allow deselect
    _selected_labels = st.multiselect(
        "Select opportunities to load",
        options=_opp_labels,
        default=_opp_labels,
        label_visibility="collapsed",
    )
else:
    _selected_labels = st.multiselect(
        "Select opportunities to load (multi-select to combine)",
        options=_opp_labels,
        default=_opp_labels[:1],  # default to most recent
        label_visibility="collapsed",
        help="Select multiple to merge all intelligence into a single view."
    )

if not _selected_labels:
    st.markdown('<div class="no-data-msg">Select at least one opportunity above.</div>',
                unsafe_allow_html=True)
    st.stop()

_selected_opps = [_sfdc_opps[_opp_labels.index(lbl)] for lbl in _selected_labels]


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — FETCH + PARSE (with 403-aware upload fallback)
# ══════════════════════════════════════════════════════════════════════════════

# Track which opps need a manual upload because server-side fetch was blocked
if "cp_upload_fallback" not in st.session_state:
    st.session_state["cp_upload_fallback"] = {}  # {opp_id: True}

_fetch_errors   = []
_fetched_docs   = []
_blocked_opps   = []  # opps that got 403 and need upload

for opp in _selected_opps:
    oid = opp['opp_id']

    # Already cached — validate it has real content before using
    if oid in _customer_cache:
        _cached = _customer_cache[oid]
        _has_content = bool(
            _cached.get('summary') or
            _cached.get('pain_points') or
            _cached.get('requirements', {}).get('must') or
            _cached.get('risks') or
            _cached.get('stakeholders')
        )
        if _has_content:
            _fetched_docs.append(_cached)
            continue
        else:
            # Stale/empty cache entry — clear it and re-fetch
            del _customer_cache[oid]
            st.session_state["cp_upload_fallback"].pop(oid, None)

    # Already known-blocked — skip fetch, go straight to upload
    if st.session_state["cp_upload_fallback"].get(oid):
        _blocked_opps.append(opp)
        continue

    # Attempt live fetch
    with st.spinner(f"Loading handover for {opp.get('opp_name', oid)}…"):
        html_content, err = _fetch_handover(oid)

    # Debug info — remove once fetch behaviour is confirmed
    _is_blocked = (
        bool(err) and ("403" in err or "Host not in allowlist" in err)
    ) or (
        bool(html_content) and ("Host not in allowlist" in html_content or len(html_content) < 100)
    )

    if _is_blocked:
        # Blocked — flag for upload fallback, don't show as error
        st.session_state["cp_upload_fallback"][oid] = True
        _blocked_opps.append(opp)
    elif err:
        _fetch_errors.append(f"**{opp.get('opp_name', oid)}**: {err}")
    else:
        parsed = parse_handover_html(html_content, selected_customer)
        _has_content = bool(
            parsed.get('summary') or parsed.get('pain_points') or
            parsed.get('requirements', {}).get('must') or parsed.get('risks')
        )
        if not _has_content:
            # Fetched successfully but got an empty/default page — treat as blocked
            st.session_state["cp_upload_fallback"][oid] = True
            _blocked_opps.append(opp)
        else:
            _customer_cache[oid] = parsed
            st.session_state["cp_handover_cache"][_cache_key] = _customer_cache
            _fetched_docs.append(parsed)

# ── Show upload fallback for any blocked opps ─────────────────────────────────
if _blocked_opps:
    _is_first_block = len(_fetched_docs) == 0 and not _fetch_errors

    st.info(
        "**RevOps site not yet reachable from Streamlit Cloud** — "
        "direct fetch is blocked until the IP is whitelisted. "
        "In the meantime, download the HTML from RevOps and upload it below. "
        "This step goes away once whitelisting is done.",
        icon="ℹ️"
    )

    for opp in _blocked_opps:
        oid = opp['opp_id']
        opp_label = opp.get('opp_name', oid) or oid

        # Build a RevOps direct link the user can click to open the report
        revops_url = f"{HANDOVER_BASE_URL}/?opp_id={oid}"

        st.markdown(
            f'<div style="font-size:13px;margin-bottom:6px">'
            f'<strong>{opp_label}</strong> &nbsp;·&nbsp; '
            f'<a href="{revops_url}" target="_blank" '
            f'style="color:#3B9EFF;font-size:12px">Open in RevOps ↗</a>'
            f' &nbsp;<span style="font-size:11px;opacity:.5">'
            f'(File → Save Page As → Webpage, HTML Only)</span>'
            f'</div>',
            unsafe_allow_html=True
        )

        uploaded = st.file_uploader(
            f"Upload HTML for {opp_label}",
            type=["html", "htm"],
            key=f"cp_upload_{oid}",
            label_visibility="collapsed",
        )

        if uploaded:
            try:
                html_content = uploaded.read().decode("utf-8", errors="replace")
                parsed = parse_handover_html(html_content, selected_customer)
                # Tag as upload source so refresh knows not to auto-fetch
                parsed['source'] = 'upload'
                _customer_cache[oid] = parsed
                st.session_state["cp_handover_cache"][_cache_key] = _customer_cache
                _fetched_docs.append(parsed)
                st.success(f"✓ Loaded from upload: {uploaded.name}")
            except Exception as e:
                st.error(f"Could not read uploaded file: {e}")

if _fetch_errors:
    for err in _fetch_errors:
        st.error(err)

if not _fetched_docs:
    st.markdown(
        '<div class="no-data-msg">No handover data loaded yet.<br>'
        '<span style="font-size:12px;opacity:.6">Upload the HTML file above to continue.</span></div>',
        unsafe_allow_html=True
    )
    st.stop()

# Merge if multiple
d = _merge_docs(_fetched_docs)

# ── Source badge text: live fetch vs upload ───────────────────────────────────
_all_sources = [doc.get('source','') for doc in _fetched_docs]
_source_label = "upload" if all(s == 'upload' for s in _all_sources) else (
                "RevOps live" if all(s == 'revops_fetch' for s in _all_sources) else
                "RevOps · partial upload")
_source_badge_color = "#27AE60" if _source_label == "RevOps live" else "#D68910"

# ── Refresh / clear uploads button ───────────────────────────────────────────
_rcol, _spacer = st.columns([1, 5])
with _rcol:
    if st.button("↺ Refresh", help="Clear cache and re-fetch (or re-upload)", type="secondary"):
        for opp in _selected_opps:
            _customer_cache.pop(opp['opp_id'], None)
            st.session_state["cp_upload_fallback"].pop(opp['opp_id'], None)
        st.session_state["cp_handover_cache"][_cache_key] = _customer_cache
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — CUSTOMER HEADER
# ══════════════════════════════════════════════════════════════════════════════

product_pills = "".join(f'<span class="pill pill-teal">{p}</span>' for p in d['products'])
opp_link_html = (
    f'<a href="{d["opp_link"]}" target="_blank" '
    f'style="font-size:12px;color:#3B9EFF;opacity:.7;margin-left:12px;text-decoration:none">'
    f'↗ SFDC opportunity</a>'
    if d.get('opp_link') else ""
)
_meta_parts = []
if d.get('arr'):    _meta_parts.append(f"ARR: {d['arr']}")
if d.get('ae'):     _meta_parts.append(f"AE: {d['ae']}")
if d.get('region'): _meta_parts.append(d['region'])
if d.get('close_date'): _meta_parts.append(f"Closed: {d['close_date']}")
_meta_str = " · ".join(_meta_parts)

# Multi-opp chips if combined view
_opp_chips_html = ""
if len(_fetched_docs) > 1:
    _opp_chips_html = "<div style='margin-bottom:10px'>" + "".join(
        f'<span class="opp-chip">◎ {doc.get("opp_name","Opp")[:40]}</span>'
        for doc in _fetched_docs
    ) + "</div>"

# ── AI Q&A panel (stub — Phase 2) ────────────────────────────────────────────
st.markdown("""
<div class="ai-stub">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px">
        <div style="font-size:13px;font-weight:700;color:#4472C4">Ask about this customer</div>
        <span style="font-size:11px;color:rgba(128,128,128,.45)">AI · available when Anthropic API is approved</span>
    </div>
    <div style="font-size:12px;opacity:.45;font-style:italic;line-height:1.6">
        e.g. "What were their biggest concerns about change management?"
        &nbsp;·&nbsp; "Which commitments carry the most risk?"
        &nbsp;·&nbsp; "Summarise the technical complexity across all opps"
    </div>
</div>
<hr class='divider' style='margin:12px 0 16px'>
""", unsafe_allow_html=True)

st.markdown(
    f"<div style='margin-bottom:4px'>"
    f"<span style='font-size:22px;font-weight:700'>{selected_customer}</span>"
    f"{opp_link_html}"
    f"<span style='font-size:9px;font-weight:700;padding:1px 7px;border-radius:8px;"
    f"background:{_source_badge_color}22;color:{_source_badge_color};"
    f"border:1px solid {_source_badge_color}44;margin-left:8px;vertical-align:middle;"
    f"letter-spacing:.4px'>{_source_label}</span>"
    f"</div>"
    f"{_opp_chips_html}"
    f"<div style='margin-bottom:8px;font-size:13px;color:rgba(128,128,128,.7)'>{_meta_str}</div>"
    f"<div style='margin-bottom:20px'>{product_pills}</div>",
    unsafe_allow_html=True
)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — NS HOURS LOOKUP (unchanged from original)
# ══════════════════════════════════════════════════════════════════════════════

def _fmt_hrs(h):
    if h is None: return "—"
    rounded = round(float(h) * 4) / 4
    return f"{rounded:g} hrs"

_ns_htd: dict        = {}
_ns_last_entry: dict = {}
_ns_tm_pids: set     = set()
_ns_person_hrs: dict = {}

if df_ns is not None and not df_ns.empty:
    _ns_id_col = "project_id" if "project_id" in df_ns.columns else None
    if _ns_id_col:
        _ns_clean = df_ns.copy()
        _ns_clean[_ns_id_col] = _ns_clean[_ns_id_col].astype(str).str.strip().str.lower()
        if "hours_to_date" in _ns_clean.columns:
            for _pid, _grp in _ns_clean.groupby(_ns_id_col):
                if _pid:
                    _v = _grp["hours_to_date"].dropna().astype(float)
                    if not _v.empty:
                        _ns_htd[_pid] = round(float(_v.max()), 2)
        elif "hours" in _ns_clean.columns:
            for _pid, _grp in _ns_clean.groupby(_ns_id_col):
                if _pid:
                    _v = _grp["hours"].dropna().astype(float)
                    if not _v.empty:
                        _ns_htd[_pid] = round(float(_v.sum()), 2)
        if "date" in _ns_clean.columns:
            _ns_clean["date"] = pd.to_datetime(_ns_clean["date"], errors="coerce")
            for _pid, _grp in _ns_clean.groupby(_ns_id_col):
                if _pid:
                    _d = _grp["date"].dropna()
                    if not _d.empty:
                        _ns_last_entry[_pid] = _d.max()
        if "billing_type" in _ns_clean.columns:
            _tm_mask = _ns_clean["billing_type"].fillna("").str.lower().str.contains("t&m|time", regex=True)
            _ns_tm_pids = {str(p) for p in _ns_clean.loc[_tm_mask, _ns_id_col].dropna() if p}
        _emp_col = "employee" if "employee" in _ns_clean.columns else None
        _hrs_col = "hours" if "hours" in _ns_clean.columns else (
                   "hours_to_date" if "hours_to_date" in _ns_clean.columns else None)
        if _emp_col and _hrs_col:
            for (_pid, _emp), _grp in _ns_clean.groupby([_ns_id_col, _emp_col]):
                _pid_k = str(_pid).strip().lower()
                _emp_k = str(_emp).strip()
                if _pid_k and _emp_k:
                    _h = round(float(_grp[_hrs_col].dropna().astype(float).sum()), 1)
                    if _h > 0:
                        if _pid_k not in _ns_person_hrs:
                            _ns_person_hrs[_pid_k] = {}
                        _ns_person_hrs[_pid_k][_emp_k] = _h


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — DRS PROJECT CARDS (unchanged from original)
# ══════════════════════════════════════════════════════════════════════════════

_drs_match = None
if df_drs is not None and not df_drs.empty and selected_customer:
    _name_col = "project_name" if "project_name" in df_drs.columns else None
    _acct_col = "account" if "account" in df_drs.columns else None
    _sel_lower = selected_customer.strip().lower()
    if _name_col:
        _extracted = df_drs[_name_col].fillna("").apply(_extract_customer_name)
        _ext_lower = _extracted.str.strip().str.lower()
        _drs_match = df_drs[
            (_ext_lower == _sel_lower) |
            _ext_lower.str.startswith(_sel_lower) |
            _ext_lower.apply(lambda x: _sel_lower.startswith(x) if x else False)
        ]
        if _drs_match.empty and _acct_col:
            _drs_match = df_drs[df_drs[_acct_col].fillna("").str.strip().str.lower() == _sel_lower]

_PHASE_ORDER = [
    "00. onboarding", "01. requirements and design", "02. configuration",
    "03. enablement/training", "04. uat", "05. prep for go-live",
    "06. go-live (hypercare)", "08. ready for support transition",
    "10. complete/pending final billing"
]
_PHASE_LABELS = ["Onboarding","Req","Config","Enablement","UAT","Go-Live","Hypercare","Support Tx","Complete"]
_COMPLETE_PHASES = {"10. complete","10. complete/pending final billing","complete","09. complete"}
_HOLD_PHASES = {"11. on hold","on hold"}

def _phase_index(phase_str):
    if not phase_str: return -1
    pl = str(phase_str).lower().strip()
    for i, p in enumerate(_PHASE_ORDER):
        if pl.startswith(p[:6]): return i
    return -1

def _project_status(phase_str):
    pl = str(phase_str).lower().strip()
    if any(pl.startswith(p[:6]) for p in _COMPLETE_PHASES): return "complete"
    if any(pl.startswith(p[:6]) for p in _HOLD_PHASES): return "hold"
    return "active"

def _build_project_card(row, proj_col, lbl_s, val_s, ns_htd=None, ns_tm_pids=None, ns_last_entry=None):
    import pandas as _pd
    phase     = str(row.get("phase","") or "").strip()
    if phase.lower() == "nan": phase = ""
    proj_name = str(row.get(proj_col,"") or "").strip()
    cons      = str(row.get("project_manager","—") or "—").strip()
    proj_type = str(row.get("project_type","—") or "—").strip()
    days      = row.get("days_inactive", None)
    last_act  = row.get("last_activity_date", row.get("last_ns_entry", None))
    status    = _project_status(phase)
    pidx      = _phase_index(phase)
    _pid_for_last = str(row.get("project_id","") or "").strip().lower()
    if last_act is None and ns_last_entry and _pid_for_last in ns_last_entry:
        last_act = ns_last_entry[_pid_for_last]
    last_str, days_str = "—", ""
    if last_act is not None:
        try:
            la = _pd.to_datetime(last_act)
            last_str = la.strftime("%b %d, %Y")
            d_val = int(days) if days is not None and str(days) != "nan" else (_pd.Timestamp.today() - la).days
            if d_val >= 0:
                days_str = "<div style=\"font-size:10px;color:rgba(128,128,128,.5)\">" + str(d_val) + " days ago</div>"
        except Exception:
            last_str = str(last_act)[:10]

    def step_color(i):
        if status == "complete": return "#27AE60"
        if i < pidx: return "#27AE60"
        if i == pidx: return "#D68910" if status == "hold" else "#4472C4"
        return "rgba(128,128,128,.2)"

    lbl_div_style = "flex:1;font-size:8px;color:rgba(128,128,128,.45);text-align:center;overflow:hidden"
    bar_labels = "".join("<div style=\"" + lbl_div_style + "\">" + l + "</div>" for l in _PHASE_LABELS)
    bar_steps  = "".join(
        "<div style=\"flex:1;height:3px;border-radius:2px;background:" + step_color(i) + "\"></div>"
        for i in range(len(_PHASE_LABELS))
    )
    if status == "active":
        pill = "<span style=\"display:inline-block;font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;background:rgba(39,174,96,.12);color:#1A7A4A;margin-top:6px\">Active</span>"
    elif status == "hold":
        pill = "<span style=\"display:inline-block;font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;background:rgba(214,137,16,.12);color:#854F0B;margin-top:6px\">On hold</span>"
    else:
        pill = "<span style=\"display:inline-block;font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;background:rgba(128,128,128,.1);color:rgba(128,128,128,.6);margin-top:6px\">Complete</span>"

    border  = "0.5px solid rgba(214,137,16,.35)" if status == "hold" else "0.5px solid rgba(128,128,128,.15)"
    opacity = "opacity:0.55;" if status == "complete" else ""

    _start_val = row.get("start_date", None)
    start_str = "—"
    if _start_val is not None and str(_start_val) not in ("NaT", "nan", "None", ""):
        try:
            import pandas as _pd2
            _sd = _pd2.to_datetime(_start_val)
            if not _pd2.isnull(_sd):
                start_str = _sd.strftime("%b %d, %Y")
        except Exception:
            pass

    _pid_key   = str(row.get("project_id","") or "").strip().lower()
    _bill_raw  = str(row.get("billing_type","") or "").lower()
    _is_tm     = "t&m" in _bill_raw or "time" in _bill_raw or (ns_tm_pids and _pid_key in ns_tm_pids)
    _ff_scope  = get_ff_scope(str(row.get("project_type","") or ""), proj_name) if not _is_tm else None
    _htd       = ns_htd.get(_pid_key) if ns_htd and _pid_key else None
    _scope     = _htd if _is_tm else (_ff_scope if _ff_scope is not None else None)

    _hours_html = ""
    if _htd is not None or _scope is not None:
        _htd_str   = _fmt_hrs(_htd) if _htd is not None else "—"
        _scope_str = _fmt_hrs(_scope) if _scope is not None else "—"
        _bal       = round(float(_scope) - float(_htd), 1) if _scope is not None and _htd is not None else None
        _bal_color = ("color:#C0392B" if _bal is not None and _bal < 0 else
                      "color:#D68910" if _bal is not None and _scope and _bal / float(_scope) <= 0.10 else
                      "color:var(--color-text-primary)")
        _bal_str   = (("+" if _bal > 0 else "") + _fmt_hrs(_bal)) if _bal is not None else "—"
        _hours_html = (
            "<div style=\"margin-top:8px;padding-top:8px;border-top:0.5px solid rgba(128,128,128,.12);"
            "display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px\">"
            "<div><div style=\"" + lbl_s + "\">Scope</div><div style=\"" + val_s + "\">" + _scope_str + "</div></div>"
            "<div><div style=\"" + lbl_s + "\">Hours to date</div><div style=\"" + val_s + "\">" + _htd_str + "</div></div>"
            "<div><div style=\"" + lbl_s + "\">Balance</div>"
            "<div style=\"font-size:12px;font-weight:500;" + _bal_color + "\">" + _bal_str + "</div></div>"
            "</div>"
        )

    parts = [
        "<div style=\"background:rgba(128,128,128,.04);border:" + border + ";border-radius:10px;padding:12px 14px;" + opacity + "\">",
        "<div style=\"font-size:12px;font-weight:500;color:var(--color-text-primary);margin-bottom:8px;line-height:1.4\">" + proj_name + "</div>",
        "<div style=\"display:flex;gap:0;margin-bottom:3px\">" + bar_labels + "</div>",
        "<div style=\"display:flex;gap:2px;margin-bottom:10px\">" + bar_steps + "</div>",
        "<div style=\"display:grid;grid-template-columns:1fr 1fr;gap:6px\">",
        "<div><div style=\"" + lbl_s + "\">Phase</div><div style=\"" + val_s + "\">" + phase + "</div></div>",
        "<div><div style=\"" + lbl_s + "\">Last activity</div><div style=\"" + val_s + "\">" + last_str + "</div>" + days_str + "</div>",
        "<div><div style=\"" + lbl_s + "\">Consultant</div><div style=\"" + val_s + "\">" + cons + "</div></div>",
        "<div><div style=\"" + lbl_s + "\">Type</div><div style=\"" + val_s + "\">" + proj_type + "</div></div>",
        "<div><div style=\"" + lbl_s + "\">Start date</div><div style=\"" + val_s + "\">" + start_str + "</div></div>",
        "</div>" + _hours_html + pill + "</div>",
    ]
    return "".join(parts)

if _drs_match is not None and not _drs_match.empty:
    _proj_col = "project_name" if "project_name" in _drs_match.columns else (
                "project" if "project" in _drs_match.columns else None)
    _lbl_s = "font-size:9px;text-transform:uppercase;letter-spacing:.5px;color:rgba(128,128,128,.5);margin-bottom:1px"
    _val_s = "font-size:12px;font-weight:500;color:var(--color-text-primary)"
    _active_rows, _complete_rows = [], []
    for _, _row in _drs_match.iterrows():
        _ph = str(_row.get("phase","") or "").lower().strip()
        if any(_ph.startswith(p[:6]) for p in _COMPLETE_PHASES) or _ph == "complete":
            _complete_rows.append(_row)
        else:
            _active_rows.append(_row)
    _n_active, _n_complete = len(_active_rows), len(_complete_rows)
    _n_total = _n_active + _n_complete
    _show_complete_key = f"cp_show_complete_{selected_customer}"
    if _show_complete_key not in st.session_state:
        st.session_state[_show_complete_key] = False
    _hdr_col1, _hdr_col2 = st.columns([3, 1])
    with _hdr_col1:
        st.markdown(
            "<div style=\"display:flex;align-items:center;gap:6px;margin-bottom:10px\">"
            "<div style=\"width:8px;height:8px;border-radius:50%;background:#4472C4;flex-shrink:0\"></div>"
            "<div style=\"font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:rgba(128,128,128,.6)\">DRS — project data</div>"
            "</div>",
            unsafe_allow_html=True
        )
    with _hdr_col2:
        _show_complete = (
            st.toggle(f"Show {_n_complete} completed", value=st.session_state[_show_complete_key],
                      key=_show_complete_key)
            if _n_complete > 0 else False
        )
    _display_rows = _active_rows + (_complete_rows if _show_complete else [])
    if _display_rows and _proj_col:
        _cards_html = "<div style=\"display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin-bottom:8px\">"
        for _row in _display_rows:
            _cards_html += _build_project_card(_row, _proj_col, _lbl_s, _val_s,
                                               ns_htd=_ns_htd, ns_tm_pids=_ns_tm_pids, ns_last_entry=_ns_last_entry)
        _cards_html += "</div>"
        st.markdown(_cards_html, unsafe_allow_html=True)
        _n_hold = sum(1 for r in _active_rows if _project_status(str(r.get("phase","") or "")) == "hold")
        _n_act  = _n_active - _n_hold
        _parts  = []
        if _n_act:      _parts.append(f"{_n_act} active")
        if _n_hold:     _parts.append(f"{_n_hold} on hold")
        if _n_complete: _parts.append(f"{_n_complete} complete")
        st.markdown(
            "<div style=\"font-size:11px;color:var(--color-text-tertiary);text-align:right;margin-bottom:8px\">"
            + f"{_n_total} project{'s' if _n_total != 1 else ''} · " + " · ".join(_parts) + "</div>",
            unsafe_allow_html=True
        )


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 9 — TABS
# ══════════════════════════════════════════════════════════════════════════════

tab_overview, tab_stakeholders, tab_requirements, tab_usecases, tab_risks, tab_notes, tab_usage = st.tabs([
    "Overview", "Stakeholders", "Requirements", "Use Cases", "Risks & Commitments", "Notes", "Usage"
])


# ─── TAB: OVERVIEW ───────────────────────────────────────────────────────────
with tab_overview:
    _ov_col1, _ov_col2 = st.columns([1, 1])
    with _ov_col1:
        if d.get("summary"):
            st.markdown('<div class="section-label">Summary</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="cp-card" style="font-size:13px;line-height:1.7;color:inherit">{d["summary"]}</div>',
                        unsafe_allow_html=True)
        if d.get("pain_points"):
            st.markdown('<div class="section-label" style="margin-top:16px">Pain points</div>',
                        unsafe_allow_html=True)
            pp_html = "".join(
                '<div class="cp-bullet">' + t + ('<span class="cp-flag">review</span>' if f else "") + "</div>"
                for t, f in d["pain_points"]
            )
            st.markdown(f'<div class="cp-card" style="padding:10px 14px">{pp_html}</div>',
                        unsafe_allow_html=True)
    with _ov_col2:
        req = d.get("requirements", {})
        req_flat = (req if isinstance(req, list) else req.get("must", []) + req.get("nice", []))
        if req_flat:
            st.markdown('<div class="section-label">Key requirements</div>', unsafe_allow_html=True)
            req_html = "".join(
                '<div class="req-row">' + (t[0] if isinstance(t, tuple) else t) + "</div>"
                for t in req_flat[:4]
            )
            st.markdown(f'<div class="cp-card" style="padding:10px 14px">{req_html}</div>',
                        unsafe_allow_html=True)
            if len(req_flat) > 4:
                st.caption(f"{len(req_flat) - 4} more in the Requirements tab.")
        _watch = [r for r in d.get("risks", []) if r.get("flagged")] + \
                 [c for c in d.get("commitments", []) if c.get("status") == "risk"]
        if _watch:
            st.markdown('<div class="section-label" style="margin-top:16px">Watch items</div>',
                        unsafe_allow_html=True)
            watch_html = ""
            for item in _watch[:4]:
                _wtext  = item.get("text", "")
                _wcolor = "#C0392B" if item.get("status") == "risk" else "#D68910"
                _wicon  = "✕" if item.get("status") == "risk" else "⚠"
                watch_html += (
                    f'<div style="display:flex;gap:8px;align-items:flex-start;padding:5px 0;'
                    f'border-bottom:0.5px solid rgba(128,128,128,.1);font-size:13px;line-height:1.6">'
                    f'<span style="color:{_wcolor};flex-shrink:0">{_wicon}</span>'
                    f'<span>{_wtext}</span></div>'
                )
            st.markdown(f'<div class="cp-card" style="padding:10px 14px">{watch_html}</div>',
                        unsafe_allow_html=True)
    if d.get("info_gaps"):
        st.markdown('<div class="section-label" style="margin-top:8px">Assumptions & information gaps</div>',
                    unsafe_allow_html=True)
        ig_html = "".join(f'<div class="info-gap-row">{g}</div>' for g in d["info_gaps"])
        st.markdown(f'<div class="cp-card" style="padding:10px 14px">{ig_html}</div>',
                    unsafe_allow_html=True)
    if d.get("competitors"):
        st.markdown('<div class="section-label" style="margin-top:8px">Competitors mentioned</div>',
                    unsafe_allow_html=True)
        comp_html = "".join(f'<div class="cp-bullet">{c}</div>' for c in d["competitors"])
        st.markdown(f'<div class="cp-card" style="padding:10px 14px">{comp_html}</div>',
                    unsafe_allow_html=True)


# ─── TAB: STAKEHOLDERS ───────────────────────────────────────────────────────
with tab_stakeholders:

    _sfdc_contacts = []
    _opp_ids_for_sfdc = [opp['opp_id'] for opp in _selected_opps]

    if df_sfdc is not None and not df_sfdc.empty:
        _opp_id_col_s  = next((c for c in ('opportunity_id','opp_id','sf_opp_id') if c in df_sfdc.columns), None)
        _first_col     = next((c for c in ('first_name',) if c in df_sfdc.columns), None)
        _last_col      = next((c for c in ('last_name',) if c in df_sfdc.columns), None)
        _title_col     = next((c for c in ('title','contact_title') if c in df_sfdc.columns), None)
        _email_col     = next((c for c in ('email',) if c in df_sfdc.columns), None)
        _roles_col     = next((c for c in ('contact_roles','roles') if c in df_sfdc.columns), None)
        _primary_col   = next((c for c in ('is_primary',) if c in df_sfdc.columns), None)

        if _opp_id_col_s:
            _sfdc_rows = df_sfdc[df_sfdc[_opp_id_col_s].astype(str).str.strip().isin(_opp_ids_for_sfdc)]
        else:
            _sfdc_rows = df_sfdc[
                df_sfdc.get('account', pd.Series(dtype=str)).fillna('').str.lower().str.contains(
                    re.escape(selected_customer.lower()[:12]), na=False, regex=True
                )
            ] if 'account' in df_sfdc.columns else pd.DataFrame()

        for _, _sr in _sfdc_rows.iterrows():
            _fname = str(_sr.get(_first_col or 'first_name', '') or '').strip()
            _lname = str(_sr.get(_last_col or 'last_name', '') or '').strip()
            _fullname = f"{_fname} {_lname}".strip()
            if not _fullname: continue
            _roles_raw = str(_sr.get(_roles_col or 'contact_roles', '') or '')
            _roles = [r.strip() for r in _roles_raw.split(',') if r.strip()]
            _sfdc_contacts.append({
                'name':       _fullname,
                'title':      str(_sr.get(_title_col or 'title', '') or '').strip(),
                'email':      str(_sr.get(_email_col or 'email', '') or '').strip(),
                'roles':      _roles,
                'is_primary': str(_sr.get(_primary_col or 'is_primary', '0')).strip() == '1',
                'source':     'sfdc',
                'internal':   False,
            })

    # Gong stakeholders from parsed handover
    _gong_stakeholders = d.get('stakeholders', [])
    _gong_external = [s for s in _gong_stakeholders if not s.get('internal')]
    _gong_internal = [s for s in _gong_stakeholders if s.get('internal')]

    # Merge Gong external + SFDC by fuzzy name match
    from rapidfuzz import fuzz as _fuzz
    _merged_external, _sfdc_used = [], set()

    for _gs in _gong_external:
        _best_score, _best_idx = 0, None
        for _i, _sc in enumerate(_sfdc_contacts):
            if _i in _sfdc_used: continue
            _score = _fuzz.token_set_ratio(_gs['name'].lower(), _sc['name'].lower())
            if _score > _best_score:
                _best_score, _best_idx = _score, _i
        if _best_idx is not None and _best_score >= 80:
            _sc = _sfdc_contacts[_best_idx]
            _sfdc_used.add(_best_idx)
            _merged_external.append({
                'name':       _sc['name'],
                'title':      _sc['title'] or _gs.get('title', ''),
                'email':      _sc['email'] or _gs.get('email', ''),
                'roles':      _sc['roles'],
                'is_primary': _sc['is_primary'],
                'role_note':  _gs.get('role_note', ''),
                'influence':  _gs.get('influence', ''),
                'sentiment':  _gs.get('sentiment', ''),
                'source':     'both',
                'internal':   False,
            })
        else:
            _merged_external.append({
                'name':       _gs['name'],
                'title':      _gs.get('title', ''),
                'email':      _gs.get('email', ''),
                'roles':      [],
                'is_primary': False,
                'role_note':  _gs.get('role_note', ''),
                'influence':  _gs.get('influence', ''),
                'sentiment':  _gs.get('sentiment', ''),
                'source':     'gong',
                'internal':   False,
            })

    for _i, _sc in enumerate(_sfdc_contacts):
        if _i not in _sfdc_used:
            _merged_external.append({**_sc, 'role_note': '', 'influence': '', 'sentiment': ''})

    _merged_external.sort(key=lambda x: (not x.get('is_primary'), x['name']))

    def _build_stakeholder_html(stakeholders):
        if not stakeholders:
            return "<div style='font-size:13px;opacity:.4;padding:8px'>None listed.</div>"
        rows_html = ""
        for s in stakeholders:
            initials = _initials(s['name'])
            av_cls   = "avatar avatar-int" if s.get('internal') else "avatar"
            # Influence pill
            infl = s.get('influence', '').upper()
            infl_html = ""
            if infl in ('HIGH', 'MED', 'LOW'):
                infl_col = {"HIGH": "#C0392B", "MED": "#D68910", "LOW": "#888"}.get(infl, "#888")
                infl_bg  = {"HIGH": "rgba(192,57,43,.1)", "MED": "rgba(243,156,18,.1)", "LOW": "rgba(128,128,128,.1)"}.get(infl, "")
                infl_html = f'<span style="font-size:9px;font-weight:700;padding:1px 5px;border-radius:6px;background:{infl_bg};color:{infl_col};margin-left:4px">{infl}</span>'
            # Sentiment
            sent = s.get('sentiment', '')
            sent_html = f'<span style="font-size:10px;color:rgba(128,128,128,.5);margin-left:4px">{sent}</span>' if sent else ""
            # Role pills
            role_pills = ""
            _priority_roles = ["Decision Maker","Primary Contact","Implementation Contact","Economic Buyer","Champion"]
            for _r in s.get('roles', []):
                if any(_r.lower() == p.lower() for p in _priority_roles):
                    role_pills += f'<span class="pill pill-blue" style="font-size:10px">{_r}</span> '
            if not role_pills and s.get('role_note'):
                role_pills = f'<span class="pill pill-grey" style="font-size:10px">{s["role_note"]}</span>'
            title_str = f'<span style="font-size:11px;opacity:.6">{s["title"]}</span>' if s.get("title") else ""
            email_str = (
                f'<a href="mailto:{s["email"]}" style="font-size:11px;color:#3B9EFF;opacity:.7">{s["email"]}</a>'
                if s.get("email") and "@" in s["email"] else ""
            )
            _src = s.get('source', '')
            _src_color = {"sfdc":"#27AE60","gong":"#D68910","both":"#4472C4","drs":"#4472C4","ns":"rgba(128,128,128,.6)"}.get(_src,"#888")
            _src_label = {"sfdc":"SFDC","gong":"Gong","both":"SFDC+Gong","drs":"DRS","ns":"NS time"}.get(_src, _src)
            src_badge = (f'<span style="font-size:9px;font-weight:700;padding:1px 5px;border-radius:8px;'
                         f'background:{_src_color}22;color:{_src_color};margin-left:4px">{_src_label}</span>') if _src else ""
            primary_dot = '<span style="color:#4472C4;font-size:10px;margin-right:3px">●</span>' if s.get('is_primary') else ""
            rows_html += (
                f'<div class="stakeholder-row">'
                f'<div class="{av_cls}">{initials}</div>'
                f'<div style="flex:1;min-width:0">'
                f'<div style="font-size:13px;font-weight:600">{primary_dot}{s["name"]}{src_badge}{infl_html}{sent_html}</div>'
                f'<div>{title_str}</div><div>{email_str}</div>'
                f'<div style="margin-top:3px">{role_pills}</div>'
                f'</div></div>'
            )
        return f'<div class="cp-card" style="padding:8px 14px">{rows_html}</div>'

    col_ext, col_int = st.columns(2)

    with col_ext:
        _src_note = (" <span style=\"font-size:10px;font-weight:400;opacity:.5;text-transform:none;letter-spacing:0\">· SFDC matched</span>"
                     if _sfdc_contacts else "")
        st.markdown(f'<div class="section-label">Customer contacts{_src_note}</div>', unsafe_allow_html=True)
        if not _merged_external:
            st.markdown('<div class="no-data-msg">No contacts found.</div>', unsafe_allow_html=True)
        else:
            st.markdown(_build_stakeholder_html(_merged_external), unsafe_allow_html=True)

    with col_int:
        st.markdown('<div class="section-label">Zone team</div>', unsafe_allow_html=True)

        # SFDC AM
        _sfdc_am_contacts = []
        if df_sfdc is not None and not df_sfdc.empty:
            _am_opp_col = next((c for c in ('opportunity_id','opp_id') if c in df_sfdc.columns), None)
            if _am_opp_col:
                _am_rows = df_sfdc[df_sfdc[_am_opp_col].astype(str).str.strip().isin(_opp_ids_for_sfdc)]
                if not _am_rows.empty:
                    _am_row = _am_rows.iloc[0]
                    _am_name  = str(_am_row.get("account_manager","") or "").strip()
                    _am_email = str(_am_row.get("account_manager_email","") or "").strip()
                    if _am_name:
                        _sfdc_am_contacts.append({
                            "name":_am_name,"title":"Account Manager","email":_am_email,
                            "roles":[],"is_primary":False,"role_note":"","source":"sfdc",
                            "internal":True,"hrs":None,"influence":"","sentiment":"",
                        })

        def _resolve_title(name):
            gr = get_role(name)
            er = EMPLOYEE_ROLES.get(name, {})
            if gr == "consultant" and not er:
                parts = name.split()
                flipped = f"{parts[-1]}, {' '.join(parts[:-1])}" if len(parts) >= 2 else name
                gr = get_role(flipped)
                er = EMPLOYEE_ROLES.get(flipped, {})
            er_role = er.get("role","")
            if gr in ("manager","manager_only") or er_role == "Project Manager": return "Project Manager"
            if "senior" in er_role.lower(): return "Senior Consultant"
            return "Implementation Consultant"

        _drs_assigned = []
        _assigned_names = set()
        if _drs_match is not None and not _drs_match.empty:
            _pm_col2 = "project_manager" if "project_manager" in _drs_match.columns else None
            _pn_col2 = "project_name" if "project_name" in _drs_match.columns else None
            _pid_col2 = "project_id" if "project_id" in _drs_match.columns else None
            if _pm_col2:
                _seen_asgn = set()
                for _, _pm_row in _drs_match.iterrows():
                    _pm_name = str(_pm_row.get(_pm_col2,"") or "").strip()
                    _pn_name = str(_pm_row.get(_pn_col2,"") or "").strip() if _pn_col2 else ""
                    _pid_val = str(_pm_row.get(_pid_col2,"") or "").strip().lower() if _pid_col2 else ""
                    if _pm_name and _pm_name != "—" and _pm_name not in _seen_asgn:
                        _seen_asgn.add(_pm_name)
                        _assigned_names.add(_pm_name.lower())
                        _p_hrs = _ns_person_hrs.get(_pid_val, {}).get(_pm_name) if _pid_val else None
                        _drs_assigned.append({
                            "name":_pm_name,"title":_resolve_title(_pm_name),"email":"",
                            "roles":[],"is_primary":False,
                            "role_note": _pn_name[:38]+("…" if len(_pn_name)>38 else ""),
                            "source":"drs","internal":True,"hrs":_p_hrs,
                            "influence":"","sentiment":"",
                        })

        _ns_contributors = []
        if _drs_match is not None and not _drs_match.empty and _ns_person_hrs:
            _pid_col3 = "project_id" if "project_id" in _drs_match.columns else None
            _contrib_seen = set()
            for _, _pr in _drs_match.iterrows():
                _pid_v = str(_pr.get(_pid_col3,"") or "").strip().lower() if _pid_col3 else ""
                if not _pid_v or _pid_v not in _ns_person_hrs: continue
                for _emp_name, _hrs in _ns_person_hrs[_pid_v].items():
                    if _emp_name.lower() in _assigned_names: continue
                    if _emp_name.lower() in _contrib_seen: continue
                    _contrib_seen.add(_emp_name.lower())
                    _ns_contributors.append({
                        "name":_emp_name,"title":_resolve_title(_emp_name),"email":"",
                        "roles":[],"is_primary":False,"role_note":"","source":"ns",
                        "internal":True,"hrs":_hrs,"influence":"","sentiment":"",
                    })
            _ns_contributors.sort(key=lambda x: x.get("hrs") or 0, reverse=True)

        _assigned_all = _sfdc_am_contacts + _drs_assigned + [{**c,"hrs":None} for c in _gong_internal]
        _seen_zone, _assigned_dedup = set(), []
        for _zt in _assigned_all:
            _n = _zt["name"].lower()
            if _n not in _seen_zone:
                _seen_zone.add(_n)
                _assigned_dedup.append(_zt)

        def _zone_row_html(person):
            _ini = "".join(p[0].upper() for p in person["name"].split()[:2])
            _title_s = (f'<div style="font-size:11px;color:var(--color-text-secondary)">{person["title"]}</div>') if person.get("title") else ""
            _rn = person.get("role_note","")
            _rn_html = (f'<div style="font-size:10px;color:var(--color-text-tertiary);margin-top:1px">{_rn}</div>') if _rn else ""
            _src = person.get("source","")
            _pill_map = {"drs":("#4472C4","DRS"),"sfdc":("#27AE60","SFDC"),"gong":("#D68910","Gong"),"ns":("rgba(128,128,128,.6)","NS time")}
            _pc, _pl = _pill_map.get(_src,("rgba(128,128,128,.6)",_src))
            _src_html = (f'<span style="font-size:9px;font-weight:700;padding:1px 5px;border-radius:8px;'
                         f'background:{_pc}22;color:{_pc}">{_pl}</span>') if _src else ""
            _hrs_v = person.get("hrs")
            _hrs_html = (f'<span style="font-size:10px;color:rgba(128,128,128,.5);margin-left:4px">{_hrs_v}h</span>'
                         if _hrs_v is not None else "")
            return (
                f'<div class="stakeholder-row">'
                f'<div class="avatar avatar-int">{_ini}</div>'
                f'<div style="flex:1;min-width:0">'
                f'<div style="font-size:13px;font-weight:600">{person["name"]}{_hrs_html} {_src_html}</div>'
                f'{_title_s}{_rn_html}'
                f'</div></div>'
            )

        _zone_rows = _assigned_dedup + _ns_contributors
        if _zone_rows:
            _zone_html = '<div class="cp-card" style="padding:8px 14px">'
            if _assigned_dedup:
                _zone_html += '<div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:rgba(128,128,128,.4);padding:4px 0 2px">Assigned</div>'
                for p in _assigned_dedup:
                    _zone_html += _zone_row_html(p)
            if _ns_contributors:
                _zone_html += '<div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:rgba(128,128,128,.4);padding:10px 0 2px;margin-top:4px;border-top:0.5px solid rgba(128,128,128,.1)">NS contributors</div>'
                for p in _ns_contributors:
                    _zone_html += _zone_row_html(p)
            _zone_html += "</div>"
            st.markdown(_zone_html, unsafe_allow_html=True)
        else:
            st.caption("Load SFDC Contacts and DRS to see Zone team.")


# ─── TAB: REQUIREMENTS ───────────────────────────────────────────────────────
with tab_requirements:
    req = d.get("requirements", {})
    must_items = req.get("must", []) if isinstance(req, dict) else req
    nice_items = req.get("nice", []) if isinstance(req, dict) else []

    col_must, col_nice = st.columns([3, 2])
    with col_must:
        st.markdown('<div class="section-label">Must-have</div>', unsafe_allow_html=True)
        if must_items:
            html_m = "".join(
                f'<div class="req-row">{(t[0] if isinstance(t, tuple) else t)}</div>'
                for t in must_items
            )
            st.markdown(f'<div class="cp-card" style="padding:10px 14px">{html_m}</div>',
                        unsafe_allow_html=True)
        else:
            st.markdown('<div class="no-data-msg">None parsed.</div>', unsafe_allow_html=True)
    with col_nice:
        st.markdown('<div class="section-label">Nice-to-have</div>', unsafe_allow_html=True)
        if nice_items:
            html_n = "".join(
                f'<div class="req-row req-nice">{(t[0] if isinstance(t, tuple) else t)}</div>'
                for t in nice_items
            )
            st.markdown(f'<div class="cp-card" style="padding:10px 14px;opacity:.8">{html_n}</div>',
                        unsafe_allow_html=True)
        else:
            st.caption("None listed.")

    if d.get("tech_env"):
        st.markdown('<div class="section-label" style="margin-top:16px">Technical environment</div>',
                    unsafe_allow_html=True)
        te_html = "".join(
            '<div class="cp-bullet">' + t + '</div>'
            for t, _ in d["tech_env"]
        )
        st.markdown(f'<div class="cp-card" style="padding:10px 14px">{te_html}</div>',
                    unsafe_allow_html=True)

    if d.get("timeline"):
        st.markdown('<div class="section-label" style="margin-top:16px">Delivery constraints & timeline</div>',
                    unsafe_allow_html=True)
        tl_html = "".join(
            '<div class="cp-bullet">' + t + '</div>'
            for t, _ in d["timeline"]
        )
        st.markdown(f'<div class="cp-card" style="padding:10px 14px">{tl_html}</div>',
                    unsafe_allow_html=True)


# ─── TAB: USE CASES ──────────────────────────────────────────────────────────
with tab_usecases:
    use_cases = d.get("use_cases", [])
    if not use_cases:
        st.markdown('<div class="no-data-msg">No use cases parsed.</div>', unsafe_allow_html=True)
    else:
        for uc in use_cases:
            title = uc.get('title', '')
            meta_str = uc.get('meta', '')
            st.markdown(
                f'<div class="cp-card" style="padding:12px 16px">'
                f'<div style="font-size:13px;font-weight:700;margin-bottom:4px">{title}</div>'
                f'<div style="font-size:12px;color:rgba(128,128,128,.65)">{meta_str}</div>'
                f'</div>',
                unsafe_allow_html=True
            )


# ─── TAB: RISKS & COMMITMENTS ────────────────────────────────────────────────
with tab_risks:

    def _classify_commit(text):
        t = text.lower()
        if any(w in t for w in ["price","discount","contract","commercial","billing","invoice","payment"]):
            return "commercial"
        if any(w in t for w in ["timeline","date","schedule","deadline","start","kickoff","week","month"]):
            return "process"
        return "ps"

    def _render_commit_list(commits, muted=False):
        icon_map = {"aligned": "✅", "review": "⚠️", "risk": "🔴"}
        html = ""
        for c in commits:
            icon    = icon_map.get(c["status"], "—")
            opacity = ' style="opacity:.5"' if muted else (' style="opacity:.6"' if c["status"] == "risk" else "")
            html += (
                f'<div style="display:flex;gap:10px;align-items:flex-start;padding:5px 0;'
                f'border-bottom:0.5px solid rgba(128,128,128,.1);font-size:13px;line-height:1.6">'
                f'<span class="commit-icon">{icon}</span><span{opacity}>{c["text"]}</span></div>'
            )
        return html

    def _render_risk_list(risks):
        badge_map = {
            "tech":     ("Technical",    "risk-tech"),
            "exp":      ("Expectation",  "risk-exp"),
            "org":      ("Org readiness","risk-org"),
            "timeline": ("Timeline",     "risk-timeline"),
        }
        html = ""
        for r in risks:
            label, css = badge_map.get(r["badge"], ("Risk","risk-tech"))
            sev = r.get("severity","")
            sev_html = (f'<span style="font-size:9px;font-weight:700;padding:1px 5px;border-radius:6px;'
                        f'background:rgba(192,57,43,.12);color:#C0392B;margin-left:4px">{sev}</span>'
                        if sev.upper() == "HIGH" else
                        f'<span style="font-size:9px;font-weight:700;padding:1px 5px;border-radius:6px;'
                        f'background:rgba(243,156,18,.1);color:#D68910;margin-left:4px">{sev}</span>'
                        if sev else "")
            html += (
                f'<div style="display:flex;gap:8px;align-items:flex-start;padding:6px 0;'
                f'border-bottom:0.5px solid rgba(128,128,128,.1);font-size:13px;line-height:1.6">'
                f'<span class="risk-badge {css}">{label}</span>'
                f'<span>{r["category"]}{sev_html}<br>'
                f'<span style="font-size:12px;opacity:.8">{r["text"]}</span></span></div>'
            )
        return html

    ps_commits   = [c for c in d.get("commitments",[]) if _classify_commit(c["text"]) == "ps"]
    comm_commits = [c for c in d.get("commitments",[]) if _classify_commit(c["text"]) == "commercial"]
    proc_commits = [c for c in d.get("commitments",[]) if _classify_commit(c["text"]) == "process"]
    tech_risks   = [r for r in d.get("risks",[]) if r["badge"] == "tech"]
    exp_risks    = [r for r in d.get("risks",[]) if r["badge"] == "exp"]
    org_risks    = [r for r in d.get("risks",[]) if r["badge"] in ("org","timeline")]

    legend = """<div style="font-size:11px;opacity:.45;margin-top:6px">
        ✅ aligned &nbsp;·&nbsp; ⚠️ needs review &nbsp;·&nbsp; 🔴 risk / flag
    </div>"""

    sub_commits, sub_risks = st.tabs(["Commitments", "Risks"])

    with sub_commits:
        if ps_commits:
            st.markdown('<div class="section-label">PS & implementation</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="cp-card" style="padding:10px 14px">{_render_commit_list(ps_commits)}</div>',
                        unsafe_allow_html=True)
        if comm_commits:
            st.markdown('<div class="section-label" style="margin-top:16px">Commercial</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="cp-card" style="padding:10px 14px;opacity:.75">{_render_commit_list(comm_commits, muted=True)}</div>',
                        unsafe_allow_html=True)
        if proc_commits:
            st.markdown('<div class="section-label" style="margin-top:16px">Process & scheduling</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="cp-card" style="padding:10px 14px;opacity:.65">{_render_commit_list(proc_commits, muted=True)}</div>',
                        unsafe_allow_html=True)
        if not d.get("commitments"):
            st.markdown('<div class="no-data-msg">No commitments parsed.</div>', unsafe_allow_html=True)
        else:
            st.markdown(legend, unsafe_allow_html=True)

    with sub_risks:
        if not d.get("risks"):
            st.markdown('<div class="no-data-msg">No risks parsed.</div>', unsafe_allow_html=True)
        else:
            if exp_risks:
                st.markdown('<div class="section-label">Expectation alignment</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="cp-card" style="padding:10px 14px">{_render_risk_list(exp_risks)}</div>',
                            unsafe_allow_html=True)
            if tech_risks:
                st.markdown('<div class="section-label" style="margin-top:16px">Technical complexity</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="cp-card" style="padding:10px 14px">{_render_risk_list(tech_risks)}</div>',
                            unsafe_allow_html=True)
            if org_risks:
                st.markdown('<div class="section-label" style="margin-top:16px">Org readiness & timeline</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="cp-card" style="padding:10px 14px">{_render_risk_list(org_risks)}</div>',
                            unsafe_allow_html=True)


# ─── TAB: NOTES ──────────────────────────────────────────────────────────────
with tab_notes:
    _notes_key = f"cp_notes_{selected_customer}"
    if _notes_key not in st.session_state:
        st.session_state[_notes_key] = []
    _notes = st.session_state[_notes_key]
    st.markdown('<div class="section-label">Add a note</div>', unsafe_allow_html=True)
    _note_text = st.text_area(
        "Note",
        placeholder="e.g. Maarten confirmed NS access is pending. Session 2 scheduled for Apr 22.",
        height=90, label_visibility="collapsed",
        key=f"cp_note_input_{selected_customer}"
    )
    if st.button("Save note", key=f"cp_note_save_{selected_customer}"):
        if _note_text.strip():
            from datetime import datetime as _dt
            st.session_state[_notes_key].insert(0, {
                "author": _session_name,
                "ts":     _dt.now().strftime("%b %d, %Y · %I:%M %p"),
                "text":   _note_text.strip()
            })
            st.rerun()
    if _notes:
        st.markdown('<div class="section-label" style="margin-top:16px">Notes</div>', unsafe_allow_html=True)
        for _n in _notes:
            _author_short = _n["author"].split(",")[0] if "," in _n["author"] else _n["author"]
            st.markdown(f"""
<div style="background:var(--cp-card-bg,rgba(128,128,128,.05));border:0.5px solid rgba(128,128,128,.15);
            border-radius:8px;padding:10px 14px;margin-bottom:8px">
    <div style="font-size:11px;color:rgba(128,128,128,.5);margin-bottom:4px">
        {_author_short} &nbsp;·&nbsp; {_n["ts"]}
    </div>
    <div style="font-size:13px;line-height:1.6;color:var(--color-text-primary)">{_n["text"]}</div>
</div>""", unsafe_allow_html=True)
    else:
        st.markdown('<div class="no-data-msg">No notes yet — add one above.</div>', unsafe_allow_html=True)
    st.markdown("""
<div style="font-size:11px;opacity:.4;margin-top:16px;text-align:center">
    Notes are session-only until persistent storage is enabled in Phase 2.
</div>""", unsafe_allow_html=True)


# ─── TAB: USAGE ──────────────────────────────────────────────────────────────
with tab_usage:
    st.markdown('<div class="section-label">NetSuite usage data</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="usage-stub">
        <div style="font-size:22px;opacity:.3;margin-bottom:8px">⏱</div>
        <div style="font-weight:600;margin-bottom:6px">Usage sync not yet configured</div>
        <div style="font-size:12px;opacity:.7">
            When connected, this section will show: first login date · active user count ·
            module adoption · recent activity trend — pulled directly from NetSuite.
        </div>
    </div>""", unsafe_allow_html=True)


# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="font-size:11px;opacity:.35;text-align:center;margin-top:16px">
    PS Projects & Tools · Internal use only · Handover data fetched live from RevOps this session
</div>""", unsafe_allow_html=True)
