"""
PS Tools — Customer Profile
Gong handover intelligence: pain points, stakeholders, requirements,
commitments, risks, and information gaps — parsed from uploaded Gong .docx exports.
"""
import streamlit as st
import pandas as pd
import re
import io
from datetime import date

st.session_state["current_page"] = "Customer Profile"

from shared.constants import (
    EMPLOYEE_ROLES, get_role, is_manager, name_matches,
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
    .ai-stub{border:1px dashed rgba(128,128,128,.25);border-radius:10px;
             padding:16px 18px;margin-top:8px;opacity:.65}
    .usage-stub{border:1px dashed rgba(128,128,128,.2);border-radius:10px;
                padding:24px;text-align:center;color:rgba(128,128,128,.6);
                font-size:13px}
    .opp-tab-active{border-bottom:2px solid #4472C4;color:#4472C4;font-weight:700}
    .req-row{font-size:13px;padding:5px 0 5px 16px;position:relative;
             line-height:1.6;border-bottom:0.5px solid rgba(128,128,128,.1)}
    .req-row:last-child{border-bottom:none}
    .req-row::before{content:'✓';position:absolute;left:2px;top:5px;
                     font-size:10px;color:rgba(128,128,128,.4)}
    .req-nice::before{content:'○';color:rgba(128,128,128,.35)}
    .no-data-msg{text-align:center;padding:36px;opacity:.4;font-size:14px}

</style>
""", unsafe_allow_html=True)

# ── Zone SVG watermark ────────────────────────────────────────────────────────
_zone_svg = """<svg style='position:absolute;right:-40px;top:50%;transform:translateY(-50%);
opacity:0.06;width:200px;height:200px;pointer-events:none'
viewBox='0 0 1482 1286.25' xmlns='http://www.w3.org/2000/svg'>
<g fill='#3B9EFF' fill-rule='evenodd'><path d='M975.127,924.953c2.608-2.68,1.744-5.496-.42-7.829l-57.415-61.872c-2.463-2.655-5.025-2.878-8.443-.991-10.398,5.739-19.024,12.314-27.949,19.885-83.252,70.621-197.471,155.494-298.93,195.556-17.993,7.105-35.256,13.178-54.191,17.329-62.148,13.627-131.853,15.491-192.702-5.298-64.93-22.183-113.878-68.722-142.715-130.542-28.647-61.415-22.393-131.406,11.352-189.217,2.598-2.793,1.405-6.055-1.389-8.184-35.341-26.918-40.303-33.439-69.367-65.686-1.449-1.607-4.102-2.401-5.903-1.138-13.105,9.189-23.232,20.534-33.172,32.961-16.499,20.629-29.73,42.605-38.718,67.541-5.127,10.469-8.378,20.486-10.885,32.065-13.633,62.973-7.701,128.685,17.402,188.142,23.839,56.463,65.297,103.638,114.77,139.169,32.418,23.283,66.848,42.548,103.476,58.385,25.142,10.871,50.281,18.994,76.934,25.12,96.392,22.153,188.876,4.496,276.774-38.393,42.916-20.94,83.188-45.685,121.922-73.568,75.733-54.514,154.643-126.72,219.571-193.435Z'/></g></svg>"""

# ── Hero banner ───────────────────────────────────────────────────────────────
st.markdown(f"""
<div style='background:#050D1F;padding:32px 40px 28px;border-radius:10px;
            margin-bottom:24px;font-family:Manrope,sans-serif;
            position:relative;overflow:hidden'>
  {_zone_svg}
  <div style='font-size:13px;font-weight:700;letter-spacing:2.5px;
              text-transform:uppercase;color:#3B9EFF;margin-bottom:10px'>
      Professional Services · Tools</div>
  <h1 style='color:white;margin:0;font-size:28px;font-family:Manrope,sans-serif;
             font-weight:800'>Customer Profile</h1>
  <p style='color:rgba(255,255,255,0.45);margin:6px 0 0;font-size:14px;
            font-family:Manrope,sans-serif;max-width:560px'>
    Gong handover intelligence — pain points, stakeholders, requirements,
    commitments and risks. Select a customer and upload their Gong export to begin.</p>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — GONG DOC PARSER
# ══════════════════════════════════════════════════════════════════════════════

def _initials(name: str) -> str:
    """Return up to 2 initials from a name string."""
    parts = str(name).strip().split()
    if not parts:
        return "?"
    if len(parts) == 1:
        return parts[0][:2].upper()
    return (parts[0][0] + parts[-1][0]).upper()


def _strip_flag(text: str) -> tuple[str, bool]:
    """Remove trailing ⚠️ and return (clean_text, was_flagged)."""
    flagged = "⚠️" in text
    clean = text.replace("⚠️", "").strip().rstrip(".")
    return clean, flagged


def _parse_section(text: str, heading: str) -> list[str]:
    """
    Extract bullet lines under a heading.
    Returns list of raw strings (may contain ⚠️).
    Handles headings with or without trailing colon.
    """
    pattern = rf"###\s+{re.escape(heading)}:?\s*\n(.*?)(?=\n###|\Z)"
    m = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    if not m:
        return []
    block = m.group(1)
    lines = []
    for line in block.split("\n"):
        line = line.strip().lstrip("-").lstrip("*").strip()
        if line:
            lines.append(line)
    return lines


def _parse_stakeholders(text: str) -> list[dict]:
    """Parse the Stakeholders section into structured dicts."""
    lines = _parse_section(text, "Stakeholders")
    stakeholders = []
    for line in lines:
        # Format: Name | Title | Internal/External | email | role_note
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 2:
            stakeholders.append({
                "name":     parts[0] if len(parts) > 0 else "",
                "title":    parts[1] if len(parts) > 1 else "",
                "internal": (parts[2].lower() == "internal") if len(parts) > 2 else False,
                "email":    parts[3] if len(parts) > 3 else "",
                "role_note": parts[4] if len(parts) > 4 else "",
            })
    return stakeholders


def _parse_requirements(text: str) -> dict:
    """Parse must-have and nice-to-have requirements.
    Handles Nice-to-Have appearing inline at end of a must-have line."""
    pattern = r"###\s+Requirements:?\s*\n(.*?)(?=\n###|\Z)"
    m = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    if not m:
        return {"must": [], "nice": []}
    block = m.group(1)
    # Nice-to-Have can appear inline (no preceding newline)
    parts = re.split(r"Nice-to-Have\s*:", block, maxsplit=1, flags=re.IGNORECASE)
    must_block = parts[0]
    nice_block = parts[1] if len(parts) > 1 else ""

    def _extract_items(b):
        items = []
        for line in b.split("\n"):
            l = line.strip().lstrip("-").lstrip("*").strip()
            if l and not re.match(r"must.have", l, re.IGNORECASE):
                items.append(l)
        return items

    return {"must": _extract_items(must_block), "nice": _extract_items(nice_block)}


def _parse_commitments(text: str) -> list[dict]:
    """Parse Sales Commitments into structured list with status."""
    pattern = r"###\s+Sales Commitments:?\s*\n(.*?)(?=\n###|\Z)"
    m = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    if not m:
        return []
    block = m.group(1)
    commits = []
    for line in block.split("\n"):
        l = line.strip().lstrip("-").lstrip("*").strip()
        if not l:
            continue
        if l.startswith("✅"):
            status = "aligned"
            text_part = l[1:].lstrip("aligned").lstrip("—").strip()
        elif l.startswith("⚠️"):
            status = "review"
            text_part = l[2:].lstrip("review").lstrip("—").strip()
        elif l.startswith("🔴"):
            status = "risk"
            text_part = l[2:].lstrip("risk").lstrip("—").strip()
        else:
            status = "aligned"
            text_part = l
        # Strip the "✅ aligned — " prefix pattern
        text_part = re.sub(r"^(aligned|review|risk)\s*[—-]\s*", "", text_part, flags=re.IGNORECASE).strip()
        if text_part:
            commits.append({"status": status, "text": text_part})
    return commits


def _parse_risks(text: str) -> list[dict]:
    """Parse Risks section into structured list."""
    lines = _parse_section(text, "Risks")
    risks = []
    for line in lines:
        clean, flagged = _strip_flag(line)
        # Format: "category — description"
        if " — " in clean:
            cat, desc = clean.split(" — ", 1)
            cat = cat.strip().lower()
        elif " - " in clean:
            cat, desc = clean.split(" - ", 1)
            cat = cat.strip().lower()
        else:
            cat, desc = "general", clean
        # Normalise category to badge type
        if "technical" in cat:
            badge = "tech"
        elif "expectation" in cat:
            badge = "exp"
        elif "org" in cat or "readiness" in cat:
            badge = "org"
        elif "timeline" in cat:
            badge = "timeline"
        else:
            badge = "tech"
        risks.append({"badge": badge, "category": cat.title(), "text": desc.strip(), "flagged": flagged})
    return risks


def _parse_summary(text: str) -> str:
    pattern = r"###\s+Summary:?\s*\n(.*?)(?=\n###|\Z)"
    m = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    if not m:
        return ""
    # Strip leading "- " from each line (docx bullet format)
    lines = []
    for line in m.group(1).split("\n"):
        line = line.strip().lstrip("-").strip()
        if line:
            lines.append(line)
    return " ".join(lines)


def _parse_technical_env(text: str) -> list[str]:
    return _parse_section(text, "Technical Environment")


def _parse_info_gaps(text: str) -> list[str]:
    return _parse_section(text, "Information Gap")


def _parse_opportunity_link(text: str) -> str:
    m = re.search(r"Opportunity Link:\s*(https?://\S+)", text)
    return m.group(1) if m else ""


def _parse_data_used(text: str) -> dict:
    pattern = r"###\s+Data Used:?\s*\n(.*?)(?=\n###|\Z)"
    m = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    if not m:
        return {}
    block = m.group(1).strip().lstrip("-").strip()
    result = {}
    # Known keys in the Data Used line
    for key in ["Account", "Sales Owner", "Calls analyzed", "Call dates", "Emails analyzed"]:
        km = re.search(
            rf"{re.escape(key)}:\s*([^\n]*?)(?=\s+(?:Account|Sales Owner|Calls analyzed|Call dates|Emails analyzed):|$)",
            block
        )
        if km:
            val = km.group(1).strip()
            if val:
                result[key] = val
    return result


def _extract_products_from_text(text: str) -> list[str]:
    """Extract Zone product tags from document — strict matching only.
    Uses explicit Zone-prefixed terms or unambiguous product names to avoid
    false positives from generic words like 'billing', 'payments', 'approvals'."""
    known = {
        "ZoneCapture":   [r"zonecapture", r"zone capture", r"zone\s+capture"],
        "ZoneApprovals": [r"zoneapprovals", r"zone approvals", r"zone\s+approvals"],
        "AP Payments":   [r"zone\s*ap\s*payments", r"zone\s*payments", r"ap payments module"],
        "e-Invoicing":   [r"e-invoicing", r"einvoicing", r"zone.*invoic"],
        "Reconcile":     [r"zonereconcile", r"zone reconcile"],
        "ZoneBilling":   [r"zonebilling", r"zone billing"],
        "ZonePremium":   [r"zonepremium", r"zone premium"],
    }
    import re as _re
    found = []
    tl = text.lower()
    for label, patterns in known.items():
        if any(_re.search(p, tl) for p in patterns):
            found.append(label)
    return found


def _read_docx_text(uploaded_file) -> tuple[str, str]:
    """Extract plain text from a .docx using stdlib only (zipfile + xml).
    No python-docx dependency needed — .docx is a zip containing XML.
    Returns (text, error_message) — error_message is empty string on success."""
    try:
        import zipfile
        from xml.etree import ElementTree as ET
        raw_bytes = uploaded_file.getvalue()
        zf = zipfile.ZipFile(io.BytesIO(raw_bytes))
        xml_content = zf.read("word/document.xml")
        tree = ET.fromstring(xml_content)
        W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        lines = []
        for para in tree.iter(f"{{{W}}}p"):
            pPr = para.find(f"{{{W}}}pPr")
            style_id = ""
            if pPr is not None:
                pStyle = pPr.find(f"{{{W}}}pStyle")
                if pStyle is not None:
                    style_id = pStyle.get(f"{{{W}}}val", "")
            text = "".join(
                node.text or "" for node in para.iter(f"{{{W}}}t")
            ).strip()
            if not text:
                continue
            is_heading = bool(re.search(r"heading|^[123]$", style_id, re.IGNORECASE))
            lines.append(f"### {text}" if is_heading else f"- {text}")
        result = "\n".join(lines)
        if not result.strip():
            return "", "Document appears empty — check the file and try again."
        return result, ""
    except Exception as e:
        return "", f"Error reading .docx: {e}"


def _parse_filename(filename: str) -> dict:
    """Extract customer, opp name, and date from Gong handover filename.
    Expected format: Handover - {OppName}-{CustomerName}-{YYYY-MM-DD}.docx
    Example: Handover - ZonePayroll 66 emps for Life Trading-Life Trading-2026-04-15.docx
    Returns dict with keys: customer, opp_name, close_date (all may be empty strings).
    """
    name = filename.replace(".docx", "").strip()
    name = re.sub(r"^Handover\s*-\s*", "", name, flags=re.IGNORECASE).strip()
    date_match = re.search(r"[-\s](\d{4}-\d{2}-\d{2})$", name)
    close_date = date_match.group(1) if date_match else ""
    if close_date:
        name = name[:date_match.start()].strip().rstrip("-").strip()
    if " - " in name:
        parts = name.rsplit(" - ", 1)
        opp_name, customer = parts[0].strip(), parts[1].strip()
    elif "-" in name:
        parts = name.rsplit("-", 1)
        opp_name, customer = parts[0].strip(), parts[1].strip()
    else:
        opp_name = customer = name
    return {"customer": customer, "opp_name": opp_name, "close_date": close_date}


def _parse_gong_doc(raw_text: str, customer_name: str) -> dict:
    """Parse full Gong doc text into structured intelligence dict."""
    return {
        "customer":       customer_name,
        "opp_link":       _parse_opportunity_link(raw_text),
        "data_used":      _parse_data_used(raw_text),
        "summary":        _parse_summary(raw_text),
        "pain_points":    [_strip_flag(l) for l in _parse_section(raw_text, "Pain Points")],
        "stakeholders":   _parse_stakeholders(raw_text),
        "requirements":   _parse_requirements(raw_text),
        "use_cases":      [_strip_flag(l) for l in _parse_section(raw_text, "Use Cases")],
        "tech_env":       [_strip_flag(l) for l in _parse_technical_env(raw_text)],
        "timeline":       [_strip_flag(l) for l in _parse_section(raw_text, "Delivery Constraints & Timeline")],
        "commitments":    _parse_commitments(raw_text),
        "info_gaps":      _parse_info_gaps(raw_text),
        "risks":          _parse_risks(raw_text),
        "products":       _extract_products_from_text(raw_text),
        "raw":            raw_text,
    }


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — CUSTOMER SELECTOR + UPLOAD
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<div style='background:var(--color-background-secondary, rgba(59,158,255,0.05));
            border-left:4px solid #3B9EFF;border-radius:6px;
            padding:16px 20px;margin:0 0 20px;font-family:Manrope,sans-serif'>
    <div style='font-size:11px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;
                color:#3B9EFF;margin-bottom:10px'>Roadmap</div>
    <div style='display:flex;gap:32px;flex-wrap:wrap'>
        <div style='flex:1;min-width:220px'>
            <span style='background:#1E2C63;color:#fff;font-size:10px;font-weight:700;
                         padding:2px 8px;border-radius:10px;letter-spacing:1px'>PHASE 1 · NOW</span>
            <p style='margin:8px 0 0;font-size:13px;color:inherit;line-height:1.6'>
                Upload a Gong AI handover <strong>.docx</strong> for any customer to surface
                pain points, stakeholders, requirements, commitments, risks, and information gaps
                — structured and ready before your first call.
            </p>
        </div>
        <div style='flex:1;min-width:220px'>
            <span style='background:rgba(59,158,255,0.15);color:#3B9EFF;font-size:10px;font-weight:700;
                         padding:2px 8px;border-radius:10px;letter-spacing:1px;
                         border:1px solid rgba(59,158,255,0.4)'>PHASE 2 · COMING SOON</span>
            <p style='margin:8px 0 0;font-size:13px;color:inherit;opacity:0.65;line-height:1.6'>
                <strong>AI Q&amp;A powered by Claude</strong> — ask natural language questions
                about any customer using their Gong intelligence, requirements, and NetSuite
                usage data as context.
            </p>
        </div>
        <div style='flex:1;min-width:220px'>
            <span style='background:rgba(59,158,255,0.15);color:#3B9EFF;font-size:10px;font-weight:700;
                         padding:2px 8px;border-radius:10px;letter-spacing:1px;
                         border:1px solid rgba(59,158,255,0.4)'>PHASE 3 · PLANNED</span>
            <p style='margin:8px 0 0;font-size:13px;color:inherit;opacity:0.65;line-height:1.6'>
                <strong>NetSuite usage sync</strong> — first login, active users, module adoption,
                and activity trend pulled directly from NetSuite alongside the handover intelligence.
            </p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Build customer list from DRS
df_drs  = st.session_state.get("df_drs")
df_sfdc = st.session_state.get("df_sfdc")
drs_customers = []
if df_drs is not None and not df_drs.empty and "project_name" in df_drs.columns:
    drs_customers = sorted(set(
        str(r).split(" - ")[0].strip()
        for r in df_drs["project_name"].dropna()
        if str(r).strip()
    ))

# Session store for uploaded Gong docs: {customer_name: [parsed_doc, ...]}
# NOTE: session-only storage — docs do not persist across page reloads.
# TODO: Replace with persistent backend (DB/blob) when hosting is configured.
if "cp_gong_docs" not in st.session_state:
    st.session_state["cp_gong_docs"] = {}

_autofill = st.session_state.pop("_cp_autofill_customer", "")
_autofill_opp = st.session_state.pop("_cp_autofill_opp", "")

col_sel, col_upload = st.columns([1, 1])

with col_sel:
    st.markdown('<div class="section-label">Customer</div>', unsafe_allow_html=True)

    if drs_customers:
        customer_options = ["— Select a customer —"] + drs_customers + ["+ Enter manually"]
        _default_idx = 0
        if _autofill and _autofill in drs_customers:
            _default_idx = drs_customers.index(_autofill) + 1
        _sel = st.selectbox("Customer", customer_options, index=_default_idx,
                            label_visibility="collapsed")
        if _sel == "+ Enter manually":
            selected_customer = st.text_input("Customer name",
                                              value=_autofill,
                                              placeholder="e.g. Habyt GmbH",
                                              label_visibility="collapsed")
        elif _sel == "— Select a customer —":
            selected_customer = _autofill  # use autofill even if not in DRS list
        else:
            selected_customer = _sel
    else:
        st.markdown('<div style="font-size:12px;opacity:.5;margin-bottom:6px">DRS not loaded — enter name manually</div>', unsafe_allow_html=True)
        selected_customer = st.text_input("Customer name",
                                          value=_autofill,
                                          placeholder="e.g. Habyt GmbH",
                                          label_visibility="collapsed")

with col_upload:
    st.markdown('<div class="section-label">Gong handover doc</div>', unsafe_allow_html=True)
    st.markdown('<div style="font-size:12px;opacity:.45;margin-bottom:6px">Session only · .docx · re-upload each session</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Gong doc",
        type=["docx"],
        label_visibility="collapsed",
    )

    # Auto-fill customer name from filename if it follows the Gong naming convention
    if uploaded and not selected_customer:
        _fn_parsed = _parse_filename(uploaded.name)
        if _fn_parsed["customer"]:
            st.session_state["_cp_autofill_customer"] = _fn_parsed["customer"]
            st.session_state["_cp_autofill_opp"] = _fn_parsed["opp_name"]
            st.rerun()

    if uploaded and selected_customer:
        raw_text, err = _read_docx_text(uploaded)
        if err:
            st.error(f"Could not read the file: {err}")
        elif not raw_text.strip():
            st.error("Document appears empty — check the file and try again.")
        else:
            parsed = _parse_gong_doc(raw_text, selected_customer)
            # Enrich with filename metadata if available
            _fn_meta = _parse_filename(uploaded.name)
            if _fn_meta["opp_name"] and not parsed.get("opp_name"):
                parsed["opp_name"] = _fn_meta["opp_name"]
            if _fn_meta["close_date"] and not parsed.get("close_date"):
                parsed["close_date"] = _fn_meta["close_date"]
            docs = st.session_state["cp_gong_docs"].get(selected_customer, [])
            existing_links = [d.get("opp_link", "") for d in docs]
            if parsed["opp_link"] not in existing_links or not parsed["opp_link"]:
                docs.append(parsed)
                st.session_state["cp_gong_docs"][selected_customer] = docs
                st.session_state["_cp_notif_doc"] = f"✓ Loaded: {uploaded.name}"
                st.rerun()
            else:
                st.session_state["_cp_notif_doc"] = "This opportunity doc is already loaded."

# ── Inline notifications (same row, after upload processing) ─────────────────
_notif_customer = st.session_state.pop("_cp_notif_customer", "")
_notif_doc = st.session_state.pop("_cp_notif_doc", "")
if _autofill or _notif_customer or _notif_doc:
    _nc1, _nc2 = st.columns([1, 1])
    with _nc1:
        if _autofill:
            _msg = f"Customer: **{_autofill}**"
            if _autofill_opp:
                _msg += f" · Opp: **{_autofill_opp}**"
            st.info(_msg)
    with _nc2:
        if _notif_doc:
            st.info(_notif_doc)

if not selected_customer:
    st.markdown("""
    <div class="no-data-msg">
        Select a customer above to view their profile.<br>
        <span style="font-size:12px;opacity:.6">Upload a Gong handover .docx to populate intelligence sections.</span>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── Docs for this customer ────────────────────────────────────────────────────
customer_docs = st.session_state["cp_gong_docs"].get(selected_customer, [])

st.markdown("<hr class='divider'>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — OPPORTUNITY SELECTOR (if multiple docs)
# ══════════════════════════════════════════════════════════════════════════════

_has_gong_doc = bool(customer_docs)

if not _has_gong_doc:
    st.markdown(f"""
    <div style='margin-bottom:16px'>
        <span style='font-size:20px;font-weight:700'>{selected_customer}</span>
    </div>
    <div class='no-data-msg' style='margin-bottom:12px'>
        No Gong handover doc uploaded yet.<br>
        <span style='font-size:12px;opacity:.6'>Upload a .docx above to populate intelligence sections — DRS and Notes are available without it.</span>
    </div>
    """, unsafe_allow_html=True)

# ── Opp selector (only if Gong doc uploaded) ─────────────────────────────────
def _opp_label(doc: dict, idx: int) -> str:
    products = " · ".join(doc["products"][:3]) if doc["products"] else f"Opportunity {idx+1}"
    calls = doc["data_used"].get("Calls analyzed", "")
    calls_str = f" · {calls} calls" if calls and calls != "0" else ""
    return f"{products}{calls_str}"

d = None  # will be set below if Gong doc available

if _has_gong_doc and len(customer_docs) > 1:
    st.markdown('<div class="section-label">Opportunities</div>', unsafe_allow_html=True)
    opp_labels = [_opp_label(d, i) for i, d in enumerate(customer_docs)]
    # Add remove option
    opp_labels_with_action = opp_labels + ["— Remove a doc —"]
    _opp_tab = st.selectbox("Select opportunity", opp_labels_with_action, label_visibility="collapsed")
    if _opp_tab == "— Remove a doc —":
        _to_remove = st.selectbox("Which doc to remove?", opp_labels, label_visibility="collapsed")
        if st.button("Remove this doc", type="secondary"):
            idx_remove = opp_labels.index(_to_remove)
            customer_docs.pop(idx_remove)
            st.session_state["cp_gong_docs"][selected_customer] = customer_docs
            st.rerun()
        st.stop()
    active_doc = customer_docs[opp_labels.index(_opp_tab)]
    d = active_doc
elif _has_gong_doc:
    d = customer_docs[0]


# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — CUSTOMER HEADER (Gong doc required)
# ══════════════════════════════════════════════════════════════════════════════

if d is not None:
    calls_info  = d["data_used"].get("Calls analyzed", "")
    emails_info = d["data_used"].get("Emails analyzed", "")
    data_meta = []
    if calls_info:
        data_meta.append(f"{calls_info} calls analysed")
    if emails_info and emails_info != "0":
        data_meta.append(f"{emails_info} emails analysed")
    _opp_id_match = re.search(r"/Opportunity/([A-Za-z0-9]+)/", d.get("opp_link", ""))
    if _opp_id_match:
        data_meta.append(f"SF Opp: {_opp_id_match.group(1)}")
    data_meta_str  = " · ".join(data_meta) if data_meta else ""
    product_pills  = "".join(f'<span class="pill pill-teal">{p}</span>' for p in d["products"])
    opp_link_html  = (
        f'<a href="{d["opp_link"]}" target="_blank" '
        f'style="font-size:12px;color:#3B9EFF;opacity:.7;margin-left:12px;text-decoration:none">↗ SFDC opportunity</a>'
        if d["opp_link"] else ""
    )
    st.markdown("""
<div class="ai-stub" style="margin-bottom:16px">
    <div style="display:flex;align-items:center;gap:10px;margin-bottom:6px">
        <div style="font-size:13px;font-weight:700;color:#4472C4">Ask about this customer</div>
        <span style="font-size:11px;color:rgba(128,128,128,.45)">AI · available when API keys are configured</span>
    </div>
    <div style="font-size:12px;opacity:.45;font-style:italic;line-height:1.6">
        e.g. "What were their biggest concerns about change management?"
        &nbsp;·&nbsp; "Which commitments carry the most risk?"
        &nbsp;·&nbsp; "Summarise the key technical complexity"
    </div>
</div>
<hr class='divider' style='margin:12px 0 16px'>
""", unsafe_allow_html=True)
    _opp_name_display = d.get("opp_name", "")
    _opp_name_html = (
        f'<div style="font-size:13px;color:rgba(128,128,128,.55);margin-top:2px;margin-bottom:2px">Opp: {_opp_name_display}</div>'
        if _opp_name_display else ""
    )
    st.markdown(f"""
<div style='margin-bottom:4px'>
    <span style='font-size:22px;font-weight:700'>{selected_customer}</span>
    {opp_link_html}
</div>
{_opp_name_html}
<div style='margin-bottom:8px;font-size:13px;color:rgba(128,128,128,.7)'>
    {data_meta_str}
</div>
<div style='margin-bottom:20px'>
    {product_pills}
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 4b — DRS PROJECT STRIP
# ══════════════════════════════════════════════════════════════════════════════

# Pull matching DRS rows for this customer
_drs_match = None
if df_drs is not None and not df_drs.empty and selected_customer:
    _name_col = "project_name" if "project_name" in df_drs.columns else None
    _acct_col = "account" if "account" in df_drs.columns else None
    if _name_col:
        _drs_match = df_drs[
            df_drs[_name_col].fillna("").str.lower().str.contains(
                re.escape(selected_customer.lower()[:12]), na=False
            )
        ]
        if _drs_match.empty and _acct_col:
            _drs_match = df_drs[
                df_drs[_acct_col].fillna("").str.lower().str.contains(
                    re.escape(selected_customer.lower()[:12]), na=False
                )
            ]

# Phase order for progress bar
_PHASE_ORDER = [
    "00. onboarding", "01. requirements and design", "02. configuration",
    "03. enablement/training", "04. uat", "05. prep for go-live",
    "06. go-live (hypercare)", "08. ready for support transition",
    "10. complete/pending final billing"
]
_PHASE_LABELS = ["Onboarding", "Req & Design", "Config", "Enablement", "UAT",
                 "Pre-Go-Live", "Hypercare", "Support Tx", "Complete"]

def _phase_index(phase_str):
    if not phase_str:
        return -1
    pl = str(phase_str).lower().strip()
    for i, p in enumerate(_PHASE_ORDER):
        if pl.startswith(p[:6]):
            return i
    return -1

if _drs_match is not None and not _drs_match.empty:
    # Project selector when customer has multiple projects
    _proj_col = "project_name" if "project_name" in _drs_match.columns else (
                "project" if "project" in _drs_match.columns else None)

    if len(_drs_match) > 1 and _proj_col:
        _proj_labels = [str(r.get(_proj_col, f"Project {i+1}")).strip()
                        for i, (_, r) in enumerate(_drs_match.iterrows())]
        # Deduplicate
        _seen_p, _proj_labels_dedup = {}, []
        for lbl in _proj_labels:
            _seen_p[lbl] = _seen_p.get(lbl, 0) + 1
            _proj_labels_dedup.append(lbl if _seen_p[lbl] == 1 else f"{lbl} ({_seen_p[lbl]})")
        st.markdown('<div class="section-label">Project</div>', unsafe_allow_html=True)
        _sel_proj = st.selectbox("Select project", _proj_labels_dedup,
                                 key=f"cp_proj_sel_{selected_customer}",
                                 label_visibility="collapsed")
        _dr = _drs_match.iloc[_proj_labels_dedup.index(_sel_proj)]
    else:
        _dr = _drs_match.iloc[0]

    _phase     = str(_dr.get("phase", "—")).strip()
    _phase_idx = _phase_index(_phase)
    _cons      = str(_dr.get("project_manager", "—")).strip()
    _proj_type = str(_dr.get("project_type", "—")).strip()
    _days      = _dr.get("days_inactive", None)
    _last_act  = _dr.get("last_activity_date", _dr.get("last_ns_entry", None))

    # Phase progress bar HTML
    _bar_labels = "".join(
        f'<div style="flex:1;font-size:9px;color:rgba(128,128,128,.5);text-align:center">{l}</div>'
        for l in _PHASE_LABELS
    )
    def _step_color(i):
        if i < _phase_idx: return "#27AE60"
        if i == _phase_idx: return "#4472C4"
        return "rgba(128,128,128,.2)"
    _bar_steps = "".join(
        f'<div style="flex:1;height:4px;border-radius:2px;background:{_step_color(i)}"></div>'
        for i in range(len(_PHASE_LABELS))
    )

    _last_str = "—"
    _days_str = ""
    if _last_act is not None:
        try:
            import pandas as pd
            _la = pd.to_datetime(_last_act)
            _last_str = _la.strftime("%b %d, %Y")
            _d = int(_days) if _days is not None and str(_days) != "nan" else (pd.Timestamp.today() - _la).days
            if _d >= 0:
                _days_str = f'<div style="font-size:11px;color:rgba(128,128,128,.5)">{_d} days ago</div>'
        except Exception:
            _last_str = str(_last_act)[:10]

    _lbl_s = "font-size:10px;text-transform:uppercase;letter-spacing:.6px;color:rgba(128,128,128,.5);margin-bottom:2px"
    _val_s = "font-size:13px;font-weight:500;color:var(--color-text-primary)"
    _drs_html = "".join([
        "<div style=\"background:rgba(128,128,128,.05);border:0.5px solid rgba(128,128,128,.15);border-radius:10px;padding:14px 18px;margin-bottom:16px\">",
        "<div style=\"display:flex;align-items:center;gap:6px;margin-bottom:10px\">",
        "<div style=\"width:8px;height:8px;border-radius:50%;background:#4472C4;flex-shrink:0\"></div>",
        "<div style=\"font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:rgba(128,128,128,.6)\">DRS \u2014 project data</div>",
        "<div style=\"margin-left:auto;font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;background:rgba(128,128,128,.08);color:rgba(128,128,128,.5);border:0.5px dashed rgba(128,128,128,.3)\">Live sync \u2014 Phase 2</div>",
        "</div>",
        f"<div style=\"display:flex;gap:0;margin-bottom:4px\">{_bar_labels}</div>",
        f"<div style=\"display:flex;gap:3px;margin-bottom:12px\">{_bar_steps}</div>",
        "<div style=\"display:grid;grid-template-columns:repeat(4,1fr);gap:12px\">",
        f"<div><div style=\"{_lbl_s}\">Phase</div><div style=\"{_val_s}\">{_phase}</div></div>",
        f"<div><div style=\"{_lbl_s}\">Last activity</div><div style=\"{_val_s}\">{_last_str}</div>{_days_str}</div>",
        f"<div><div style=\"{_lbl_s}\">Consultant</div><div style=\"{_val_s}\">{_cons}</div></div>",
        f"<div><div style=\"{_lbl_s}\">Project type</div><div style=\"{_val_s}\">{_proj_type}</div></div>",
        "</div></div>",
    ])
    st.markdown(_drs_html, unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — TABS
# ══════════════════════════════════════════════════════════════════════════════

tab_overview, tab_stakeholders, tab_requirements, tab_usecases, tab_risks, tab_notes, tab_usage = st.tabs([
    "Overview", "Stakeholders", "Requirements", "Use Cases", "Risks & Commitments", "Notes", "Usage"
])

# ─── TAB: OVERVIEW ───────────────────────────────────────────────────────────
with tab_overview:
    if d is None:
        st.markdown('<div class="no-data-msg">Upload a Gong handover doc to populate intelligence.</div>',
                    unsafe_allow_html=True)
    else:
        _ov_col1, _ov_col2 = st.columns([1, 1])
        with _ov_col1:
            if d["summary"]:
                st.markdown('<div class="section-label">Summary</div>', unsafe_allow_html=True)
                st.markdown(f'<div class="cp-card" style="font-size:13px;line-height:1.7;color:inherit">{d["summary"]}</div>',
                            unsafe_allow_html=True)
            if d["pain_points"]:
                st.markdown('<div class="section-label" style="margin-top:16px">Pain points</div>',
                            unsafe_allow_html=True)
                pp_html = "".join(
                    '<div class="cp-bullet">' + t + ('<span class="cp-flag">review</span>' if f else "") + "</div>"
                    for t, f in d["pain_points"]
                )
                st.markdown(f'<div class="cp-card" style="padding:10px 14px">{pp_html}</div>',
                            unsafe_allow_html=True)
        with _ov_col2:
            req_flat = []
            if d.get("requirements"):
                r = d["requirements"]
                req_flat = (r if isinstance(r, list) else r.get("must", []) + r.get("nice", []))
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
            _watch = [r for r in d.get("risks", []) if r.get("flagged")] +                      [c for c in d.get("commitments", []) if c.get("status") == "risk"]
            if _watch:
                st.markdown('<div class="section-label" style="margin-top:16px">Watch items</div>',
                            unsafe_allow_html=True)
                watch_html = ""
                for item in _watch[:4]:
                    _wtext = item.get("text", "")
                    _wcolor = "#C0392B" if item.get("status") == "risk" else "#D68910"
                    _wicon = "✕" if item.get("status") == "risk" else "⚠"
                    watch_html += (
                        f'<div style="display:flex;gap:8px;align-items:flex-start;padding:5px 0;'
                        f'border-bottom:0.5px solid rgba(128,128,128,.1);font-size:13px;line-height:1.6">'
                        f'<span style="color:{_wcolor};flex-shrink:0">{_wicon}</span>'
                        f'<span>{_wtext}</span></div>'
                    )
                st.markdown(f'<div class="cp-card" style="padding:10px 14px">{watch_html}</div>',
                            unsafe_allow_html=True)
        if d.get("info_gaps"):
            st.markdown('<div class="section-label" style="margin-top:8px">Information gaps</div>',
                        unsafe_allow_html=True)
            ig_html = "".join(f'<div class="info-gap-row">{g}</div>' for g in d["info_gaps"])
            st.markdown(f'<div class="cp-card" style="padding:10px 14px">{ig_html}</div>',
                        unsafe_allow_html=True)


# ─── TAB: STAKEHOLDERS ───────────────────────────────────────────────────────
with tab_stakeholders:

    # ── Build SFDC contact list keyed by Opp ID ───────────────────────────────
    _sfdc_contacts = []  # list of dicts
    _opp_id = d.get("opp_link", "") if d else ""
    import re as _re
    _opp_id_match = _re.search(r"/Opportunity/([A-Za-z0-9]+)/", _opp_id)
    _opp_id_clean = _opp_id_match.group(1) if _opp_id_match else None

    if df_sfdc is not None and not df_sfdc.empty and _opp_id_clean and "opportunity_id" in df_sfdc.columns:
        _sfdc_rows = df_sfdc[df_sfdc["opportunity_id"].astype(str).str.strip() == _opp_id_clean]
        for _, _sr in _sfdc_rows.iterrows():
            _fname = str(_sr.get("first_name", "") or "").strip()
            _lname = str(_sr.get("last_name", "") or "").strip()
            _fullname = f"{_fname} {_lname}".strip()
            if not _fullname:
                continue
            _roles_raw = str(_sr.get("contact_roles", "") or "")
            _roles = [r.strip() for r in _roles_raw.split(",") if r.strip()]
            _is_primary = str(_sr.get("is_primary", "0")).strip() == "1"
            _sfdc_contacts.append({
                "name":      _fullname,
                "title":     str(_sr.get("title", "") or "").strip(),
                "email":     str(_sr.get("email", "") or "").strip(),
                "roles":     _roles,
                "is_primary": _is_primary,
                "source":    "sfdc",
                "internal":  False,
            })

    # Fall back to account name match if no opp ID
    elif df_sfdc is not None and not df_sfdc.empty and selected_customer and "account" in df_sfdc.columns:
        _sfdc_rows = df_sfdc[
            df_sfdc["account"].fillna("").str.lower().str.contains(
                selected_customer.lower()[:12], na=False
            )
        ]
        for _, _sr in _sfdc_rows.iterrows():
            _fname = str(_sr.get("first_name", "") or "").strip()
            _lname = str(_sr.get("last_name", "") or "").strip()
            _fullname = f"{_fname} {_lname}".strip()
            if not _fullname:
                continue
            _roles_raw = str(_sr.get("contact_roles", "") or "")
            _roles = [r.strip() for r in _roles_raw.split(",") if r.strip()]
            _sfdc_contacts.append({
                "name":      _fullname,
                "title":     str(_sr.get("title", "") or "").strip(),
                "email":     str(_sr.get("email", "") or "").strip(),
                "roles":     _roles,
                "is_primary": str(_sr.get("is_primary", "0")).strip() == "1",
                "source":    "sfdc",
                "internal":  False,
            })

    # ── Merge Gong + SFDC contacts ────────────────────────────────────────────
    # Gong stakeholders marked internal stay as-is (Zone team)
    # External Gong stakeholders: enrich from SFDC by name match, or add as Gong-only
    # SFDC contacts not matched in Gong: add as SFDC-only
    from rapidfuzz import fuzz as _fuzz

    _gong_external = [s for s in (d["stakeholders"] if d else []) if not s["internal"]]
    _gong_internal = [s for s in (d["stakeholders"] if d else []) if s["internal"]]

    _merged_external = []
    _sfdc_used = set()

    for _gs in _gong_external:
        _best_score, _best_idx = 0, None
        for _i, _sc in enumerate(_sfdc_contacts):
            if _i in _sfdc_used:
                continue
            _score = _fuzz.token_set_ratio(_gs["name"].lower(), _sc["name"].lower())
            if _score > _best_score:
                _best_score, _best_idx = _score, _i
        if _best_idx is not None and _best_score >= 80:
            # Merge: SFDC wins on email/title if Gong is empty, keep Gong role_note
            _sc = _sfdc_contacts[_best_idx]
            _sfdc_used.add(_best_idx)
            _merged_external.append({
                "name":      _sc["name"],
                "title":     _sc["title"] or _gs["title"],
                "email":     _sc["email"] or _gs["email"],
                "roles":     _sc["roles"],
                "is_primary": _sc["is_primary"],
                "role_note": _gs.get("role_note", ""),
                "source":    "both",
                "internal":  False,
            })
        else:
            # Gong only
            _merged_external.append({
                "name":      _gs["name"],
                "title":     _gs["title"],
                "email":     _gs["email"],
                "roles":     [],
                "is_primary": False,
                "role_note": _gs.get("role_note", ""),
                "source":    "gong",
                "internal":  False,
            })

    # Add unmatched SFDC contacts
    for _i, _sc in enumerate(_sfdc_contacts):
        if _i not in _sfdc_used:
            _merged_external.append({**_sc, "role_note": ""})

    # Sort: primary first, then by name
    _merged_external.sort(key=lambda x: (not x.get("is_primary"), x["name"]))

    # ── Render ────────────────────────────────────────────────────────────────
    def _build_stakeholder_html(stakeholders, show_source=True):
        if not stakeholders:
            return "<div style='font-size:13px;opacity:.4;padding:8px'>None listed.</div>"
        rows_html = ""
        for s in stakeholders:
            initials = _initials(s["name"])
            av_cls = "avatar avatar-int" if s.get("internal") else "avatar"
            # Role pills from SFDC contact_roles
            role_pills = ""
            _priority_roles = ["Decision Maker", "Primary Contact", "Implementation Contact", "Economic Buyer"]
            for _r in s.get("roles", []):
                if any(_r.lower() == p.lower() for p in _priority_roles):
                    role_pills += f'<span class="pill pill-blue" style="font-size:10px">{_r}</span> '
            # Gong role note as fallback
            if not role_pills and s.get("role_note"):
                role_pills = f'<span class="pill pill-grey" style="font-size:10px">{s["role_note"]}</span>'
            title_str = (f'<span style="font-size:11px;opacity:.6">{s["title"]}</span>') if s.get("title") else ""
            email_str = (
                f'<a href="mailto:{s["email"]}" style="font-size:11px;color:#3B9EFF;opacity:.7">{s["email"]}</a>'
                if s.get("email") and "@" in s["email"] else ""
            )
            # Source badge
            _src = s.get("source", "")
            if show_source and _src:
                _src_color = {"sfdc": "#27AE60", "gong": "#D68910", "both": "#4472C4"}.get(_src, "#888")
                _src_label = {"sfdc": "SFDC", "gong": "Gong", "both": "SFDC + Gong"}.get(_src, _src)
                src_badge = (f'<span style="font-size:9px;font-weight:700;padding:1px 5px;'
                             f'border-radius:8px;background:{_src_color}22;color:{_src_color};'
                             f'margin-left:4px">{_src_label}</span>')
            else:
                src_badge = ""
            primary_dot = '<span style="color:#4472C4;font-size:10px;margin-right:3px">●</span>' if s.get("is_primary") else ""
            rows_html += (
                f'<div class="stakeholder-row">'
                f'<div class="{av_cls}">{initials}</div>'
                f'<div style="flex:1;min-width:0">'
                f'<div style="font-size:13px;font-weight:600">{primary_dot}{s["name"]}{src_badge}</div>'
                f'<div>{title_str}</div>'
                f'<div>{email_str}</div>'
                f'<div style="margin-top:3px">{role_pills}</div>'
                f'</div></div>'
            )
        return f'<div class="cp-card" style="padding:8px 14px">{rows_html}</div>'

    col_ext, col_int = st.columns(2)

    with col_ext:
        _has_sfdc = bool(_sfdc_contacts)
        _src_note = " <span style=\"font-size:10px;font-weight:400;opacity:.5;text-transform:none;letter-spacing:0\">· SFDC matched</span>" if _has_sfdc else ""
        st.markdown(f'<div class="section-label">Customer contacts{_src_note}</div>', unsafe_allow_html=True)
        if not _merged_external:
            if df_sfdc is None:
                st.caption("Upload SFDC Contacts to see customer stakeholders, or upload a Gong doc.")
            else:
                st.markdown('<div class="no-data-msg">No contacts found for this opportunity.</div>', unsafe_allow_html=True)
        else:
            st.markdown(_build_stakeholder_html(_merged_external), unsafe_allow_html=True)

    with col_int:
        st.markdown('<div class="section-label">Zone team</div>', unsafe_allow_html=True)
        # AM from SFDC
        _sfdc_am_contacts = []
        if df_sfdc is not None and not df_sfdc.empty and _opp_id_clean and "opportunity_id" in df_sfdc.columns:
            _am_rows = df_sfdc[df_sfdc["opportunity_id"].astype(str).str.strip() == _opp_id_clean]
            if not _am_rows.empty:
                _am_row = _am_rows.iloc[0]
                _am_name = str(_am_row.get("account_manager", "") or "").strip()
                _am_email = str(_am_row.get("account_manager_email", "") or "").strip()
                if _am_name:
                    _sfdc_am_contacts.append({
                        "name": _am_name, "title": "Account Manager",
                        "email": _am_email, "roles": [], "is_primary": False,
                        "role_note": "", "source": "sfdc", "internal": True,
                    })
        _zone_team = _sfdc_am_contacts + _gong_internal
        # Deduplicate Zone team by name
        _seen_zone = set()
        _zone_dedup = []
        for _zt in _zone_team:
            if _zt["name"] not in _seen_zone:
                _seen_zone.add(_zt["name"])
                _zone_dedup.append(_zt)
        if not _zone_dedup:
            st.caption("Zone team contacts will appear here from Gong doc or SFDC.")
        else:
            st.markdown(_build_stakeholder_html(_zone_dedup, show_source=False), unsafe_allow_html=True)


# ─── TAB: REQUIREMENTS ───────────────────────────────────────────────────────
with tab_requirements:
    if d is None:
        st.markdown('<div class="no-data-msg">Upload a Gong handover doc to view Requirements.</div>',
                    unsafe_allow_html=True)
    else:
        req = d["requirements"]
        must = req.get("must", []) if isinstance(req, dict) else []
        nice = req.get("nice", []) if isinstance(req, dict) else []
        if not must and not nice:
            st.markdown('<div class="no-data-msg">No requirements parsed.</div>', unsafe_allow_html=True)
        else:
            col_must, col_nice = st.columns(2)
            with col_must:
                st.markdown(
                    f'<div class="section-label">Must-have <span style="font-size:11px;font-weight:400;opacity:.5;text-transform:none;letter-spacing:0">({len(must)})</span></div>',
                    unsafe_allow_html=True)
                rows = ""
                for item, flagged in [_strip_flag(r) for r in must]:
                    flag_html = '<span class="cp-flag">review</span>' if flagged else ""
                    rows += f'<div class="req-row">{item}{flag_html}</div>'
                st.markdown(f'<div class="cp-card" style="padding:10px 14px">{rows}</div>',
                            unsafe_allow_html=True)
            with col_nice:
                st.markdown(
                    f'<div class="section-label">Nice-to-have <span style="font-size:11px;font-weight:400;opacity:.5;text-transform:none;letter-spacing:0">({len(nice)})</span></div>',
                    unsafe_allow_html=True)
                rows = ""
                for item, flagged in [_strip_flag(r) for r in nice]:
                    flag_html = '<span class="cp-flag">review</span>' if flagged else ""
                    rows += f'<div class="req-nice req-row">{item}{flag_html}</div>'
                st.markdown(f'<div class="cp-card" style="padding:10px 14px">{rows}</div>',
                            unsafe_allow_html=True)


# ─── TAB: USE CASES ──────────────────────────────────────────────────────────
with tab_usecases:
    if d is None:
        st.markdown('<div class="no-data-msg">Upload a Gong handover doc to view Use Cases.</div>',
                    unsafe_allow_html=True)
    else:
        if not d["use_cases"]:
            st.markdown('<div class="no-data-msg">No use cases parsed.</div>', unsafe_allow_html=True)
        else:
            uc_html = ""
            for text, flagged in d["use_cases"]:
                flag_html = '<span class="cp-flag">review</span>' if flagged else ""
                uc_html += f'<div class="cp-bullet">{text}{flag_html}</div>'
            st.markdown(f'<div class="cp-card" style="padding:10px 14px">{uc_html}</div>',
                        unsafe_allow_html=True)

# ─── TAB: RISKS & COMMITMENTS ────────────────────────────────────────────────
with tab_risks:
    if d is None:
        st.markdown('<div class="no-data-msg">Upload a Gong handover doc to view Risks & Commitments.</div>',
                    unsafe_allow_html=True)
    else:
        _PRODUCT_KW = [
            "solution", "zone capture", "zone approvals", "capture tool",
            "approval workflow", "non-netsuite users", "auto-populate",
            "gen ai", "ocr", "out-of-the-box", "capabilities",
            "vendor payments directly from netsuite", "gl accounts tagged",
            "invoice capture", "participate in the approval", "pending approval",
            "existing approval workflows", "needing a netsuite license",
            "standard implementation", "premium implementation",
            "configur", "integrat", "suiteapp",
        ]
        _COMMERCIAL_KW = [
            "discount", "pricing", "price", "% discount", "usd", "free add-on",
            "tiering breakdown", "work closely with", "q1", "q2", "threshold",
            "approved discount", "same 50%", "35% approved", "advanced 1000",
            "higher tiers", "50% license", "50% discount on licens",
        ]
        _PROCESS_KW = [
            "send general material", "send email timeslots", "send timeslots",
            "refresh the team", "ahead of the demo", "first week of february",
            "last week of january", "ready for the wednesday",
            "share notes and the agenda", "session is recorded",
            "will be distributed", "provide a quote and estimate",
        ]

        def _classify_commit(text):
            tl = text.lower()
            if any(kw in tl for kw in _PRODUCT_KW):
                return "ps"
            if any(kw in tl for kw in _COMMERCIAL_KW):
                return "commercial"
            if any(kw in tl for kw in _PROCESS_KW):
                return "process"
            return "ps"

        def _render_commit_list(commits, muted=False):
            html = ""
            for c in commits:
                if c["status"] == "aligned":
                    icon = '<span style="color:#27AE60;font-size:15px">✓</span>'
                elif c["status"] == "review":
                    icon = '<span style="color:#D68910;font-size:15px">⚠</span>'
                else:
                    icon = '<span style="color:#C0392B;font-size:15px">✕</span>'
                opacity = ' style="opacity:.5"' if muted else (' style="opacity:.6"' if c["status"] == "risk" else "")
                html += (f'<div style="display:flex;gap:10px;align-items:flex-start;padding:5px 0;'
                         f'border-bottom:0.5px solid rgba(128,128,128,.1);font-size:13px;line-height:1.6">'
                         f'{icon}<span{opacity}>{c["text"]}</span></div>')
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
                label, css = badge_map.get(r["badge"], ("Risk", "risk-tech"))
                flag_html = '<span class="cp-flag">review</span>' if r["flagged"] else ""
                html += (f'<div style="display:flex;gap:8px;align-items:flex-start;padding:6px 0;'
                         f'border-bottom:0.5px solid rgba(128,128,128,.1);font-size:13px;line-height:1.6">'
                         f'<span class="risk-badge {css}">{label}</span>'
                         f'<span>{r["text"]}{flag_html}</span></div>')
            return html

        ps_commits   = [c for c in d["commitments"] if _classify_commit(c["text"]) == "ps"]
        comm_commits = [c for c in d["commitments"] if _classify_commit(c["text"]) == "commercial"]
        proc_commits = [c for c in d["commitments"] if _classify_commit(c["text"]) == "process"]
        tech_risks   = [r for r in d["risks"] if r["badge"] == "tech"]
        exp_risks    = [r for r in d["risks"] if r["badge"] == "exp"]
        org_risks    = [r for r in d["risks"] if r["badge"] in ("org", "timeline")]

        legend = """<div style="font-size:11px;opacity:.45;margin-top:6px">
            ✓ aligned &nbsp;·&nbsp; ⚠ needs review &nbsp;·&nbsp; ✕ risk / flag
        </div>"""

        sub_commits, sub_risks = st.tabs(["Commitments", "Risks"])

        with sub_commits:
            st.markdown('<div class="section-label">PS & implementation</div>', unsafe_allow_html=True)
            if ps_commits:
                st.markdown(f'<div class="cp-card" style="padding:10px 14px">{_render_commit_list(ps_commits)}</div>',
                            unsafe_allow_html=True)
            else:
                st.caption("None parsed.")
            if comm_commits:
                st.markdown('<div class="section-label" style="margin-top:16px">Commercial</div>',
                            unsafe_allow_html=True)
                st.markdown(f'<div class="cp-card" style="padding:10px 14px;opacity:.75">{_render_commit_list(comm_commits, muted=True)}</div>',
                            unsafe_allow_html=True)
            if proc_commits:
                st.markdown('<div class="section-label" style="margin-top:16px">Process & scheduling</div>',
                            unsafe_allow_html=True)
                st.markdown(f'<div class="cp-card" style="padding:10px 14px;opacity:.65">{_render_commit_list(proc_commits, muted=True)}</div>',
                            unsafe_allow_html=True)
            if d["commitments"]:
                st.markdown(legend, unsafe_allow_html=True)

        with sub_risks:
            if not d["risks"]:
                st.markdown('<div class="no-data-msg">No risks parsed.</div>', unsafe_allow_html=True)
            else:
                if tech_risks:
                    st.markdown('<div class="section-label">Technical complexity</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="cp-card" style="padding:10px 14px">{_render_risk_list(tech_risks)}</div>',
                                unsafe_allow_html=True)
                if exp_risks:
                    st.markdown('<div class="section-label" style="margin-top:16px">Expectation alignment</div>',
                                unsafe_allow_html=True)
                    st.markdown(f'<div class="cp-card" style="padding:10px 14px">{_render_risk_list(exp_risks)}</div>',
                                unsafe_allow_html=True)
                if org_risks:
                    st.markdown('<div class="section-label" style="margin-top:16px">Org readiness & timeline</div>',
                                unsafe_allow_html=True)
                    st.markdown(f'<div class="cp-card" style="padding:10px 14px">{_render_risk_list(org_risks)}</div>',
                                unsafe_allow_html=True)


# ─── TAB: NOTES ──────────────────────────────────────────────────────────────
with tab_notes:
    # Session-keyed notes: stored as list of dicts {author, ts, text}
    _notes_key = f"cp_notes_{selected_customer}"
    if _notes_key not in st.session_state:
        st.session_state[_notes_key] = []

    _notes = st.session_state[_notes_key]

    # New note input
    st.markdown('<div class="section-label">Add a note</div>', unsafe_allow_html=True)
    _note_text = st.text_area(
        "Note",
        placeholder="e.g. Maarten confirmed NS access is pending. Session 2 scheduled for Apr 22.",
        height=90,
        label_visibility="collapsed",
        key=f"cp_note_input_{selected_customer}"
    )
    if st.button("Save note", key=f"cp_note_save_{selected_customer}"):
        if _note_text.strip():
            from datetime import datetime as _dt
            st.session_state[_notes_key].insert(0, {
                "author": _session_name,
                "ts": _dt.now().strftime("%b %d, %Y · %I:%M %p"),
                "text": _note_text.strip()
            })
            st.rerun()

    # Display existing notes
    if _notes:
        st.markdown('<div class="section-label" style="margin-top:16px">Notes</div>',
                    unsafe_allow_html=True)
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
        st.markdown('<div class="no-data-msg">No notes yet — add one above.</div>',
                    unsafe_allow_html=True)

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
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# ── Remove doc + Footer ───────────────────────────────────────────────────────
_rcol1, _rcol2, _rcol3 = st.columns([3, 1, 3])
with _rcol2:
    if st.button("Remove doc", type="secondary", use_container_width=True):
        st.session_state["cp_gong_docs"][selected_customer] = []
        st.rerun()
st.markdown("""
<div style="font-size:11px;opacity:.35;text-align:center;margin-top:16px">
    PS Projects & Tools · Internal use only · Data loaded this session only
</div>
""", unsafe_allow_html=True)
