"""
Zone & Co brand Excel formatter for PS Revenue Report.
Matches Utilization Report styling exactly:
  - Font: Manrope throughout
  - Header banner: Navy #1E2C63, white text 16pt bold
  - Subtitle: grey #808080 10pt
  - Section headers: Navy fill, white text 11pt bold
  - Table headers: Blue #4472C4 fill, white text 9pt bold
  - Table rows: alternating white / #F2F2F2, Manrope 10pt
  - Metric values: Navy #1E2C63 14pt bold; alert red #E74C3C on #FDECED
"""

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# ── Brand colours ──────────────────────────────────────────────────────────────
NAVY       = "1E2C63"
TEAL       = "08A9B7"
BLUE       = "4472C4"
WHITE      = "FFFFFF"
GREY       = "808080"
LIGHT_GREY = "F2F2F2"
RED        = "E74C3C"
RED_BG     = "FDECED"
ORANGE     = "FF4B40"

def _font(name="Manrope", size=10, bold=False, color="000000"):
    return Font(name=name, size=size, bold=bold, color=color)

def _fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def _center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def _left():
    return Alignment(horizontal="left", vertical="center", wrap_text=False)

def _thin_border():
    s = Side(style="thin", color="D0D0D0")
    return Border(bottom=s)


def format_header_banner(ws, title, subtitle, blurb, start_row=2):
    """Write title banner, subtitle and blurb rows."""
    # Title row — navy fill, white bold 16pt
    cell = ws.cell(row=start_row, column=2, value=title)
    cell.font      = _font(size=16, bold=True, color=WHITE)
    cell.fill      = _fill(NAVY)
    cell.alignment = _left()
    ws.row_dimensions[start_row].height = 28

    # Subtitle — grey 10pt
    sub = ws.cell(row=start_row+1, column=2, value=subtitle)
    sub.font = _font(size=10, color=GREY)

    # Blurb — grey 9pt
    bl = ws.cell(row=start_row+2, column=2, value=blurb)
    bl.font      = _font(size=9, color=GREY)
    bl.alignment = Alignment(wrap_text=True)
    ws.row_dimensions[start_row+2].height = 36


def format_section_header(ws, row, col, text, n_cols=6):
    """Navy fill section header spanning n_cols columns."""
    cell = ws.cell(row=row, column=col, value=text)
    cell.font      = _font(size=11, bold=True, color=WHITE)
    cell.fill      = _fill(NAVY)
    cell.alignment = _left()
    ws.row_dimensions[row].height = 20
    # Fill neighbouring cells in same row for visual span
    for c in range(col+1, col+n_cols):
        nc = ws.cell(row=row, column=c)
        nc.fill = _fill(NAVY)


def format_metric_card(ws, row, col, label, value, alert=False):
    """Label row + value row metric card."""
    lbl = ws.cell(row=row, column=col, value=label)
    lbl.font = _font(size=10, color=GREY)

    val = ws.cell(row=row+1, column=col, value=value)
    if alert:
        val.font      = _font(size=14, bold=True, color=RED)
        val.fill      = _fill(RED_BG)
        val.alignment = _center()
    else:
        val.font = _font(size=14, bold=True, color=NAVY)


def format_table_headers(ws, row, headers, start_col=1):
    """Blue table header row."""
    for i, h in enumerate(headers):
        cell = ws.cell(row=row, column=start_col+i, value=h)
        cell.font      = _font(size=9, bold=True, color=WHITE)
        cell.fill      = _fill(BLUE)
        cell.alignment = _center()
    ws.row_dimensions[row].height = 18


def format_table_rows(ws, data_start_row, n_rows, n_cols, start_col=1,
                      currency_cols=None, pct_cols=None):
    """Alternating white/light-grey rows with Manrope 10pt."""
    currency_cols = currency_cols or []
    pct_cols      = pct_cols or []
    for r in range(data_start_row, data_start_row + n_rows):
        fill = _fill(LIGHT_GREY) if (r - data_start_row) % 2 == 1 else _fill(WHITE)
        for c in range(start_col, start_col + n_cols):
            cell = ws.cell(row=r, column=c)
            cell.font      = _font(size=10, color="000000")
            cell.fill      = fill
            cell.alignment = _left()
            col_idx = c - start_col
            if col_idx in currency_cols:
                cell.number_format = '$#,##0.00'
            if col_idx in pct_cols:
                cell.number_format = '0.0%'


def format_data_sheet(ws, title_text):
    """Apply brand formatting to a data sheet with a simple header."""
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"

    # Title banner in row 1 across all used columns
    max_col = ws.max_column or 1
    cell = ws.cell(row=1, column=1, value=title_text)
    cell.font      = _font(size=12, bold=True, color=WHITE)
    cell.fill      = _fill(NAVY)
    cell.alignment = _left()
    ws.row_dimensions[1].height = 20

    for c in range(2, max_col + 1):
        ws.cell(row=1, column=c).fill = _fill(NAVY)

    # Header row (row 2) — blue
    if ws.max_row >= 2:
        for cell in ws[2]:
            if cell.value is not None or True:
                cell.font      = _font(size=9, bold=True, color=WHITE)
                cell.fill      = _fill(BLUE)
                cell.alignment = _center()
        ws.row_dimensions[2].height = 18

    # Data rows — alternating
    for r_idx, row in enumerate(ws.iter_rows(min_row=3, max_row=ws.max_row)):
        fill = _fill(LIGHT_GREY) if r_idx % 2 == 1 else _fill(WHITE)
        for cell in row:
            cell.font      = _font(size=10)
            cell.fill      = fill
            cell.alignment = _left()

    # Auto-width columns (capped at 40)
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                max_len = max(max_len, len(str(cell.value or "")))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 8), 40)


def build_dashboard(wb, metrics: dict, date_str: str, blurb: str):
    """Build Dashboard sheet matching Utilization report layout."""
    if "Dashboard" in wb.sheetnames:
        del wb["Dashboard"]
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2   # left margin
    ws.column_dimensions["B"].width = 22
    for col in ["C","D","E","F","G","H"]:
        ws.column_dimensions[col].width = 18

    # ── Header banner ─────────────────────────────────────────────────────────
    format_header_banner(
        ws,
        title    = "Professional Services — Revenue Report",
        subtitle = f"Data through {date_str}",
        blurb    = blurb,
        start_row = 2
    )

    # ── EXECUTIVE SUMMARY ─────────────────────────────────────────────────────
    format_section_header(ws, row=5, col=2, text="EXECUTIVE SUMMARY", n_cols=7)

    exec_metrics = [
        ("YTD Revenue",          metrics.get("ytd"),       False),
        ("QTD Revenue",          metrics.get("qtd"),       False),
        ("MTD Revenue",          metrics.get("mtd"),       False),
        ("Full Month Forecast",  metrics.get("full_mo"),   False),
        ("Run Rate (ARR)",       metrics.get("run_rate"),  False),
        ("MoM Growth",           metrics.get("mom"),       False),
    ]
    for i, (label, value, alert) in enumerate(exec_metrics):
        format_metric_card(ws, row=6, col=2+i, label=label, value=value, alert=alert)
    ws.row_dimensions[6].height = 16
    ws.row_dimensions[7].height = 22

    # ── FF / T&M / Reconcile summary row ──────────────────────────────────────
    format_section_header(ws, row=9, col=2, text="REVENUE BY TYPE", n_cols=7)
    type_metrics = [
        ("FF Revenue YTD",        metrics.get("ff_ytd"),      False),
        ("T&M Revenue YTD",       metrics.get("tm_ytd"),      False),
        ("Reconcile Carve YTD",   metrics.get("recon_ytd"),   False),
        ("FF Projects",           metrics.get("ff_projects"), False),
        ("T&M Projects",          metrics.get("tm_projects"), False),
        ("Carve Flags",           metrics.get("flag_count"),
         (metrics.get("flag_count") or 0) > 0),
    ]
    for i, (label, value, alert) in enumerate(type_metrics):
        format_metric_card(ws, row=10, col=2+i, label=label, value=value, alert=alert)
    ws.row_dimensions[10].height = 16
    ws.row_dimensions[11].height = 22

    # ── ANALYST DETAIL — Revenue by Month ────────────────────────────────────
    format_section_header(ws, row=13, col=2, text="ANALYST DETAIL — Revenue by Month (USD)", n_cols=7)
    trend_data = metrics.get("trend_rows", [])
    if trend_data:
        hdrs = list(trend_data[0].keys())
        format_table_headers(ws, row=14, headers=hdrs, start_col=2)
        for r_idx, row_data in enumerate(trend_data):
            fill = _fill(LIGHT_GREY) if r_idx % 2 == 1 else _fill(WHITE)
            for c_idx, val in enumerate(row_data.values()):
                cell = ws.cell(row=15+r_idx, column=2+c_idx, value=val)
                cell.font  = _font(size=10)
                cell.fill  = fill
                cell.alignment = _left()
        last_trend_row = 15 + len(trend_data)
    else:
        last_trend_row = 15

    # ── By Region + By Product ────────────────────────────────────────────────
    format_section_header(ws, row=last_trend_row+1, col=2,
                          text="BY REGION — YTD", n_cols=3)
    format_section_header(ws, row=last_trend_row+1, col=5,
                          text="BY PRODUCT — YTD", n_cols=4)

    region_rows  = metrics.get("region_rows", [])
    product_rows = metrics.get("product_rows", [])

    if region_rows:
        hdrs_r = list(region_rows[0].keys())
        format_table_headers(ws, row=last_trend_row+2, headers=hdrs_r, start_col=2)
        for r_idx, rd in enumerate(region_rows):
            fill = _fill(LIGHT_GREY) if r_idx % 2 == 1 else _fill(WHITE)
            for c_idx, val in enumerate(rd.values()):
                cell = ws.cell(row=last_trend_row+3+r_idx, column=2+c_idx, value=val)
                cell.font = _font(size=10); cell.fill = fill; cell.alignment = _left()

    if product_rows:
        hdrs_p = list(product_rows[0].keys())
        format_table_headers(ws, row=last_trend_row+2, headers=hdrs_p, start_col=5)
        for r_idx, pd_row in enumerate(product_rows):
            fill = _fill(LIGHT_GREY) if r_idx % 2 == 1 else _fill(WHITE)
            for c_idx, val in enumerate(pd_row.values()):
                cell = ws.cell(row=last_trend_row+3+r_idx, column=5+c_idx, value=val)
                cell.font = _font(size=10); cell.fill = fill; cell.alignment = _left()

    return ws


def apply_zone_formatting(excel_bytes: bytes, metrics: dict,
                           date_str: str, blurb: str) -> bytes:
    """
    Take raw xlsxwriter bytes, load with openpyxl,
    add Dashboard tab and apply Zone brand formatting to all sheets.
    Returns formatted bytes.
    """
    wb = load_workbook(io.BytesIO(excel_bytes))

    # Build Dashboard first (inserted at position 0)
    build_dashboard(wb, metrics, date_str, blurb)

    # Format all data sheets
    skip = {"Dashboard"}
    for sheet_name in wb.sheetnames:
        if sheet_name in skip:
            continue
        ws = wb[sheet_name]
        # Insert a blank row 1 for the title banner, shift existing data down
        ws.insert_rows(1)
        ws.insert_rows(1)
        format_data_sheet(ws, sheet_name)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()
