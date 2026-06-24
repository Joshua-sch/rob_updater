import streamlit as st
import pandas as pd
import openpyxl
import io
import re
import datetime

# ── CSV parsing ───────────────────────────────────────────────────────────────

MONTH_ABBR = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
DAILY_RE = re.compile(r"^\d{2}-\d{2}-\d{4}")


def is_formula(value) -> bool:
    return isinstance(value, str) and value.strip().startswith("=")


def parse_csv(file_bytes: bytes) -> pd.DataFrame:
    raw = pd.read_csv(io.BytesIO(file_bytes), header=None, dtype=str, encoding="utf-8-sig")
    df = raw.iloc[2:].reset_index(drop=True)
    df.columns = range(df.shape[1])
    return df


def classify_row(date_str: str):
    """Return ('daily', date) | ('monthly', (year, month)) | (None, None)"""
    if not isinstance(date_str, str):
        return None, None
    date_str = date_str.strip()
    if DAILY_RE.match(date_str):
        try:
            d = datetime.datetime.strptime(date_str[:10], "%m-%d-%Y").date()
            return "daily", d
        except ValueError:
            return None, None
    parts = date_str.split()
    if len(parts) == 2 and parts[0][:3].lower() in MONTH_ABBR:
        try:
            month = MONTH_ABBR[parts[0][:3].lower()]
            year = int(parts[1])
            return "monthly", (year, month)
        except ValueError:
            pass
    return None, None


def safe_float(val):
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


# ── ROB Update ───────────────────────────────────────────────────────────────

ROB_SHEETS = ["wk one", "wk two", "wk three", "wk four", "wk five", "wk six"]


def find_secondary_col(ws, block_start):
    candidates = []
    for cell in ws[block_start]:
        if cell.column <= 5:
            continue
        if isinstance(cell.value, str) and "variance" in cell.value.strip().lower():
            candidates.append(cell.column)
    return min(candidates) if candidates else None


def build_rob_change_plan(df, ws):
    today = datetime.date.today()
    current_month = today.month
    current_year = today.year
    changes = []

    # E4 = as-of date
    changes.append({
        "row": 4, "col": 5, "label": "As-of date", "month": None,
        "new_value": today, "skip_reason": None,
    })

    for _, row in df.iterrows():
        date_str = str(row[0]).strip() if row[0] else ""
        kind, info = classify_row(date_str)
        if kind != "monthly":
            continue
        year, month = info
        if year != current_year or month < current_month:
            continue

        month_index = month - 1
        block_start = 4 + 8 * month_index

        rev     = safe_float(row[5])
        rms     = safe_float(row[1])
        grp_pu  = safe_float(row[7])
        grp_npu = safe_float(row[8])
        grp_rvn = safe_float(row[9])

        grp_sold = (grp_pu or 0) + (grp_npu or 0) if grp_pu is not None and grp_npu is not None else None
        sec_col = find_secondary_col(ws, block_start)

        entries = [
            (block_start + 1, 5, "Revenue",        rev,       False),
            (block_start + 2, 5, "Room Nights",     rms,       False),
            (block_start + 4, 5, "Group Rms Sold",  grp_sold,  False),
            (block_start + 5, 5, "Group Rm Rev",    grp_rvn,   False),
        ]
        if sec_col:
            from openpyxl.utils import get_column_letter
            sec_letter = get_column_letter(sec_col)
            npu_row    = block_start + 4
            adr_row    = block_start + 6
            npu_formula = f"={sec_letter}{npu_row}*E{adr_row}"
            entries.append((npu_row,     sec_col, "Group Not P/U rooms",    grp_npu,     False))
            entries.append((npu_row + 1, sec_col, "Group Not P/U rev (formula)", npu_formula, True))

        for r, c, label, val, is_formula_write in entries:
            skip = None
            if r >= 100:
                skip = "row≥100"
            elif not is_formula_write and is_formula(ws.cell(r, c).value):
                skip = "formula"
            changes.append({"row": r, "col": c, "label": label, "month": month,
                             "new_value": val, "skip_reason": skip})

    return changes


def apply_rob_changes(wb, sheet_name, changes):
    ws = wb[sheet_name]
    for ch in changes:
        if ch["skip_reason"]:
            continue
        ws.cell(ch["row"], ch["col"]).value = ch["new_value"]


def first_uncolored_sheet(wb, sheet_names):
    """Return the first sheet in sheet_names whose tab has no color set."""
    for name in sheet_names:
        ws = wb[name]
        tc = ws.sheet_properties.tabColor
        if tc is None:
            return name
    return sheet_names[-1]  # fallback: last sheet


def color_tab_done(wb, sheet_name):
    """Mark a sheet tab green to indicate it has been completed."""
    from openpyxl.styles.colors import Color
    wb[sheet_name].sheet_properties.tabColor = Color(rgb="FF00B050")


def strip_tables(wb):
    """Remove Excel table definitions to prevent openpyxl save corruption."""
    for ws in wb.worksheets:
        ws.tables.clear()


# ── Strategy Report ───────────────────────────────────────────────────────────

STRATEGY_SHEETS = ["WKONE", "WKTWO", "WKTHREE", "WKFOUR", "WKFIVE"]
FORECAST_SHEETS = ["FCST-WK1", "FCST-WK2", "FCST-WK3", "FCST-WK4",
                   "FCST-WK5", "FCST-WK6", "FCST-WK7", "FCST-WK8", "FCST-WK9"]

# openpyxl col index → (header label, CSV col index)
STRATEGY_COLS = {
    4:  ("OTB TY Trans (Indiv Count)", 15),
    9:  ("GRP PU TY",                   7),
    14: ("GRP N/PU TY",                 8),
    21: ("OOO RMS",                     4),
    35: ("Trans Rev TY",               16),
    43: ("Grp Rev TY",                  9),
}


def build_date_row_map(wb):
    ws = wb["WKONE"]
    mapping = {}
    for row_num in range(5, ws.max_row + 1):
        val = ws.cell(row_num, 3).value
        if isinstance(val, datetime.datetime):
            mapping[val.date()] = row_num
        elif isinstance(val, datetime.date):
            mapping[val] = row_num
    return mapping


def build_strategy_change_plan(df, wb, sheet_name):
    today = datetime.date.today()
    scope_start = today.replace(day=1)
    scope_end = datetime.date(today.year, 12, 31)

    date_row_map = build_date_row_map(wb)
    ws = wb[sheet_name]
    changes = []

    for _, row in df.iterrows():
        date_str = str(row[0]).strip() if row[0] else ""
        kind, info = classify_row(date_str)
        if kind != "daily":
            continue
        d = info
        if d < scope_start or d > scope_end:
            continue
        if d not in date_row_map:
            continue

        excel_row = date_row_map[d]
        for excel_col, (label, csv_col) in STRATEGY_COLS.items():
            val = safe_float(row[csv_col])
            skip = "formula" if is_formula(ws.cell(excel_row, excel_col).value) else None
            changes.append({
                "date": d, "row": excel_row, "col": excel_col,
                "label": label, "new_value": val, "skip_reason": skip,
            })

    return changes


def apply_strategy_changes(wb, sheet_name, changes):
    ws = wb[sheet_name]
    for ch in changes:
        if ch["skip_reason"]:
            continue
        ws.cell(ch["row"], ch["col"]).value = ch["new_value"]


def parse_rate_csv(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_csv(io.BytesIO(file_bytes), dtype=str, encoding="utf-8-sig")
    df.columns = [c.strip() for c in df.columns]
    return df


def find_header_col(ws, keyword, header_rows=(2, 3, 4)):
    """Find the leftmost column whose concatenated header text (rows 2-4) contains keyword."""
    keyword = keyword.strip().lower()
    col_texts = {}
    for row_num in header_rows:
        for cell in ws[row_num]:
            col = cell.column
            if col not in col_texts:
                col_texts[col] = ""
            if isinstance(cell.value, str):
                col_texts[col] += cell.value.strip()
    matches = [col for col, text in col_texts.items() if keyword in text.lower()]
    return min(matches) if matches else None


def build_rates_change_plan(rate_df, wb, sheet_name):
    today = datetime.date.today()
    scope_start = today.replace(day=1)
    scope_end = datetime.date(today.year, 12, 31)

    date_row_map = build_date_row_map(wb)
    ws = wb[sheet_name]

    restric_col = find_header_col(ws, "restric")
    hotel_col   = (restric_col + 1) if restric_col else None

    changes = []
    warnings = []
    if not restric_col:
        warnings.append("Could not find Restrictions column in sheet headers.")

    for _, row in rate_df.iterrows():
        date_str = str(row.get("Date", "")).strip()
        d = None
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
            try:
                d = datetime.datetime.strptime(date_str, fmt).date()
                break
            except ValueError:
                continue
        if d is None:
            continue
        if d < scope_start or d > scope_end:
            continue
        if d not in date_row_map:
            continue

        excel_row  = date_row_map[d]
        double_val = safe_float(row.get("Double", ""))
        mlos_val   = safe_float(row.get("Min Length of Stay", ""))

        if hotel_col:
            skip = "formula" if is_formula(ws.cell(excel_row, hotel_col).value) else None
            changes.append({"date": d, "row": excel_row, "col": hotel_col,
                            "label": "Hotel Rate", "new_value": double_val, "skip_reason": skip})
        if restric_col:
            skip = "formula" if is_formula(ws.cell(excel_row, restric_col).value) else None
            changes.append({"date": d, "row": excel_row, "col": restric_col,
                            "label": "Restrictions (MLOS)", "new_value": mlos_val, "skip_reason": skip})

    return changes, warnings


# ── Forecast helpers ─────────────────────────────────────────────────────────

DATE_FORMATS = [
    "%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d", "%m/%d/%y", "%m-%d-%y",
    "%B %d, %Y", "%b %d, %Y", "%B %d %Y", "%b %d %Y",
    "%d-%b-%Y", "%d/%b/%Y",
]

def parse_any_date(val):
    """Parse a date from a datetime object, int serial, or string in many formats."""
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.date() if isinstance(val, datetime.datetime) else val
    if isinstance(val, (int, float)):
        # Excel serial date (days since 1900-01-01, with Lotus bug offset)
        return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(val)))
    if isinstance(val, str):
        val = val.strip()
        for fmt in DATE_FORMATS:
            try:
                return datetime.datetime.strptime(val, fmt).date()
            except ValueError:
                continue
    return None


def build_forecast_date_col_map(ws, wb=None):
    """Return {date: col_index} from row 4. Falls back to WK1 for formula-only sheets."""
    month_start = parse_any_date(ws.cell(4, 2).value)

    # If this sheet's col B is a formula, find the start date from any WK sheet with a literal
    if month_start is None and wb is not None:
        for sname in wb.sheetnames:
            if "glance" in sname.lower():
                continue
            candidate = parse_any_date(wb[sname].cell(4, 2).value)
            if candidate is not None:
                month_start = candidate
                break

    if month_start is None:
        return {}

    col_map = {}
    col = 2
    while col <= ws.max_column:
        cell = ws.cell(4, col)
        if isinstance(cell.value, str) and "total" in cell.value.lower():
            break
        if cell.value is None and col > 2:
            break
        col_map[month_start + datetime.timedelta(days=col - 2)] = col
        col += 1
    return col_map


def row_is_filled(ws, r):
    """Return True if cols B-AF contain actual numbers OR cross-sheet formula refs."""
    for col in range(2, 33):
        v = ws.cell(r, col).value
        if isinstance(v, (int, float)):
            return True
        if isinstance(v, str) and v.startswith("='"):
            return True
    return False


def find_next_pickup_data_row(ws):
    """Find the next available row in the pick-up tracking chart.

    Strategy:
    1. Locate the 'Day of Week' section header (last one before 'Total Pick UP').
    2. Within the section, find the last row that has actual numbers in B-AF
       (cross-sheet formula refs look empty to openpyxl, so only direct numbers count).
    3. Scan forward from there for the first row with no numbers and no skip keyword
       — works for WK1/WK2 (back-to-back entries) and WK3+ (Pick UP rows between).
    """
    skip_keywords = {"pick", "day", "total", "forecast", "budget", "last year"}

    # Find Total Pick UP boundary
    total_row = None
    for r in range(40, 150):
        if "total pick" in str(ws.cell(r, 1).value or "").lower():
            total_row = r
            break
    search_end = total_row if total_row else 150

    # Find last 'Day of Week' header before the boundary
    section_start = None
    for r in range(40, search_end):
        if "day of week" in str(ws.cell(r, 1).value or "").lower():
            section_start = r
    if section_start is None:
        return None

    # Find the last row in the section that has actual numbers in B-AF
    last_filled = None
    for r in range(section_start + 2, search_end):
        if row_is_filled(ws, r):
            last_filled = r

    if last_filled is None:
        # Nothing filled yet — return first non-header row in section
        last_filled = section_start + 1

    # Scan forward from last_filled+1 for the first row with no numbers
    for r in range(last_filled + 1, search_end):
        a_val = str(ws.cell(r, 1).value or "").strip().lower()
        if any(k in a_val for k in skip_keywords):
            continue
        if not row_is_filled(ws, r):
            return r

    return None


def build_forecast_change_plan(df, ws):
    """Build list of cell writes for the Forecast sheet."""
    today = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)
    month_start = today.replace(day=1)

    col_map = build_forecast_date_col_map(ws, ws.parent)
    if not col_map:
        return [], ["Could not read date row from forecast sheet."]

    changes = []
    warnings = []

    # A2 = today's date
    changes.append({
        "label": "As-of date", "row": 2, "col": 1,
        "new_value": today, "skip_reason": "formula" if is_formula(ws.cell(2, 1).value) else None,
    })

    # Build lookup from CSV: date -> row dict
    daily_rows = {}
    for _, row in df.iterrows():
        date_str = str(row.iloc[0]).strip()
        if not re.match(r"^\d{2}-\d{2}-\d{4}", date_str):
            continue
        try:
            d = datetime.datetime.strptime(date_str[:10], "%m-%d-%Y").date()
        except ValueError:
            continue
        daily_rows[d] = row

    for d, col in col_map.items():
        if d not in daily_rows:
            continue
        csv_row = daily_rows[d]
        rms = safe_float(csv_row.iloc[1])
        adr = safe_float(csv_row.iloc[6])
        rev = safe_float(csv_row.iloc[5])

        is_future = d >= today
        is_past   = d <= yesterday and d >= month_start

        # Row 6: Rooms Sold (future)
        if is_future:
            skip = "formula" if is_formula(ws.cell(6, col).value) else None
            changes.append({"label": f"Rooms Sold (future) {d}", "row": 6, "col": col,
                            "new_value": rms, "skip_reason": skip})
            # Row 14: ADR OTB (future)
            skip = "formula" if is_formula(ws.cell(14, col).value) else None
            changes.append({"label": f"ADR OTB {d}", "row": 14, "col": col,
                            "new_value": adr, "skip_reason": skip})

        # Row 16: Rooms Sold (actuals)
        if is_past:
            skip = "formula" if is_formula(ws.cell(16, col).value) else None
            changes.append({"label": f"Rooms Sold (actual) {d}", "row": 16, "col": col,
                            "new_value": rms, "skip_reason": skip})
            # Row 19: Revenue (actuals)
            skip = "formula" if is_formula(ws.cell(19, col).value) else None
            changes.append({"label": f"Revenue (actual) {d}", "row": 19, "col": col,
                            "new_value": rev, "skip_reason": skip})

    # Pick-up tracking row: write full month rooms sold to next available row
    target_row = find_next_pickup_data_row(ws)

    if target_row:
        changes.append({"label": "Pickup tracking: date", "row": target_row, "col": 1,
                        "new_value": today, "skip_reason": None})
        for d, col in col_map.items():
            if d not in daily_rows:
                continue
            rms = safe_float(daily_rows[d].iloc[1])
            skip = "formula" if is_formula(ws.cell(target_row, col).value) else None
            changes.append({"label": f"Pickup tracking: Rooms Sold {d}",
                            "row": target_row, "col": col,
                            "new_value": rms, "skip_reason": skip})
    else:
        warnings.append("No available row found in pick-up tracking chart.")

    return changes, warnings


def build_next_month_forecast_plan(df, ws):
    """For weeks 3 & 4: write Rooms Sold (row 6) and ADR OTB (row 14) for ALL
    dates in the next month (everything in the CSV beyond the current month).
    Also writes A2 = today and the pick-up tracking row.
    """
    today = datetime.date.today()
    current_month_end = (today.replace(day=1) + datetime.timedelta(days=32)).replace(day=1) - datetime.timedelta(days=1)

    col_map = build_forecast_date_col_map(ws, ws.parent)
    if not col_map:
        return [], ["Could not read date row from next-month forecast sheet."]

    changes = []
    warnings = []

    # A2 = today's date
    changes.append({
        "label": "As-of date", "row": 2, "col": 1,
        "new_value": today,
        "skip_reason": "formula" if is_formula(ws.cell(2, 1).value) else None,
    })

    daily_rows = {}
    for _, row in df.iterrows():
        date_str = str(row.iloc[0]).strip()
        if not re.match(r"^\d{2}-\d{2}-\d{4}", date_str):
            continue
        try:
            d = datetime.datetime.strptime(date_str[:10], "%m-%d-%Y").date()
        except ValueError:
            continue
        daily_rows[d] = row

    for d, col in col_map.items():
        # Only next month dates (after current month end)
        if d <= current_month_end:
            continue
        if d not in daily_rows:
            continue
        csv_row = daily_rows[d]
        rms = safe_float(csv_row.iloc[1])
        adr = safe_float(csv_row.iloc[6])

        skip6  = "formula" if is_formula(ws.cell(6, col).value) else None
        skip14 = "formula" if is_formula(ws.cell(14, col).value) else None
        changes.append({"label": f"Rooms Sold (future) {d}", "row": 6,  "col": col, "new_value": rms, "skip_reason": skip6})
        changes.append({"label": f"ADR OTB {d}",             "row": 14, "col": col, "new_value": adr, "skip_reason": skip14})

    # Pick-up tracking row (same logic — full month of next month dates)
    target_row = find_next_pickup_data_row(ws)
    if target_row:
        changes.append({"label": "Pickup tracking: date", "row": target_row, "col": 1,
                        "new_value": today, "skip_reason": None})
        for d, col in col_map.items():
            if d not in daily_rows:
                continue
            rms = safe_float(daily_rows[d].iloc[1])
            skip = "formula" if is_formula(ws.cell(target_row, col).value) else None
            changes.append({"label": f"Pickup tracking: Rooms Sold {d}",
                            "row": target_row, "col": col, "new_value": rms, "skip_reason": skip})
    else:
        warnings.append("No available row found in next-month pick-up tracking chart.")

    return changes, warnings


def apply_forecast_changes(wb, sheet_name, changes):
    ws = wb[sheet_name]
    for ch in changes:
        if ch["skip_reason"]:
            continue
        ws.cell(ch["row"], ch["col"]).value = ch["new_value"]


# ── UI ────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Linchris Weekly Tools", layout="wide")

st.markdown("""
<style>
  .stTabs [data-baseweb="tab-list"] { gap: 8px; }
  .stTabs [data-baseweb="tab"] {
    background: #1C2D4E; color: #C9A84C;
    border-radius: 6px 6px 0 0; padding: 8px 20px; font-weight: 600;
  }
  .stTabs [aria-selected="true"] { background: #C9A84C !important; color: #0F1B2D !important; }
  div[data-testid="metric-container"] { background: #1C2D4E; border-radius: 8px; padding: 12px; }
</style>
""", unsafe_allow_html=True)

st.title("Linchris Hotel Corporation — Weekly Update Tools")

tab_rob, tab_strategy, tab_forecast = st.tabs(["ROB Update", "Strategy Report", "Forecast"])

# ── Tab 1: ROB ────────────────────────────────────────────────────────────────
with tab_rob:
    st.header("ROB Master Workbook Update")
    csv_file = st.file_uploader("Upload CSV (Business on the Books)", type=["csv"], key="rob_csv")
    xl_file  = st.file_uploader("Upload ROB Master Workbook (.xlsx)", type=["xlsx"], key="rob_xl")

    if csv_file and xl_file:
        csv_bytes = csv_file.read()
        xl_bytes  = xl_file.read()

        df = parse_csv(csv_bytes)
        wb = openpyxl.load_workbook(io.BytesIO(xl_bytes), data_only=False)

        auto_sheet = first_uncolored_sheet(wb, ROB_SHEETS)
        sheet_choice = st.selectbox("Week tab", ROB_SHEETS,
                                    index=ROB_SHEETS.index(auto_sheet), key="rob_sheet")
        st.caption(f"Auto-detected next tab: **{auto_sheet}**")

        if st.button("Preview Changes", key="rob_preview"):
            ws = wb[sheet_choice]
            changes = build_rob_change_plan(df, ws)
            st.session_state["rob_changes"]   = changes
            st.session_state["rob_wb_bytes"]  = xl_bytes
            st.session_state["rob_sheet_sel"] = sheet_choice

        if "rob_changes" in st.session_state:
            changes    = st.session_state["rob_changes"]
            will_write = [c for c in changes if not c["skip_reason"]]
            skipped    = [c for c in changes if c["skip_reason"]]

            c1, c2 = st.columns(2)
            c1.metric("Cells to update", len(will_write))
            c2.metric("Skipped",         len(skipped))

            preview_rows = []
            for c in changes:
                preview_rows.append({
                    "Month":  c["month"] or "—",
                    "Label":  c["label"],
                    "Row":    c["row"],
                    "Col":    c["col"],
                    "Value":  c["new_value"],
                    "Status": "✅ will write" if not c["skip_reason"] else f"⚠️ skip ({c['skip_reason']})",
                })
            st.dataframe(preview_rows, use_container_width=True)

            if st.button("Confirm and Apply Changes", key="rob_apply"):
                wb2 = openpyxl.load_workbook(io.BytesIO(st.session_state["rob_wb_bytes"]), data_only=False)
                apply_rob_changes(wb2, st.session_state["rob_sheet_sel"], changes)
                color_tab_done(wb2, st.session_state["rob_sheet_sel"])
                strip_tables(wb2)
                out = io.BytesIO()
                wb2.save(out)
                st.download_button(
                    "Download Updated ROB Workbook",
                    data=out.getvalue(),
                    file_name="ROB_Master_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ── Tab 2: Strategy ───────────────────────────────────────────────────────────
with tab_strategy:
    st.header("Strategy Report Update")

    col_a, col_b = st.columns(2)
    with col_a:
        csv_file2 = st.file_uploader("Upload CSV (Business on the Books)", type=["csv"], key="str_csv")
    with col_b:
        rate_file2 = st.file_uploader("Upload Rates & Restrictions CSV", type=["csv"], key="str_rate")

    col_c, col_d = st.columns(2)
    with col_c:
        xl_file2 = st.file_uploader("Upload Strategy Report Workbook (.xlsx)", type=["xlsx"], key="str_xl")

    if xl_file2:
        xl_bytes2 = xl_file2.read()
        wb2_peek  = openpyxl.load_workbook(io.BytesIO(xl_bytes2), data_only=False)

        auto_sheet2   = first_uncolored_sheet(wb2_peek, STRATEGY_SHEETS)
        sheet_choice2 = st.selectbox("Week tab", STRATEGY_SHEETS,
                                     index=STRATEGY_SHEETS.index(auto_sheet2), key="str_sheet")
        st.caption(f"Auto-detected next tab: **{auto_sheet2}**")

        if (csv_file2 or rate_file2) and st.button("Preview Changes", key="str_preview"):
            wb2_full = openpyxl.load_workbook(io.BytesIO(xl_bytes2), data_only=False)
            all_changes = []

            if csv_file2:
                df2 = parse_csv(csv_file2.read())
                all_changes += build_strategy_change_plan(df2, wb2_full, sheet_choice2)

            rate_warnings = []
            if rate_file2:
                rate_df = parse_rate_csv(rate_file2.read())
                rate_changes, rate_warnings = build_rates_change_plan(
                    rate_df, wb2_full, sheet_choice2)
                all_changes += rate_changes

            st.session_state["str_changes"]   = all_changes
            st.session_state["str_wb_bytes"]  = xl_bytes2
            st.session_state["str_sheet_sel"] = sheet_choice2
            st.session_state["str_warnings"]  = rate_warnings

        if "str_changes" in st.session_state:
            for w in st.session_state.get("str_warnings", []):
                st.warning(w)

            changes2    = st.session_state["str_changes"]
            will_write2 = [c for c in changes2 if not c["skip_reason"]]
            skipped2    = [c for c in changes2 if c["skip_reason"]]

            c1, c2, c3 = st.columns(3)
            c1.metric("Cells to update", len(will_write2))
            c2.metric("Skipped",         len(skipped2))
            c3.metric("Days in scope",   len({c["date"] for c in changes2}))

            preview_rows2 = []
            for c in changes2:
                preview_rows2.append({
                    "Date":   str(c["date"]),
                    "Label":  c["label"],
                    "Row":    c["row"],
                    "Col":    c["col"],
                    "Value":  c["new_value"],
                    "Status": "✅ will write" if not c["skip_reason"] else f"⚠️ skip ({c['skip_reason']})",
                })
            st.dataframe(preview_rows2, use_container_width=True)

            if st.button("Confirm and Apply Changes", key="str_apply"):
                wb3 = openpyxl.load_workbook(io.BytesIO(st.session_state["str_wb_bytes"]), data_only=False)
                apply_strategy_changes(wb3, st.session_state["str_sheet_sel"], changes2)
                color_tab_done(wb3, st.session_state["str_sheet_sel"])
                strip_tables(wb3)
                out2 = io.BytesIO()
                wb3.save(out2)
                st.download_button(
                    "Download Updated Strategy Workbook",
                    data=out2.getvalue(),
                    file_name="Strategy_Report_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

# ── Tab 3: Forecast ───────────────────────────────────────────────────────────
with tab_forecast:
    st.header("Forecast Update")

    fcst_csv     = st.file_uploader("Upload CSV (Business on the Books)", type=["csv"], key="fcst_csv")
    fcst_xl      = st.file_uploader("Upload Current Month Forecast Workbook (.xlsx)", type=["xlsx"], key="fcst_xl")
    st.caption("Weeks 3 & 4 only: also upload next month's forecast workbook.")
    fcst_xl_next = st.file_uploader("Upload Next Month Forecast Workbook (.xlsx)", type=["xlsx"], key="fcst_xl_next")

    if fcst_xl:
        fcst_bytes = fcst_xl.read()
        wb_fcst_peek = openpyxl.load_workbook(io.BytesIO(fcst_bytes), data_only=False)

        avail_fcst = [s for s in FORECAST_SHEETS if s in wb_fcst_peek.sheetnames]
        auto_fcst  = first_uncolored_sheet(wb_fcst_peek, avail_fcst) if avail_fcst else None
        fcst_sheet = st.selectbox("Week tab", avail_fcst,
                                  index=avail_fcst.index(auto_fcst) if auto_fcst else 0,
                                  key="fcst_sheet")
        if auto_fcst:
            st.caption(f"Auto-detected next tab: **{auto_fcst}**")

        if fcst_xl_next:
            fcst_next_bytes = fcst_xl_next.read()
            wb_next_peek    = openpyxl.load_workbook(io.BytesIO(fcst_next_bytes), data_only=False)
            avail_next      = [s for s in FORECAST_SHEETS if s in wb_next_peek.sheetnames]
            auto_next       = first_uncolored_sheet(wb_next_peek, avail_next) if avail_next else None
            fcst_sheet_next = st.selectbox("Next month week tab", avail_next,
                                           index=avail_next.index(auto_next) if auto_next else 0,
                                           key="fcst_sheet_next")
            if auto_next:
                st.caption(f"Auto-detected next month tab: **{auto_next}**")

        if fcst_csv and st.button("Preview Changes", key="fcst_preview"):
            csv_bytes = fcst_csv.read()

            # Current month workbook
            wb_fcst_full = openpyxl.load_workbook(io.BytesIO(fcst_bytes), data_only=False)
            df_fcst = parse_csv(csv_bytes)
            ws_fcst = wb_fcst_full[fcst_sheet]
            fcst_changes, fcst_warnings = build_forecast_change_plan(df_fcst, ws_fcst)

            st.session_state["fcst_changes"]   = fcst_changes
            st.session_state["fcst_wb_bytes"]  = fcst_bytes
            st.session_state["fcst_sheet_sel"] = fcst_sheet
            st.session_state["fcst_warnings"]  = fcst_warnings

            # Next month workbook (weeks 3 & 4)
            if fcst_xl_next:
                wb_next_full = openpyxl.load_workbook(io.BytesIO(fcst_next_bytes), data_only=False)
                ws_next      = wb_next_full[fcst_sheet_next]
                next_changes, next_warnings = build_next_month_forecast_plan(df_fcst, ws_next)
                st.session_state["fcst_next_changes"]   = next_changes
                st.session_state["fcst_next_wb_bytes"]  = fcst_next_bytes
                st.session_state["fcst_next_sheet_sel"] = fcst_sheet_next
                st.session_state["fcst_next_warnings"]  = next_warnings
            else:
                st.session_state.pop("fcst_next_changes", None)

        if "fcst_changes" in st.session_state:
            for w in st.session_state.get("fcst_warnings", []):
                st.warning(w)

            fcst_ch   = st.session_state["fcst_changes"]
            will_fcst = [c for c in fcst_ch if not c["skip_reason"]]
            skip_fcst = [c for c in fcst_ch if c["skip_reason"]]

            st.subheader("Current month changes")
            c1, c2 = st.columns(2)
            c1.metric("Cells to update", len(will_fcst))
            c2.metric("Skipped",         len(skip_fcst))

            preview_fcst = []
            for c in fcst_ch:
                preview_fcst.append({
                    "Label":  c["label"],
                    "Row":    c["row"],
                    "Col":    c["col"],
                    "Value":  c["new_value"],
                    "Status": "✅ will write" if not c["skip_reason"] else f"⚠️ skip ({c['skip_reason']})",
                })
            st.dataframe(preview_fcst, use_container_width=True)

            if "fcst_next_changes" in st.session_state:
                next_ch = st.session_state["fcst_next_changes"]
                for w in st.session_state.get("fcst_next_warnings", []):
                    st.warning(w)
                will_next = [c for c in next_ch if not c["skip_reason"]]
                skip_next = [c for c in next_ch if c["skip_reason"]]
                st.subheader("Next month changes")
                n1, n2 = st.columns(2)
                n1.metric("Cells to update", len(will_next))
                n2.metric("Skipped",         len(skip_next))
                preview_next = []
                for c in next_ch:
                    preview_next.append({
                        "Label":  c["label"],
                        "Row":    c["row"],
                        "Col":    c["col"],
                        "Value":  c["new_value"],
                        "Status": "✅ will write" if not c["skip_reason"] else f"⚠️ skip ({c['skip_reason']})",
                    })
                st.dataframe(preview_next, use_container_width=True)

            if st.button("Confirm and Apply Changes", key="fcst_apply"):
                # Apply current month
                wb_out = openpyxl.load_workbook(io.BytesIO(st.session_state["fcst_wb_bytes"]), data_only=False)
                apply_forecast_changes(wb_out, st.session_state["fcst_sheet_sel"], fcst_ch)
                color_tab_done(wb_out, st.session_state["fcst_sheet_sel"])
                strip_tables(wb_out)
                out3 = io.BytesIO()
                wb_out.save(out3)
                st.download_button(
                    "Download Updated Current Month Forecast",
                    data=out3.getvalue(),
                    file_name="Forecast_current_updated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                # Apply next month (if uploaded)
                if "fcst_next_changes" in st.session_state:
                    wb_next_out = openpyxl.load_workbook(io.BytesIO(st.session_state["fcst_next_wb_bytes"]), data_only=False)
                    apply_forecast_changes(wb_next_out, st.session_state["fcst_next_sheet_sel"],
                                           st.session_state["fcst_next_changes"])
                    color_tab_done(wb_next_out, st.session_state["fcst_next_sheet_sel"])
                    strip_tables(wb_next_out)
                    out4 = io.BytesIO()
                    wb_next_out.save(out4)
                    st.download_button(
                        "Download Updated Next Month Forecast",
                        data=out4.getvalue(),
                        file_name="Forecast_next_month_updated.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
