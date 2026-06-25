import streamlit as st
import pandas as pd
import openpyxl
import io
import re
import datetime
import bcrypt
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

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

# CSV column index for each field (0-based) — source data never changes
STRATEGY_CSV_COLS = {
    "otb_trans":    (15, "OTB TY Trans (Indiv Count)"),
    "grp_pu_ty":    ( 7, "GRP PU TY"),
    "grp_npu_ty":   ( 8, "GRP N/PU TY"),
    "ooo_rms":      ( 4, "OOO RMS"),
    "trans_rev_ty": (16, "Trans Rev TY"),
    "grp_rev_ty":   ( 9, "Grp Rev TY"),
}

# Each field: list of (row3_keyword, row4_keyword) pairs to try in order.
# Match = both keywords found (case-insensitive) in their respective rows of that column.
# A None keyword means "don't check that row."
# "!LY" suffix on a keyword means the column must NOT contain "LY" in either header row.
STRATEGY_FIELD_PATTERNS = {
    "otb_trans":    [("OTB TY", "TRANS"),       ("TRANS!LY", "SOLD!LY")],
    "grp_pu_ty":    [("GRP PU", "TY!LY"),       ("GROUP!LY", "SOLD!LY")],
    "grp_npu_ty":   [("GRP N/PU", "TY!LY"),     ("GRP RMS", "N/PU"),    ("N/PU", None)],
    "ooo_rms":      [("OOO", None)],
    "trans_rev_ty": [("TY TRANS", "REV"),        ("TRAN", "REV TY")],
    "grp_rev_ty":   [("GRP TY", "REV"),          ("GRP", "REV TY")],
    "otb_lst_wk":   [("OTB LST", None),          ("LST WK", None),       ("LAST WK", None), ("LST WEK", None)],
}


def _kw_matches(cell_val, keyword, r3_val, r4_val):
    """Check if keyword matches cell_val. If keyword ends with !LY,
    also verify neither r3_val nor r4_val contains 'LY'."""
    must_not_ly = keyword.endswith("!LY")
    kw = keyword.replace("!LY", "").strip()
    if kw.upper() not in str(cell_val or "").upper():
        return False
    if must_not_ly:
        if "LY" in str(r3_val or "").upper() or "LY" in str(r4_val or "").upper():
            return False
    return True


def detect_strategy_columns(ws):
    """Scan rows 3+4 of ws and return {field_key: col_index} for each field.
    Raises ValueError listing any fields that could not be found.
    """
    max_col = ws.max_column
    # Build lookup: col → (r3_text, r4_text)
    headers = {}
    for c in range(1, max_col + 1):
        headers[c] = (
            str(ws.cell(3, c).value or "").strip(),
            str(ws.cell(4, c).value or "").strip(),
        )

    col_map = {}
    for field, patterns in STRATEGY_FIELD_PATTERNS.items():
        found = None
        for r3_kw, r4_kw in patterns:
            for c, (r3v, r4v) in headers.items():
                r3_ok = r3_kw is None or _kw_matches(r3v, r3_kw, r3v, r4v)
                r4_ok = r4_kw is None or _kw_matches(r4v, r4_kw, r3v, r4v)
                if r3_ok and r4_ok and (r3_kw or r4_kw):
                    found = c
                    break
            if found:
                break
        if found:
            col_map[field] = found
        else:
            col_map[field] = None  # will surface as a warning, not a crash

    return col_map


def detect_date_column(ws):
    """Find the column whose data rows (5+) contain the earliest daily dates —
    i.e. the column that maps to each row's actual calendar date.
    Scans cols 1-10 only (dates are always on the left side).
    """
    import collections
    col_dates = collections.defaultdict(list)
    for r in range(5, min(ws.max_row + 1, 15)):
        for c in range(1, 11):
            v = ws.cell(r, c).value
            if isinstance(v, datetime.datetime):
                col_dates[c].append(v.date())
            elif isinstance(v, datetime.date):
                col_dates[c].append(v)

    # Pick the col with the most dates that form a consecutive daily sequence
    best_col, best_score = 3, 0  # fallback to col 3
    for c, dates in col_dates.items():
        if len(dates) < 3:
            continue
        dates_sorted = sorted(dates)
        consecutive = sum(
            1 for i in range(1, len(dates_sorted))
            if (dates_sorted[i] - dates_sorted[i-1]).days == 1
        )
        # Prefer the col with earliest starting date (first of month)
        score = consecutive * 10 - dates_sorted[0].day
        if score > best_score:
            best_score = score
            best_col = c
    return best_col


def build_date_row_map(wb):
    """Build {date: row_number} from WKONE using auto-detected date column."""
    ws = wb["WKONE"]
    date_col = detect_date_column(ws)
    mapping = {}
    for row_num in range(5, ws.max_row + 1):
        val = ws.cell(row_num, date_col).value
        if isinstance(val, datetime.datetime):
            mapping[val.date()] = row_num
        elif isinstance(val, datetime.date):
            mapping[val] = row_num
    return mapping


def find_otb_date_cell(ws):
    """Return (row, col) of the as-of date cell — one row above the OTB/TRANS header.
    Tries both Plymouth-style (OTB TY / TRANS) and Long Beach-style (TRANS / SOLD).
    """
    col_map = detect_strategy_columns(ws)
    otb_col = col_map.get("otb_trans")
    if otb_col:
        # Find which of rows 3 or 4 has the header, then go one above
        for r in range(2, 6):
            v = str(ws.cell(r, otb_col).value or "").strip().upper()
            if "OTB" in v or "TRANS" in v or "SOLD" in v:
                return r - 1, otb_col
    return 2, 4  # fallback


def build_strategy_change_plan(df, wb, sheet_name, drive_service=None, hotel_id=None, hotel_name=None):
    today = datetime.date.today()
    scope_start = today.replace(day=1)
    scope_end = datetime.date(today.year, 12, 31)

    date_row_map = build_date_row_map(wb)
    ws = wb[sheet_name]

    # Detect actual column positions from headers — no guessing
    col_map = detect_strategy_columns(ws)
    missing = [f for f, c in col_map.items() if c is None
               and f != "otb_lst_wk"]  # otb_lst_wk only expected on WKONE
    if missing:
        st.warning(f"Strategy: could not locate columns for: {', '.join(missing)}")

    # For WKONE, pull OTB Lst Wek from previous month's SR
    prev_otb_map = {}
    if sheet_name == "WKONE" and col_map.get("otb_lst_wk") and drive_service and hotel_id:
        prev_otb_map = get_prev_month_otb_trans(drive_service, hotel_id, hotel_name or "", scope_start)

    changes = []

    # Today's date above the OTB TY TRANS header
    date_row, date_col = find_otb_date_cell(ws)
    if date_row >= 1:
        changes.append({
            "date": today, "row": date_row, "col": date_col,
            "label": "As-of date", "new_value": today,
            "skip_reason": "formula" if is_formula(ws.cell(date_row, date_col).value) else None,
        })

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
        for field, (csv_col, label) in STRATEGY_CSV_COLS.items():
            excel_col = col_map.get(field)
            if excel_col is None:
                continue  # column not found in this sheet — already warned above
            val = safe_float(row[csv_col])
            skip = "formula" if is_formula(ws.cell(excel_row, excel_col).value) else None
            changes.append({
                "date": d, "row": excel_row, "col": excel_col,
                "label": label, "new_value": val, "skip_reason": skip,
            })

        # OTB Lst Wek — WKONE only, sourced from previous month's SR
        lst_wk_col = col_map.get("otb_lst_wk")
        if lst_wk_col and d in prev_otb_map:
            skip = "formula" if is_formula(ws.cell(excel_row, lst_wk_col).value) else None
            changes.append({
                "date": d, "row": excel_row, "col": lst_wk_col,
                "label": "OTB Lst Wek", "new_value": prev_otb_map[d], "skip_reason": skip,
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


# ── Google Drive ──────────────────────────────────────────────────────────────

@st.cache_data(ttl=300)
def get_hotels_from_drive():
    """Return list of (display_name, folder_id) for every top-level folder
    that contains a 'REVENUE REPORTS' subfolder — i.e. each hotel folder.
    Cached for 5 minutes so it doesn't hit Drive on every rerender.
    """
    try:
        svc = get_drive_service()
        # All folders visible to the service account
        q = "mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        result = svc.files().list(q=q, fields="files(id, name)", pageSize=100).execute()
        folders = result.get("files", [])

        hotels = []
        for folder in folders:
            name = folder["name"]
            # Skip revenue report folders, year folders, month folders
            if "REVENUE REPORTS" in name.upper():
                continue
            if re.search(r'\b20\d{2}\b', name):
                continue
            # Must have a child folder containing "REVENUE REPORTS"
            child_q = ("'%s' in parents and trashed = false and "
                       "mimeType = 'application/vnd.google-apps.folder'") % folder["id"]
            children = svc.files().list(q=child_q, fields="files(name)", pageSize=20).execute()
            has_rev = any("REVENUE REPORTS" in c["name"].upper() for c in children.get("files", []))
            if has_rev:
                hotels.append((name, folder["id"]))

        return sorted(hotels, key=lambda x: x[0])
    except Exception:
        return []

WORKBOOK_TYPES = ["ROB", "Strategy Report", "Forecast"]

# Maps workbook type → partial filename keyword to search for in Drive
WORKBOOK_KEYWORDS = {
    "ROB":             "ROB",
    "Strategy Report": "STRATEGY",
    "Forecast":        "FORECAST",
}

CREDS_PATH = "credentials.json.json"
SCOPES     = ["https://www.googleapis.com/auth/drive"]


@st.cache_resource
def get_drive_service():
    creds = service_account.Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def drive_find_folder_by_keyword(service, keyword, parent_id=None):
    """Return the first folder whose name contains keyword (case-insensitive)."""
    q = "mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    if parent_id:
        q += " and '%s' in parents" % parent_id
    result = service.files().list(q=q, fields="files(id, name)", pageSize=100).execute()
    for f in result.get("files", []):
        if keyword.lower() in f["name"].lower():
            return f["id"], f["name"]
    return None, None


def drive_find_file(service, keyword, parent_id):
    """Return (file_id, file_name) for first xlsx whose name contains keyword,
    excluding files whose name also contains 'copy' (to skip backup copies)."""
    q = ("'%s' in parents and trashed = false "
         "and mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'") % parent_id
    result = service.files().list(q=q, fields="files(id, name)").execute()
    for f in result.get("files", []):
        name_lower = f["name"].lower()
        if keyword.lower() in name_lower and "copy" not in name_lower:
            return f["id"], f["name"]
    return None, None


def drive_download(service, file_id) -> bytes:
    buf = io.BytesIO()
    req = service.files().get_media(fileId=file_id)
    dl  = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    return buf.getvalue()


def drive_upload(service, file_id, file_bytes: bytes, file_name: str):
    """Overwrite an existing Drive file with new bytes."""
    buf   = io.BytesIO(file_bytes)
    media = MediaIoBaseUpload(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    service.files().update(fileId=file_id, media_body=media).execute()


def drive_find_or_create_month_folder(service, rev_id: str, year_id: str, month_date: datetime.date, hotel_name: str):
    """Return (folder_id, folder_name) for the month folder under year_id.
    Folder name pattern: '[LETTER]: [MON][YEAR] REVENUE REPORTS [HOTEL_UPPER]'
    where LETTER = A-L for Jan-Dec.
    Creates the folder if it doesn't exist.
    """
    month_kw = month_date.strftime("%b%Y").upper()  # e.g. JUL2026

    # Try to find existing folder first
    q = ("mimeType='application/vnd.google-apps.folder' and trashed=false "
         "and '%s' in parents") % year_id
    result = service.files().list(q=q, fields="files(id,name)", pageSize=100).execute()
    for f in result.get("files", []):
        if month_kw in f["name"].upper():
            return f["id"], f["name"]

    # Not found — infer name from existing sibling folder naming pattern
    # Pattern: "[LETTER]: [MON][YEAR] REVENUE REPORTS [HOTEL]"
    letter = chr(ord("A") + month_date.month - 1)  # A=Jan, B=Feb, ..., L=Dec

    # Try to infer hotel suffix from existing month folders in this year folder
    hotel_suffix = hotel_name.upper()
    for f in result.get("files", []):
        name_upper = f["name"].upper()
        if "REVENUE REPORTS" in name_upper:
            # Extract everything after "REVENUE REPORTS "
            idx = name_upper.find("REVENUE REPORTS ")
            if idx != -1:
                hotel_suffix = f["name"][idx + len("REVENUE REPORTS "):].strip()
                break

    new_name = f"{letter}: {month_kw} REVENUE REPORTS {hotel_suffix}"
    folder_meta = {
        "name": new_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [year_id],
    }
    created = service.files().create(body=folder_meta, fields="id,name").execute()
    return created["id"], created["name"]


def drive_copy_file(service, source_file_id: str, new_name: str, parent_folder_id: str):
    """Copy a Drive file to a new name in the given folder. Returns (new_file_id, new_name)."""
    body = {"name": new_name, "parents": [parent_folder_id]}
    copied = service.files().copy(fileId=source_file_id, body=body, fields="id,name").execute()
    return copied["id"], copied["name"]


def find_sr_master(service, hotel_id: str):
    """Search the hotel's REVENUE REPORTS tree for the SR master file.
    Returns (file_id, file_name) or (None, error_str).
    """
    rev_id, rev_name = drive_find_folder_by_keyword(service, "REVENUE REPORTS", parent_id=hotel_id)
    if not rev_id:
        return None, "No REVENUE REPORTS folder found."

    # Search all subfolders for a file with MASTER and STRATEGY in the name
    q = ("trashed=false "
         "and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
         "and name contains 'MASTER' and name contains 'STRATEGY'")
    result = service.files().list(q=q, fields="files(id,name,parents)", pageSize=50).execute()
    for f in result.get("files", []):
        return f["id"], f["name"]
    return None, "No STRATEGY master file found in Drive."


def setup_new_sr_month(service, hotel_id: str, hotel_name: str, target_month: datetime.date):
    """
    New month SR setup:
      1. Find SR master, copy it, rename to [MON][YEAR] STRATEGY [HOTEL].xlsx
      2. Find or create month folder under the year folder
      3. Move the copy into that folder
    Returns (new_file_name, error_str).
    """
    year_kw = str(target_month.year)
    month_kw = target_month.strftime("%b%Y").upper()

    # Resolve REVENUE REPORTS and year folder
    rev_id, _ = drive_find_folder_by_keyword(service, "REVENUE REPORTS", parent_id=hotel_id)
    if not rev_id:
        return None, "No REVENUE REPORTS folder."
    year_id, _ = drive_find_folder_by_keyword(service, year_kw, parent_id=rev_id)
    if not year_id:
        return None, f"No {year_kw} folder."

    # Find or create month folder
    month_id, month_name = drive_find_or_create_month_folder(service, rev_id, year_id, target_month, hotel_name)

    # Check if SR already exists in that folder
    existing_id, existing_name = drive_find_file(service, "STRATEGY", month_id)
    if existing_id and "master" not in existing_name.lower():
        return existing_name, None  # already set up

    # Find master
    master_id, master_name = find_sr_master(service, hotel_id)
    if not master_id:
        return None, master_name  # error string

    # Infer hotel suffix from master file name for the new file name
    # e.g. "MASTER 2026 STRATEGY PLYMOUTH.xlsx" → "PLYMOUTH"
    hotel_suffix = hotel_name.upper()
    name_upper = master_name.upper().replace(".XLSX", "")
    if "STRATEGY" in name_upper:
        after = name_upper[name_upper.find("STRATEGY") + len("STRATEGY"):].strip()
        if after:
            hotel_suffix = master_name[master_name.upper().find("STRATEGY") + len("STRATEGY"):].strip().replace(".xlsx", "").replace(".XLSX", "").strip()

    new_file_name = f"{month_kw} STRATEGY {hotel_suffix}.xlsx"
    _, created_name = drive_copy_file(service, master_id, new_file_name, month_id)
    return created_name, None


def get_prev_month_otb_trans(service, hotel_id: str, hotel_name: str, current_month: datetime.date):
    """Pull OTB TY Trans values for current_month dates from previous month's SR last filled tab.
    Returns {date: value} or {} on any failure.
    """
    prev_month = (current_month.replace(day=1) - datetime.timedelta(days=1)).replace(day=1)
    result, err = resolve_drive_workbook(service, hotel_id, hotel_name, "Strategy Report", month_date=prev_month)
    if err or not result:
        return {}

    file_id, _ = result
    try:
        file_bytes = drive_download(service, file_id)
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=False)
    except Exception:
        return {}

    # Find last filled (colored) tab
    last_filled_sheet = None
    for sheet_name in STRATEGY_SHEETS:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        tab_color = ws.sheet_properties.tabColor
        if tab_color is not None:
            last_filled_sheet = sheet_name

    if not last_filled_sheet:
        return {}

    ws = wb[last_filled_sheet]
    col_map = detect_strategy_columns(ws)
    otb_col = col_map.get("otb_trans")
    date_col = detect_date_column(ws)
    if not otb_col:
        return {}

    result_map = {}
    for r in range(5, ws.max_row + 1):
        date_val = ws.cell(r, date_col).value
        if isinstance(date_val, datetime.datetime):
            d = date_val.date()
        elif isinstance(date_val, datetime.date):
            d = date_val
        else:
            continue
        if d < current_month:
            continue
        cell_val = ws.cell(r, otb_col).value
        if cell_val is not None and not is_formula(cell_val):
            result_map[d] = safe_float(cell_val)

    return result_map


def resolve_drive_workbook(service, hotel_id: str, hotel_name: str, workbook_type: str, month_date: datetime.date = None):
    """
    Walk Drive: Hotel > REVENUE REPORTS > Year > Month > file.
    Returns ((file_id, file_name), None) or (None, error_message).
    Never touches files whose name contains 'master'.
    """
    if month_date is None:
        month_date = datetime.date.today()

    month_kw   = month_date.strftime("%b%Y").upper()
    year_kw    = str(month_date.year)
    wb_keyword = WORKBOOK_KEYWORDS[workbook_type]

    # Revenue Reports folder
    rev_id, rev_name = drive_find_folder_by_keyword(service, "REVENUE REPORTS", parent_id=hotel_id)
    if not rev_id:
        return None, f"No 'REVENUE REPORTS' folder found under '{hotel_name}'."

    # Year folder
    year_id, year_name = drive_find_folder_by_keyword(service, year_kw, parent_id=rev_id)
    if not year_id:
        return None, f"No '{year_kw}' folder found under '{rev_name}'."

    # Month folder
    month_id, month_name = drive_find_folder_by_keyword(service, month_kw, parent_id=year_id)
    if not month_id:
        return None, f"No '{month_kw}' folder found under '{year_name}'."

    # File — never match master docs
    file_id, file_name = drive_find_file(service, wb_keyword, month_id)
    if not file_id:
        return None, f"No '{wb_keyword}' workbook found in '{month_name}'."
    if "master" in file_name.lower():
        return None, f"Resolved file '{file_name}' looks like a master doc — aborting."

    return (file_id, file_name), None


# ── UI ────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Linchris Weekly Tools", layout="wide")

# ── Login gate ────────────────────────────────────────────────────────────────
def check_login(username: str, password: str) -> bool:
    correct_user = st.secrets["auth"]["username"]
    correct_hash = st.secrets["auth"]["password_hash"].encode()
    return username == correct_user and bcrypt.checkpw(password.encode(), correct_hash)

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("Linchris Hotel Corporation")
    st.subheader("Please log in to continue")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log In")
        if submitted:
            if check_login(username, password):
                st.session_state["authenticated"] = True
                st.rerun()
            else:
                st.error("Incorrect username or password.")
    st.stop()

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

title_col, toggle_col = st.columns([6, 1])
with title_col:
    st.title("Linchris Hotel Corporation — Weekly Update Tools")
with toggle_col:
    st.write("")
    test_mode = st.toggle("Test Mode", value=False, key="test_mode")

# ── Manual upload (collapsed) ─────────────────────────────────────────────────
with st.expander("Manual Upload", expanded=False):
    with st.expander("ROB Update"):
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
    
    with st.expander("Strategy Report"):
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
    
    with st.expander("Forecast"):
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
    
st.divider()
# ── Google Drive Update ───────────────────────────────────────────────────────
st.header("Google Drive Update")
st.caption(f"Current month: **{datetime.date.today().strftime('%B %Y')}**")

hotels = get_hotels_from_drive()
hotel_names = [h[0] for h in hotels]
hotel_id_map = {h[0]: h[1] for h in hotels}

col_h, col_w = st.columns(2)
with col_h:
    hotel_sel = st.selectbox("Hotel", hotel_names if hotel_names else ["(no hotels found)"], key="drive_hotel")
with col_w:
    wb_sels = st.multiselect(
        "Workbooks to update",
        WORKBOOK_TYPES,
        default=WORKBOOK_TYPES,
        key="drive_wb",
    )

drive_csv = st.file_uploader("Upload CSV (Business on the Books)", type=["csv"], key="drive_csv")
drive_rate_csv = None
if "Strategy Report" in (wb_sels or []):
    drive_rate_csv = st.file_uploader("Upload Rates & Restrictions CSV", type=["csv"], key="drive_rate_csv")
forecast_next_month = False
if "Forecast" in (wb_sels or []):
    forecast_next_month = st.checkbox("Forecast next month as well?", key="drive_fcst_next")

start_new_month = st.checkbox("Start new month", key="drive_new_month")
if start_new_month:
    with st.container(border=True):
        st.markdown("**New Month Setup**")
        next_month_dt = (datetime.date.today().replace(day=1) + datetime.timedelta(days=32)).replace(day=1)
        st.caption(f"Target month: **{next_month_dt.strftime('%B %Y')}**")
        new_month_hotel = hotel_sel
        if st.button("Set Up Strategy Report for New Month", key="btn_new_month_sr"):
            try:
                svc = get_drive_service()
                hotel_id_nm = hotel_id_map.get(new_month_hotel, "")
                with st.spinner(f"Setting up {next_month_dt.strftime('%b %Y')} SR for {new_month_hotel}..."):
                    created_name, err = setup_new_sr_month(svc, hotel_id_nm, new_month_hotel, next_month_dt)
                if err:
                    st.error(f"New month setup failed: {err}")
                else:
                    st.success(f"Created **{created_name}** in Drive — ready for {next_month_dt.strftime('%B %Y')}.")
            except Exception as e:
                st.error(f"New month setup error: {e}")


def build_all_plans(svc, hotel_sel, hotel_id, wb_sels, df, rate_df, forecast_next_month=False):
    all_plans = {}
    for wb_type in wb_sels:
        result, err = resolve_drive_workbook(svc, hotel_id, hotel_sel, wb_type)
        if err:
            st.error(f"{wb_type}: {err}")
            continue
        file_id, file_name = result
        wb_bytes = drive_download(svc, file_id)
        wb       = openpyxl.load_workbook(io.BytesIO(wb_bytes), data_only=False)
        if wb_type == "ROB":
            avail    = [s for s in ROB_SHEETS if s in wb.sheetnames]
            auto     = first_uncolored_sheet(wb, avail)
            sheet    = auto or avail[0]
            changes  = build_rob_change_plan(df, wb[sheet])
            warnings = []
        elif wb_type == "Strategy Report":
            avail    = [s for s in STRATEGY_SHEETS if s in wb.sheetnames]
            auto     = first_uncolored_sheet(wb, avail)
            sheet    = auto or avail[0]
            changes  = build_strategy_change_plan(df, wb, sheet, drive_service=svc, hotel_id=hotel_id, hotel_name=hotel_sel)
            warnings = []
            if rate_df is not None:
                rate_changes, rate_warnings = build_rates_change_plan(rate_df, wb, sheet)
                changes  += rate_changes
                warnings += rate_warnings
        else:  # Forecast — current month
            avail    = [s for s in FORECAST_SHEETS if s in wb.sheetnames]
            auto     = first_uncolored_sheet(wb, avail)
            sheet    = auto or avail[0]
            changes, warnings = build_forecast_change_plan(df, wb[sheet])
        all_plans[wb_type] = {
            "file_id":   file_id,
            "file_name": file_name,
            "wb_bytes":  wb_bytes,
            "sheet":     sheet,
            "changes":   changes,
            "warnings":  warnings,
        }

    # Next-month Forecast: only when checkbox is ticked
    if "Forecast" in wb_sels and forecast_next_month:
        next_month_dt = (datetime.date.today().replace(day=1) + datetime.timedelta(days=32)).replace(day=1)
        nm_result, nm_err = resolve_drive_workbook(svc, hotel_id, hotel_sel, "Forecast", month_date=next_month_dt)
        if nm_err:
            st.warning(f"Next month Forecast: {nm_err}")
        else:
            nm_file_id, nm_file_name = nm_result
            nm_bytes = drive_download(svc, nm_file_id)
            nm_wb    = openpyxl.load_workbook(io.BytesIO(nm_bytes), data_only=False)
            nm_avail = [s for s in FORECAST_SHEETS if s in nm_wb.sheetnames]
            nm_auto  = first_uncolored_sheet(nm_wb, nm_avail)
            nm_sheet = nm_auto or nm_avail[0]
            nm_changes, nm_warnings = build_next_month_forecast_plan(df, nm_wb[nm_sheet])
            all_plans["Forecast (next month)"] = {
                "file_id":   nm_file_id,
                "file_name": nm_file_name,
                "wb_bytes":  nm_bytes,
                "sheet":     nm_sheet,
                "changes":   nm_changes,
                "warnings":  nm_warnings,
            }

    return all_plans


def apply_and_upload(svc, all_plans):
    saved, errors = [], []
    for wb_type, plan in all_plans.items():
        try:
            wb_apply = openpyxl.load_workbook(io.BytesIO(plan["wb_bytes"]), data_only=False)
            if wb_type == "ROB":
                apply_rob_changes(wb_apply, plan["sheet"], plan["changes"])
            elif wb_type == "Strategy Report":
                apply_strategy_changes(wb_apply, plan["sheet"], plan["changes"])
            else:
                apply_forecast_changes(wb_apply, plan["sheet"], plan["changes"])
            color_tab_done(wb_apply, plan["sheet"])
            strip_tables(wb_apply)
            out = io.BytesIO()
            wb_apply.save(out)
            drive_upload(svc, plan["file_id"], out.getvalue(), plan["file_name"])
            saved.append(plan["file_name"])
        except Exception as e:
            errors.append(f"{wb_type}: {e}")
    return saved, errors


ready = drive_csv and wb_sels

if test_mode:
    # ── Test mode: preview first, then confirm ────────────────────────────────
    if ready and st.button("Preview Changes", key="drive_preview"):
        try:
            svc     = get_drive_service()
            df      = parse_csv(drive_csv.read())
            rate_df = parse_rate_csv(drive_rate_csv.read()) if drive_rate_csv else None
            st.session_state["drive_plans"]     = build_all_plans(svc, hotel_sel, hotel_id_map.get(hotel_sel, ""), wb_sels, df, rate_df, forecast_next_month)
            st.session_state["drive_hotel_sel"] = hotel_sel
        except Exception as e:
            st.error(f"Drive error: {e}")

    if "drive_plans" in st.session_state:
        all_plans = st.session_state["drive_plans"]
        for wb_type, plan in all_plans.items():
            st.subheader(wb_type)
            st.caption(f"File: **{plan['file_name']}** — Tab: **{plan['sheet']}**")
            for w in plan["warnings"]:
                st.warning(w)
            ch = plan["changes"]
            will_write = [c for c in ch if not c["skip_reason"]]
            skipped    = [c for c in ch if c["skip_reason"]]
            c1, c2 = st.columns(2)
            c1.metric("Cells to update", len(will_write))
            c2.metric("Skipped",         len(skipped))
            st.dataframe([{
                "Label":  c["label"],
                "Row":    c["row"],
                "Col":    c["col"],
                "Value":  c["new_value"],
                "Status": "✅ will write" if not c["skip_reason"] else f"⚠️ skip ({c['skip_reason']})",
            } for c in ch], use_container_width=True)

        if st.button("Confirm and Save All to Google Drive", key="drive_apply"):
            try:
                saved, errors = apply_and_upload(get_drive_service(), all_plans)
                for name in saved:
                    st.success(f"Saved **{name}** to Google Drive.")
                for err in errors:
                    st.error(err)
            except Exception as e:
                st.error(f"Drive error: {e}")
else:
    # ── Normal mode: one click ────────────────────────────────────────────────
    if ready and st.button("Upload Data to Workbooks", key="drive_go", type="primary"):
        try:
            svc     = get_drive_service()
            df      = parse_csv(drive_csv.read())
            rate_df = parse_rate_csv(drive_rate_csv.read()) if drive_rate_csv else None
            with st.spinner("Updating workbooks in Google Drive..."):
                all_plans       = build_all_plans(svc, hotel_sel, hotel_id_map.get(hotel_sel, ""), wb_sels, df, rate_df, forecast_next_month)
                saved, errors   = apply_and_upload(svc, all_plans)
            for name in saved:
                st.success(f"Saved **{name}** to Google Drive.")
            for err in errors:
                st.error(err)
        except Exception as e:
            st.error(f"Drive error: {e}")

