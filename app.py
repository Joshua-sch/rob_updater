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

tab_rob, tab_strategy = st.tabs(["ROB Update", "Strategy Report"])

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
