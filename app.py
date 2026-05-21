import os
import io
import uuid
import traceback
from copy import copy
from datetime import datetime, date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

APP_TITLE = "Huliot Travel Expense Entry"
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_FILE = os.path.join(BASE_DIR, "Travelling Expenses Sheet.xlsx")
TEMPLATE_SHEET_NAME = "Travel Reimbur. Form"
WORKSHEET_NAME_DEFAULT = "Entries"

# Google Sheet database columns. This matches your Google Sheet backend format.
COLUMNS = [
    "id", "empName", "designation", "location", "reportMonthText",
    "date", "from", "to", "vehicle", "company", "contact", "invoice",
    "toll", "fuel", "lodging", "food", "tel", "courier", "rikshaw",
    "remarks", "total", "created_at"
]

AMOUNT_COLUMNS = ["toll", "fuel", "lodging", "food", "tel", "courier", "rikshaw"]

CSS = """
<style>
    .stApp { background: #f4f7fb; color: #1f2937; }
    .main .block-container {
        max-width: 1120px;
        padding-top: 0 !important;
        padding-bottom: 34px;
    }
    .html-header {
        background: #1e40af;
        color: white;
        padding: 15px 16px 16px 16px;
        text-align: center;
        margin: 0 -2rem 18px -2rem;
    }
    .brand {
        display: inline-block;
        margin-bottom: 7px;
        padding: 4px 9px;
        border: 1px solid rgba(255,255,255,.55);
        border-radius: 999px;
        font-size: 10px;
        letter-spacing: .3px;
        background: rgba(255,255,255,.12);
    }
    .html-header h1 { margin: 0; font-size: 22px; line-height: 1.1; font-weight: 800; }
    .html-header p { margin: 6px 0 0; font-size: 12px; opacity: .95; }

    section[data-testid="stSidebar"] { display: none; }
    div[data-testid="stForm"], .card-box {
        background: white;
        border-radius: 12px;
        padding: 16px 16px 14px 16px;
        margin-bottom: 14px;
        box-shadow: 0 4px 14px rgba(15, 23, 42, .08);
        border: 0;
    }
    .card-title { margin: 0 0 12px 0; font-size: 21px; font-weight: 800; color: #111827; }
    .muted { color: #64748b; font-size: 12px; margin-top: 8px; }
    .ok-text { color: #166534; font-weight: 800; font-size: 13px; }
    .warn-text { color: #b91c1c; font-weight: 800; font-size: 13px; }

    label, [data-testid="stWidgetLabel"] p {
        font-size: 11px !important;
        font-weight: 800 !important;
        color: #374151 !important;
        margin-bottom: 4px !important;
    }
    input, textarea, select {
        border: 1px solid #cbd5e1 !important;
        border-radius: 6px !important;
        font-size: 13px !important;
        background: white !important;
        min-height: 34px !important;
    }
    textarea { min-height: 58px !important; }
    div[data-baseweb="input"] > div,
    div[data-baseweb="select"] > div,
    div[data-baseweb="textarea"] > div {
        border-radius: 6px !important;
        border-color: #cbd5e1 !important;
        min-height: 34px !important;
    }
    div[data-testid="stVerticalBlock"] { gap: .55rem; }
    div[data-testid="column"] { padding-left: .2rem; padding-right: .2rem; }
    .stButton>button, .stDownloadButton>button, div[data-testid="stFormSubmitButton"] button {
        border: 0;
        border-radius: 6px;
        padding: 7px 13px;
        font-size: 13px;
        font-weight: 800;
        min-height: 34px;
    }
    div[data-testid="stFormSubmitButton"] button { background: #2563eb; color: white; }
    .stDownloadButton>button { background: #16a34a; color: white; }
    button[kind="secondary"] { background: #e5e7eb; color: #111827; }

    .summary-box {
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        border-radius: 8px;
        padding: 10px;
        text-align: center;
        min-height: 58px;
    }
    .summary-box span { display: block; font-size: 11px; color: #475569; }
    .summary-box strong { display: block; margin-top: 4px; font-size: 15px; color: #1e3a8a; }
    [data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }
    .small-note { font-size: 11px; color: #64748b; margin-top: 4px; }
    .connection-detail { font-size: 11px; color: #64748b; }
</style>
"""


def money(value):
    try:
        if value in (None, ""):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def make_str(value):
    if value is None:
        return ""
    return str(value).strip()


def current_stamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_entry_date(value):
    text = make_str(value)
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except Exception:
            pass
    return date.today()


def format_date_for_sheet(value: date):
    return value.strftime("%Y-%m-%d")


def get_secret_text(section, key, default=""):
    try:
        if section:
            return st.secrets.get(section, {}).get(key, default)
        return st.secrets.get(key, default)
    except Exception:
        return default


@st.cache_resource(show_spinner=False)
def get_google_worksheet_cached():
    import gspread
    from google.oauth2.service_account import Credentials

    if "gcp_service_account" not in st.secrets:
        raise RuntimeError("Missing [gcp_service_account] in Streamlit Secrets.")
    if "gsheets" not in st.secrets:
        raise RuntimeError("Missing [gsheets] section in Streamlit Secrets.")

    spreadsheet_id = make_str(st.secrets["gsheets"].get("spreadsheet_id", ""))
    if not spreadsheet_id:
        raise RuntimeError("Missing [gsheets] spreadsheet_id in Streamlit Secrets.")

    worksheet_name = make_str(st.secrets.get("WORKSHEET_NAME", WORKSHEET_NAME_DEFAULT)) or WORKSHEET_NAME_DEFAULT

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_info(dict(st.secrets["gcp_service_account"]), scopes=scopes)
    client = gspread.authorize(creds)
    spreadsheet = client.open_by_key(spreadsheet_id)

    try:
        ws = spreadsheet.worksheet(worksheet_name)
    except Exception:
        ws = spreadsheet.add_worksheet(title=worksheet_name, rows=1000, cols=len(COLUMNS) + 5)
        ws.update("A1", [COLUMNS])
        return ws

    values = ws.get_all_values()
    if not values:
        ws.update("A1", [COLUMNS])
    elif values[0] != COLUMNS:
        # This backend sheet is only for data storage, so headers are standardized.
        ws.clear()
        ws.update("A1", [COLUMNS])
    return ws


def get_google_worksheet():
    try:
        return get_google_worksheet_cached(), ""
    except Exception:
        return None, traceback.format_exc()


def read_entries() -> pd.DataFrame:
    ws, err = get_google_worksheet()
    if ws is None:
        return pd.DataFrame(columns=COLUMNS)
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[COLUMNS].fillna("")


def append_entry(row: dict):
    row = {col: row.get(col, "") for col in COLUMNS}
    ws, err = get_google_worksheet()
    if ws is None:
        raise RuntimeError(err or "Google Sheets not connected.")
    ws.append_row([row[col] for col in COLUMNS], value_input_option="USER_ENTERED")


def rewrite_entries(df: pd.DataFrame):
    df = df.copy()
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[COLUMNS].fillna("")
    ws, err = get_google_worksheet()
    if ws is None:
        raise RuntimeError(err or "Google Sheets not connected.")
    ws.clear()
    ws.update("A1", [COLUMNS] + df.astype(str).values.tolist())


def update_entry_by_id(entry_id: str, updated_row: dict):
    df = read_entries()
    if df.empty:
        return False
    mask = df["id"].astype(str) == str(entry_id)
    if not mask.any():
        return False
    updated_row = {col: updated_row.get(col, "") for col in COLUMNS}
    for col in COLUMNS:
        df.loc[mask, col] = updated_row[col]
    rewrite_entries(df)
    return True


def delete_entry_by_id(entry_id: str):
    df = read_entries()
    if df.empty:
        return False
    new_df = df[df["id"].astype(str) != str(entry_id)].copy()
    rewrite_entries(new_df)
    return True


def safe_set(ws, cell_ref: str, value):
    cell = ws[cell_ref]
    for mr in ws.merged_cells.ranges:
        if cell.coordinate in mr:
            ws.cell(mr.min_row, mr.min_col).value = value
            return
    cell.value = value


def clear_cell(ws, row: int, col: int):
    coord = ws.cell(row, col).coordinate
    for mr in ws.merged_cells.ranges:
        if coord in mr:
            if row == mr.min_row and col == mr.min_col:
                ws.cell(row, col).value = None
            return
    ws.cell(row, col).value = None


def copy_row_style(ws, source_row: int, target_row: int):
    for col in range(1, ws.max_column + 1):
        src = ws.cell(source_row, col)
        tgt = ws.cell(target_row, col)
        if src.has_style:
            tgt.font = copy(src.font)
            tgt.border = copy(src.border)
            tgt.fill = copy(src.fill)
            tgt.number_format = src.number_format
            tgt.protection = copy(src.protection)
            tgt.alignment = copy(src.alignment)
            tgt._style = copy(src._style)
    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def generate_excel(df: pd.DataFrame, employee: dict) -> bytes:
    if not os.path.exists(TEMPLATE_FILE):
        raise FileNotFoundError("Travelling Expenses Sheet.xlsx is missing in GitHub repository.")

    wb = load_workbook(TEMPLATE_FILE)
    ws = wb[TEMPLATE_SHEET_NAME] if TEMPLATE_SHEET_NAME in wb.sheetnames else wb.active

    safe_set(ws, "A1", f"Travel Reimbursement Form : {employee.get('reportMonthText', '')}")
    safe_set(ws, "A2", f"Name : {employee.get('empName', '')}")
    safe_set(ws, "B2", f"Designation : {employee.get('designation', '')}")
    safe_set(ws, "C2", f"Location : {employee.get('location', '')}")

    filtered = df.copy()
    if employee.get("empName"):
        filtered = filtered[filtered["empName"].astype(str) == employee.get("empName")]
    if employee.get("reportMonthText"):
        filtered = filtered[filtered["reportMonthText"].astype(str) == employee.get("reportMonthText")]

    data_start = 5
    template_end = 36
    fixed_rows = template_end - data_start + 1
    entry_count = len(filtered)

    if entry_count > fixed_rows:
        extra = entry_count - fixed_rows
        ws.insert_rows(template_end + 1, extra)
        for r in range(template_end + 1, template_end + 1 + extra):
            copy_row_style(ws, template_end, r)

    last_data_row = data_start + max(entry_count, fixed_rows) - 1
    for r in range(data_start, last_data_row + 1):
        for c in range(1, 16):
            clear_cell(ws, r, c)

    for r, (_, row) in enumerate(filtered.iterrows(), start=data_start):
        ws.cell(r, 1).value = row.get("date", "")
        ws.cell(r, 2).value = row.get("from", "")
        ws.cell(r, 3).value = row.get("to", "")
        ws.cell(r, 4).value = row.get("company", "")
        ws.cell(r, 5).value = row.get("contact", "")
        ws.cell(r, 6).value = row.get("invoice", "")
        ws.cell(r, 7).value = money(row.get("toll"))
        ws.cell(r, 8).value = money(row.get("fuel"))
        ws.cell(r, 9).value = 0
        ws.cell(r, 10).value = money(row.get("lodging"))
        ws.cell(r, 11).value = money(row.get("food"))
        ws.cell(r, 12).value = money(row.get("tel"))
        ws.cell(r, 13).value = money(row.get("courier"))
        ws.cell(r, 14).value = money(row.get("rikshaw"))
        ws.cell(r, 15).value = f"=SUM(G{r}:N{r})"

    total_row_1 = last_data_row + 1
    total_row_2 = last_data_row + 2
    safe_set(ws, f"N{total_row_1}", "Total ")
    safe_set(ws, f"O{total_row_1}", f"=SUM(O{data_start}:O{last_data_row})")
    safe_set(ws, f"N{total_row_2}", "Total ")
    safe_set(ws, f"O{total_row_2}", f"=O{total_row_1}")

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def rupee(value):
    return f"₹{money(value):.2f}"


def summary_box(label, value):
    st.markdown(f'<div class="summary-box"><span>{label}</span><strong>{value}</strong></div>', unsafe_allow_html=True)


def default_employee():
    return {
        "empName": "Mr. Umesh Nikam",
        "designation": "Technical manager",
        "location": "Pune",
        "reportMonthText": "01May to 30May",
    }


def entry_total(row):
    return sum(money(row.get(col, 0)) for col in AMOUNT_COLUMNS)


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)

    st.markdown(
        '<div class="html-header"><div class="brand">HULIOT INDIA</div>'
        '<h1>Travel Expense Entry App</h1>'
        '<p>Save entries in browser and download in your attached Excel format.</p></div>',
        unsafe_allow_html=True,
    )

    if "employee" not in st.session_state:
        st.session_state.employee = default_employee()

    ws, connection_error = get_google_worksheet()
    with st.container():
        st.markdown('<div class="card-box"><h2 class="card-title">Google Sheet Connection</h2>', unsafe_allow_html=True)
        if ws is not None:
            st.markdown('<p class="ok-text">Storage: Google Sheets connected. Entries are saved live.</p>', unsafe_allow_html=True)
        else:
            st.markdown('<p class="warn-text">Storage: Google Sheets not connected.</p>', unsafe_allow_html=True)
            with st.expander("Show Google Sheet connection error"):
                st.code(connection_error or "No error details available.")
        sheet_id = get_secret_text("gsheets", "spreadsheet_id", "")
        worksheet_name = get_secret_text(None, "WORKSHEET_NAME", WORKSHEET_NAME_DEFAULT)
        st.markdown(f'<p class="connection-detail">Sheet ID: {sheet_id} | Worksheet: {worksheet_name}</p></div>', unsafe_allow_html=True)

    st.markdown('<div class="card-box"><h2 class="card-title">Employee Details</h2>', unsafe_allow_html=True)
    e1, e2, e3, e4 = st.columns(4)
    empName = e1.text_input("Name", value=st.session_state.employee.get("empName", ""))
    designation = e2.text_input("Designation", value=st.session_state.employee.get("designation", ""))
    location = e3.text_input("Location", value=st.session_state.employee.get("location", ""))
    reportMonthText = e4.text_input("Month / Period Text", value=st.session_state.employee.get("reportMonthText", ""))
    st.markdown('<p class="muted">These details are saved automatically in this browser.</p></div>', unsafe_allow_html=True)

    employee = {
        "empName": make_str(empName),
        "designation": make_str(designation),
        "location": make_str(location),
        "reportMonthText": make_str(reportMonthText),
    }
    st.session_state.employee = employee

    with st.form("entry_form", clear_on_submit=True):
        st.markdown('<h2 class="card-title">Add Travel Expense</h2>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        entry_date = c1.date_input("Date", value=None)
        from_loc = c2.text_input("From", placeholder="Start location")
        to_loc = c3.text_input("To", placeholder="End location")
        vehicle = c4.selectbox("2 Wheeler / 4 Wheeler", ["", "2 Wheeler", "4 Wheeler", "Public Transport", "Train", "Bus", "Flight", "Other"])

        c5, c6 = st.columns(2)
        company = c5.text_input("Company Name / Contact / MEP Consultant", placeholder="Site / company / consultant name")
        contact = c6.text_input("Contact Person / Meeting With", placeholder="Person name")

        c7, c8, c9, c10 = st.columns(4)
        invoice = c7.text_input("Invoice No.", placeholder="Optional")
        toll = c8.number_input("Toll / Parking", min_value=0.0, value=0.0, step=1.0)
        fuel = c9.number_input("Petrol / Diesel", min_value=0.0, value=0.0, step=1.0)
        lodging = c10.number_input("Lodging / Boarding", min_value=0.0, value=0.0, step=1.0)

        c11, c12, c13, c14 = st.columns(4)
        food = c11.number_input("Food / Beverages", min_value=0.0, value=0.0, step=1.0)
        tel = c12.number_input("Tel / Internet", min_value=0.0, value=0.0, step=1.0)
        courier = c13.number_input("Courier / Stationary", min_value=0.0, value=0.0, step=1.0)
        rikshaw = c14.number_input("Rikshaw / Bus / Ola", min_value=0.0, value=0.0, step=1.0)

        remarks = st.text_area("Remarks / Purpose", placeholder="Meeting details, site visit purpose, or notes")
        save = st.form_submit_button("Save Entry")

        if save:
            if not entry_date:
                st.error("Please enter Date.")
            elif not make_str(from_loc) or not make_str(to_loc):
                st.error("Please enter From and To.")
            else:
                row = {
                    "id": str(uuid.uuid4())[:8],
                    **employee,
                    "date": format_date_for_sheet(entry_date),
                    "from": make_str(from_loc),
                    "to": make_str(to_loc),
                    "vehicle": make_str(vehicle),
                    "company": make_str(company),
                    "contact": make_str(contact),
                    "invoice": make_str(invoice),
                    "toll": toll,
                    "fuel": fuel,
                    "lodging": lodging,
                    "food": food,
                    "tel": tel,
                    "courier": courier,
                    "rikshaw": rikshaw,
                    "remarks": make_str(remarks),
                    "created_at": current_stamp(),
                }
                row["total"] = entry_total(row)
                try:
                    append_entry(row)
                    st.success("Entry saved successfully.")
                    st.cache_resource.clear()
                    st.rerun()
                except Exception as exc:
                    st.error("Entry not saved. Google Sheet error:")
                    st.code(str(exc))

    df = read_entries()
    filtered = df.copy()
    if employee["empName"] and not filtered.empty:
        filtered = filtered[filtered["empName"].astype(str) == employee["empName"]]
    if employee["reportMonthText"] and not filtered.empty:
        filtered = filtered[filtered["reportMonthText"].astype(str) == employee["reportMonthText"]]

    for col in AMOUNT_COLUMNS + ["total"]:
        if col not in filtered.columns:
            filtered[col] = 0

    total_entries = len(filtered)
    sum_toll = filtered["toll"].apply(money).sum() if total_entries else 0
    sum_fuel = filtered["fuel"].apply(money).sum() if total_entries else 0
    sum_stay_food = (filtered["lodging"].apply(money).sum() + filtered["food"].apply(money).sum()) if total_entries else 0
    grand_total = filtered["total"].apply(money).sum() if total_entries else 0

    st.markdown('<div class="card-box"><h2 class="card-title">Expense Summary</h2>', unsafe_allow_html=True)
    s1, s2, s3, s4, s5 = st.columns(5)
    with s1: summary_box("Total Entries", str(total_entries))
    with s2: summary_box("Toll / Parking", rupee(sum_toll))
    with s3: summary_box("Fuel", rupee(sum_fuel))
    with s4: summary_box("Lodging + Food", rupee(sum_stay_food))
    with s5: summary_box("Grand Total", rupee(grand_total))

    b1, b2, b3, b4, b5 = st.columns([1.4, 1.1, 1.0, 1.1, 1.1])
    with b1:
        try:
            excel_bytes = generate_excel(df, employee)
            st.download_button(
                "Download Same Excel Format",
                data=excel_bytes,
                file_name=f"Travelling_Expenses_{employee['empName'].replace(' ', '_').replace('.', '')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as exc:
            st.button("Download Same Excel Format", disabled=True, use_container_width=True)
            st.caption(str(exc))
    with b2:
        st.download_button("Export Simple CSV", filtered.to_csv(index=False).encode("utf-8"), "travel_expense_backup.csv", "text/csv", use_container_width=True)
    with b3:
        if st.button("Refresh", use_container_width=True):
            st.rerun()
    st.markdown('<p class="muted">The Excel download uses your attached workbook template. Format is kept as-is; only employee details and entry cells are filled.</p></div>', unsafe_allow_html=True)

    st.markdown('<div class="card-box"><h2 class="card-title">Saved Entries</h2>', unsafe_allow_html=True)
    if filtered.empty:
        st.info("No entries saved for this employee and period.")
    else:
        display = filtered[["date", "from", "to", "company", "contact", "invoice", "vehicle", "toll", "fuel", "lodging", "food", "tel", "courier", "rikshaw", "total", "remarks", "id"]].copy()
        display = display.rename(columns={
            "date": "Date", "from": "From", "to": "To", "company": "Company / Consultant", "contact": "Contact", "invoice": "Invoice",
            "vehicle": "Vehicle", "toll": "Toll", "fuel": "Fuel", "lodging": "Lodging", "food": "Food", "tel": "Tel",
            "courier": "Courier", "rikshaw": "Rikshaw", "total": "Total", "remarks": "Remarks", "id": "EntryID"
        })
        st.dataframe(display, use_container_width=True, hide_index=True, height=280)
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="card-box"><h2 class="card-title">Edit / Delete Entry</h2>', unsafe_allow_html=True)
    if filtered.empty:
        st.info("No entry available to edit.")
    else:
        options = []
        for _, r in filtered.iterrows():
            label = f"{r.get('date','')} | {r.get('from','')} to {r.get('to','')} | {r.get('company','')} | ID: {r.get('id','')}"
            options.append((label, str(r.get("id", ""))))
        selected_label = st.selectbox("Select entry to edit", [x[0] for x in options])
        selected_id = dict(options).get(selected_label, "")
        selected_rows = filtered[filtered["id"].astype(str) == selected_id]
        if not selected_rows.empty:
            selected = selected_rows.iloc[0].to_dict()
            with st.form("edit_form"):
                e1, e2, e3, e4 = st.columns(4)
                edit_date = e1.date_input("Date", value=parse_entry_date(selected.get("date")))
                edit_from = e2.text_input("From", value=make_str(selected.get("from")))
                edit_to = e3.text_input("To", value=make_str(selected.get("to")))
                vehicle_options = ["", "2 Wheeler", "4 Wheeler", "Public Transport", "Train", "Bus", "Flight", "Other"]
                current_vehicle = make_str(selected.get("vehicle"))
                edit_vehicle = e4.selectbox("2 Wheeler / 4 Wheeler", vehicle_options, index=vehicle_options.index(current_vehicle) if current_vehicle in vehicle_options else 0)

                e5, e6 = st.columns(2)
                edit_company = e5.text_input("Company Name / Contact / MEP Consultant", value=make_str(selected.get("company")))
                edit_contact = e6.text_input("Contact Person / Meeting With", value=make_str(selected.get("contact")))

                e7, e8, e9, e10 = st.columns(4)
                edit_invoice = e7.text_input("Invoice No.", value=make_str(selected.get("invoice")))
                edit_toll = e8.number_input("Toll / Parking", min_value=0.0, value=money(selected.get("toll")), step=1.0)
                edit_fuel = e9.number_input("Petrol / Diesel", min_value=0.0, value=money(selected.get("fuel")), step=1.0)
                edit_lodging = e10.number_input("Lodging / Boarding", min_value=0.0, value=money(selected.get("lodging")), step=1.0)

                e11, e12, e13, e14 = st.columns(4)
                edit_food = e11.number_input("Food / Beverages", min_value=0.0, value=money(selected.get("food")), step=1.0)
                edit_tel = e12.number_input("Tel / Internet", min_value=0.0, value=money(selected.get("tel")), step=1.0)
                edit_courier = e13.number_input("Courier / Stationary", min_value=0.0, value=money(selected.get("courier")), step=1.0)
                edit_rikshaw = e14.number_input("Rikshaw / Bus / Ola", min_value=0.0, value=money(selected.get("rikshaw")), step=1.0)
                edit_remarks = st.text_area("Remarks / Purpose", value=make_str(selected.get("remarks")))

                u1, u2 = st.columns(2)
                update_clicked = u1.form_submit_button("Update Selected Entry")
                delete_clicked = u2.form_submit_button("Delete Selected Entry")

                if update_clicked:
                    updated = {
                        "id": selected_id,
                        **employee,
                        "date": format_date_for_sheet(edit_date),
                        "from": make_str(edit_from),
                        "to": make_str(edit_to),
                        "vehicle": make_str(edit_vehicle),
                        "company": make_str(edit_company),
                        "contact": make_str(edit_contact),
                        "invoice": make_str(edit_invoice),
                        "toll": edit_toll,
                        "fuel": edit_fuel,
                        "lodging": edit_lodging,
                        "food": edit_food,
                        "tel": edit_tel,
                        "courier": edit_courier,
                        "rikshaw": edit_rikshaw,
                        "remarks": make_str(edit_remarks),
                        "created_at": make_str(selected.get("created_at")) or current_stamp(),
                    }
                    updated["total"] = entry_total(updated)
                    try:
                        if update_entry_by_id(selected_id, updated):
                            st.success("Entry updated successfully.")
                            st.cache_resource.clear()
                            st.rerun()
                        else:
                            st.error("Entry not found.")
                    except Exception as exc:
                        st.error("Update failed.")
                        st.code(str(exc))

                if delete_clicked:
                    try:
                        delete_entry_by_id(selected_id)
                        st.success("Entry deleted successfully.")
                        st.cache_resource.clear()
                        st.rerun()
                    except Exception as exc:
                        st.error("Delete failed.")
                        st.code(str(exc))
    st.markdown('</div>', unsafe_allow_html=True)


if __name__ == "__main__":
    main()
