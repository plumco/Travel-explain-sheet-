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

VEHICLE_OPTIONS = [
    "",
    "2 Wheeler",
    "4 Wheeler",
    "Public Transport",
    "Train",
    "Bus",
    "Flight",
    "Other",
]

HEADERS = [
    "id",
    "empName",
    "designation",
    "location",
    "reportMonthText",
    "date",
    "from",
    "to",
    "vehicle",
    "company",
    "contact",
    "invoice",
    "toll",
    "fuel",
    "lodging",
    "food",
    "tel",
    "courier",
    "rikshaw",
    "remarks",
    "total",
    "created_at",
]

AMOUNT_COLUMNS = ["toll", "fuel", "lodging", "food", "tel", "courier", "rikshaw"]

CSS = """
<style>
    .stApp {
        background: #f4f7fb;
        color: #1f2937;
        font-family: Arial, sans-serif;
    }

    .main .block-container {
        max-width: 1180px;
        padding-top: 0;
        padding-bottom: 34px;
    }

    .html-header {
        background: #1e40af;
        color: white;
        padding: 18px 16px;
        text-align: center;
        margin: 0 -2rem 18px -2rem;
    }

    .html-header .brand {
        display: inline-block;
        margin-bottom: 8px;
        padding: 5px 10px;
        border: 1px solid rgba(255,255,255,.55);
        border-radius: 999px;
        font-size: 12px;
        letter-spacing: .3px;
        background: rgba(255,255,255,.12);
    }

    .html-header h1 {
        margin: 0;
        font-size: 22px;
        line-height: 1.1;
        font-weight: 800;
    }

    .html-header p {
        margin: 6px 0 0;
        font-size: 14px;
        opacity: .9;
    }

    .card-box {
        background: white;
        border-radius: 14px;
        padding: 16px;
        margin-bottom: 16px;
        box-shadow: 0 4px 14px rgba(15, 23, 42, .08);
        border: 0;
    }

    div[data-testid="stForm"] {
        background: white;
        border-radius: 14px;
        padding: 16px;
        margin-bottom: 16px;
        box-shadow: 0 4px 14px rgba(15, 23, 42, .08);
        border: 0;
    }

    .card-title {
        margin: 0 0 14px 0;
        font-size: 22px;
        font-weight: 800;
        color: #111827;
    }

    .muted {
        color: #64748b;
        font-size: 13px;
        margin-top: 8px;
    }

    .ok-text {
        color: #166534;
        font-weight: 700;
        font-size: 13px;
    }

    .warn-text {
        color: #92400e;
        font-weight: 700;
        font-size: 13px;
    }

    .error-text {
        color: #b91c1c;
        font-weight: 700;
        font-size: 13px;
    }

    label, [data-testid="stWidgetLabel"] p {
        font-size: 12px !important;
        font-weight: 700 !important;
        color: #374151 !important;
        margin-bottom: 5px !important;
    }

    input, textarea, select {
        border: 1px solid #cbd5e1 !important;
        border-radius: 8px !important;
        font-size: 14px !important;
        background: white !important;
    }

    textarea {
        min-height: 70px !important;
    }

    div[data-baseweb="input"] > div,
    div[data-baseweb="select"] > div,
    div[data-baseweb="textarea"] > div {
        border-radius: 8px !important;
        border-color: #cbd5e1 !important;
        background: white !important;
    }

    .stButton>button,
    .stDownloadButton>button,
    div[data-testid="stFormSubmitButton"] button {
        border: 0;
        border-radius: 8px;
        padding: 10px 14px;
        font-size: 14px;
        font-weight: 800;
        min-height: 42px;
    }

    div[data-testid="stFormSubmitButton"] button {
        background: #2563eb;
        color: white;
    }

    .stDownloadButton>button {
        background: #16a34a;
        color: white;
    }

    button[kind="secondary"] {
        background: #e5e7eb;
        color: #111827;
    }

    .summary-box {
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        border-radius: 12px;
        padding: 12px;
        text-align: center;
        min-height: 74px;
    }

    .summary-box span {
        display: block;
        font-size: 12px;
        color: #475569;
    }

    .summary-box strong {
        display: block;
        margin-top: 5px;
        font-size: 18px;
        color: #1e3a8a;
    }

    [data-testid="stDataFrame"] {
        border-radius: 8px;
        overflow: hidden;
    }

    .block-gap {
        height: 8px;
    }

    @media (max-width: 800px) {
        .html-header h1 {
            font-size: 19px;
        }
    }
</style>
"""


def make_str(value):
    if value is None:
        return ""
    return str(value).strip()


def money(value):
    try:
        if value in (None, ""):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def current_stamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def calc_total(row):
    return sum(money(row.get(col, 0)) for col in AMOUNT_COLUMNS)


def default_employee():
    return {
        "empName": "Mr.Umesh Nikam",
        "designation": "Executive-Technical Support",
        "location": "Pune",
        "reportMonthText": "01st Jun 2026 To 30th Jun 2026",
    }


def parse_entry_date(value):
    text = make_str(value)
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except Exception:
            pass
    return date.today()


def summary_box(label, value):
    st.markdown(
        f'<div class="summary-box"><span>{label}</span><strong>{value}</strong></div>',
        unsafe_allow_html=True,
    )


def get_secret_value(name, default_value=""):
    try:
        return st.secrets.get(name, default_value)
    except Exception:
        return default_value


@st.cache_resource(show_spinner=False)
def get_google_worksheet():
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        if "gcp_service_account" not in st.secrets:
            raise RuntimeError("Missing [gcp_service_account] in Streamlit Secrets.")

        spreadsheet_id = ""
        if "gsheets" in st.secrets:
            spreadsheet_id = st.secrets["gsheets"].get("spreadsheet_id", "")

        if not spreadsheet_id:
            raise RuntimeError("Missing [gsheets] spreadsheet_id in Streamlit Secrets.")

        worksheet_name = get_secret_value("WORKSHEET_NAME", WORKSHEET_NAME_DEFAULT)
        worksheet_name = worksheet_name or WORKSHEET_NAME_DEFAULT

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]

        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]),
            scopes=scopes,
        )

        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(spreadsheet_id)

        try:
            ws = spreadsheet.worksheet(worksheet_name)
        except Exception:
            ws = spreadsheet.add_worksheet(
                title=worksheet_name,
                rows=1000,
                cols=len(HEADERS) + 5,
            )
            ws.append_row(HEADERS, value_input_option="USER_ENTERED")
            return ws

        values = ws.get_all_values()

        if not values:
            ws.append_row(HEADERS, value_input_option="USER_ENTERED")
        elif values[0] != HEADERS:
            ws.clear()
            ws.update("A1", [HEADERS])

        return ws

    except Exception as e:
        st.session_state["gsheet_error"] = traceback.format_exc()
        return None


def read_entries():
    ws = get_google_worksheet()

    if ws is None:
        return pd.DataFrame(columns=HEADERS)

    try:
        records = ws.get_all_records()
        df = pd.DataFrame(records)
    except Exception:
        df = pd.DataFrame(columns=HEADERS)

    for col in HEADERS:
        if col not in df.columns:
            df[col] = ""

    df = df[HEADERS].fillna("")
    return df


def append_entry(row):
    ws = get_google_worksheet()

    if ws is None:
        raise RuntimeError("Google Sheets not connected.")

    row = {col: row.get(col, "") for col in HEADERS}
    ws.append_row([row[col] for col in HEADERS], value_input_option="USER_ENTERED")


def rewrite_entries(df):
    ws = get_google_worksheet()

    if ws is None:
        raise RuntimeError("Google Sheets not connected.")

    for col in HEADERS:
        if col not in df.columns:
            df[col] = ""

    df = df[HEADERS].fillna("").astype(str)
    ws.clear()
    ws.update("A1", [HEADERS] + df.values.tolist())


def update_entry(entry_id, updated_row):
    df = read_entries()
    mask = df["id"].astype(str) == str(entry_id)

    if not mask.any():
        return False

    for col in HEADERS:
        df.loc[mask, col] = updated_row.get(col, "")

    rewrite_entries(df)
    return True


def delete_entry(entry_id):
    df = read_entries()
    new_df = df[df["id"].astype(str) != str(entry_id)].copy()
    rewrite_entries(new_df)


def copy_row_style(ws, source_row, target_row):
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


def safe_set(ws, cell_ref, value):
    cell = ws[cell_ref]
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            ws.cell(merged_range.min_row, merged_range.min_col).value = value
            return
    cell.value = value


def clear_cell(ws, row, col):
    coord = ws.cell(row, col).coordinate
    for merged_range in ws.merged_cells.ranges:
        if coord in merged_range:
            if row == merged_range.min_row and col == merged_range.min_col:
                ws.cell(row, col).value = None
            return
    ws.cell(row, col).value = None


def generate_excel(df, employee):
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
        extra_rows = entry_count - fixed_rows
        ws.insert_rows(template_end + 1, extra_rows)
        for r in range(template_end + 1, template_end + 1 + extra_rows):
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


def show_connection_status():
    st.markdown('<div class="card-box"><h2 class="card-title">Google Sheet Connection</h2>', unsafe_allow_html=True)

    ws = get_google_worksheet()

    if ws is not None:
        st.markdown('<p class="ok-text">Storage: Google Sheets connected. Entries are saved live.</p>', unsafe_allow_html=True)
    else:
        st.markdown('<p class="error-text">Storage: Google Sheets not connected.</p>', unsafe_allow_html=True)
        with st.expander("Show Google Sheet connection error"):
            st.code(st.session_state.get("gsheet_error", "No error captured."))

    sheet_id = ""
    if "gsheets" in st.secrets:
        sheet_id = st.secrets["gsheets"].get("spreadsheet_id", "")

    worksheet_name = get_secret_value("WORKSHEET_NAME", WORKSHEET_NAME_DEFAULT)
    st.caption(f"Sheet ID: {sheet_id} | Worksheet: {worksheet_name}")
    st.markdown("</div>", unsafe_allow_html=True)


def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)

    st.markdown(
        """
        <div class="html-header">
            <div class="brand">HULIOT INDIA</div>
            <h1>Travel Expense Entry App</h1>
            <p>Save entries live in Google Sheets and download in your original Excel format.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if "employee" not in st.session_state:
        st.session_state.employee = default_employee()

    show_connection_status()

    st.markdown('<div class="card-box"><h2 class="card-title">Employee Details</h2>', unsafe_allow_html=True)

    ec1, ec2, ec3, ec4 = st.columns(4)

    emp_name = ec1.text_input("Name", value=st.session_state.employee.get("empName", ""), key="emp_name")
    designation = ec2.text_input("Designation", value=st.session_state.employee.get("designation", ""), key="emp_designation")
    location = ec3.text_input("Location", value=st.session_state.employee.get("location", ""), key="emp_location")
    report_text = ec4.text_input("Month / Period Text", value=st.session_state.employee.get("reportMonthText", ""), key="emp_report_text")

    st.markdown('<p class="muted">These details are used in the Excel header and for filtering your saved entries.</p></div>', unsafe_allow_html=True)

    employee = {
        "empName": make_str(emp_name),
        "designation": make_str(designation),
        "location": make_str(location),
        "reportMonthText": make_str(report_text),
    }
    st.session_state.employee = employee

    with st.form("entry_form", clear_on_submit=True):
        st.markdown('<h2 class="card-title">Add Travel Expense</h2>', unsafe_allow_html=True)

        c1, c2, c3, c4 = st.columns(4)
        entry_date = c1.date_input("Date", value=date.today())
        from_loc = c2.text_input("From", placeholder="Start location")
        to_loc = c3.text_input("To", placeholder="End location")
        vehicle = c4.selectbox("2 Wheeler / 4 Wheeler", VEHICLE_OPTIONS, index=1)

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

        save_clicked = st.form_submit_button("Save Entry")

        if save_clicked:
            try:
                if not make_str(from_loc) or not make_str(to_loc):
                    st.error("Please enter From and To.")
                else:
                    row = {
                        "id": str(uuid.uuid4())[:8],
                        "empName": employee["empName"],
                        "designation": employee["designation"],
                        "location": employee["location"],
                        "reportMonthText": employee["reportMonthText"],
                        "date": entry_date.strftime("%Y-%m-%d"),
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
                    row["total"] = calc_total(row)

                    append_entry(row)
                    st.success("Entry saved successfully in Google Sheet.")
                    st.rerun()

            except Exception as e:
                st.error(f"Entry not saved. Google Sheet error: {e}")
                with st.expander("Show full error"):
                    st.code(traceback.format_exc())

    df = read_entries()

    filtered = df.copy()

    if employee["empName"]:
        filtered = filtered[filtered["empName"].astype(str) == employee["empName"]]

    if employee["reportMonthText"]:
        filtered = filtered[filtered["reportMonthText"].astype(str) == employee["reportMonthText"]]

    for col in AMOUNT_COLUMNS + ["total"]:
        if col not in filtered.columns:
            filtered[col] = 0

    total_entries = len(filtered)
    sum_toll = filtered["toll"].apply(money).sum() if total_entries else 0
    sum_fuel = filtered["fuel"].apply(money).sum() if total_entries else 0
    sum_stay_food = (
        filtered["lodging"].apply(money).sum() + filtered["food"].apply(money).sum()
    ) if total_entries else 0
    grand_total = filtered["total"].apply(money).sum() if total_entries else 0

    st.markdown('<div class="card-box"><h2 class="card-title">Expense Summary</h2>', unsafe_allow_html=True)

    s1, s2, s3, s4, s5 = st.columns(5)

    with s1:
        summary_box("Total Entries", str(total_entries))
    with s2:
        summary_box("Toll / Parking", f"₹{sum_toll:.2f}")
    with s3:
        summary_box("Fuel", f"₹{sum_fuel:.2f}")
    with s4:
        summary_box("Lodging + Food", f"₹{sum_stay_food:.2f}")
    with s5:
        summary_box("Grand Total", f"₹{grand_total:.2f}")

    st.markdown('<div class="block-gap"></div>', unsafe_allow_html=True)

    b1, b2, b3 = st.columns([1.4, 1.1, 1.1])

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
        except Exception as e:
            st.error(f"Excel download error: {e}")

    with b2:
        st.download_button(
            "Backup CSV",
            data=filtered.to_csv(index=False).encode("utf-8"),
            file_name="travel_expense_backup.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with b3:
        if st.button("Refresh", use_container_width=True):
            st.rerun()

    st.markdown('<p class="muted">Excel download uses your original workbook. Only employee details and entry cells are filled.</p></div>', unsafe_allow_html=True)

    st.markdown('<div class="card-box"><h2 class="card-title">Saved Entries</h2>', unsafe_allow_html=True)

    if filtered.empty:
        st.info("No entries saved for this employee and period.")
    else:
        display_cols = [
            "date",
            "from",
            "to",
            "company",
            "contact",
            "invoice",
            "vehicle",
            "toll",
            "fuel",
            "lodging",
            "food",
            "tel",
            "courier",
            "rikshaw",
            "total",
            "remarks",
            "id",
        ]

        show_df = filtered[[c for c in display_cols if c in filtered.columns]].copy()
        show_df = show_df.rename(columns={
            "date": "Date",
            "from": "From",
            "to": "To",
            "company": "Company / Consultant",
            "contact": "Contact",
            "invoice": "Invoice",
            "vehicle": "Vehicle",
            "toll": "Toll",
            "fuel": "Fuel",
            "lodging": "Lodging",
            "food": "Food",
            "tel": "Tel",
            "courier": "Courier",
            "rikshaw": "Rikshaw",
            "total": "Total",
            "remarks": "Remarks",
            "id": "EntryID",
        })

        st.dataframe(show_df, use_container_width=True, hide_index=True, height=330)

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card-box"><h2 class="card-title">Edit / Delete Entry</h2>', unsafe_allow_html=True)

    if filtered.empty:
        st.info("No entry available to edit.")
    else:
        edit_options = []
        for _, r in filtered.iterrows():
            label = f"{r.get('date', '')} | {r.get('from', '')} to {r.get('to', '')} | {r.get('company', '')} | ID: {r.get('id', '')}"
            edit_options.append((label, str(r.get("id", ""))))

        selected_label = st.selectbox("Select entry to edit", [x[0] for x in edit_options])
        selected_id = dict(edit_options).get(selected_label, "")

        selected_rows = filtered[filtered["id"].astype(str) == selected_id]

        if not selected_rows.empty:
            selected = selected_rows.iloc[0].to_dict()

            with st.form("edit_form"):
                e1, e2, e3, e4 = st.columns(4)

                edit_date = e1.date_input("Date", value=parse_entry_date(selected.get("date")), key=f"edit_date_{selected_id}")
                edit_from = e2.text_input("From", value=make_str(selected.get("from")), key=f"edit_from_{selected_id}")
                edit_to = e3.text_input("To", value=make_str(selected.get("to")), key=f"edit_to_{selected_id}")

                current_vehicle = make_str(selected.get("vehicle"))
                vehicle_index = VEHICLE_OPTIONS.index(current_vehicle) if current_vehicle in VEHICLE_OPTIONS else 0
                edit_vehicle = e4.selectbox("2 Wheeler / 4 Wheeler", VEHICLE_OPTIONS, index=vehicle_index, key=f"edit_vehicle_{selected_id}")

                e5, e6 = st.columns(2)
                edit_company = e5.text_input("Company Name / Contact / MEP Consultant", value=make_str(selected.get("company")), key=f"edit_company_{selected_id}")
                edit_contact = e6.text_input("Contact Person / Meeting With", value=make_str(selected.get("contact")), key=f"edit_contact_{selected_id}")

                e7, e8, e9, e10 = st.columns(4)
                edit_invoice = e7.text_input("Invoice No.", value=make_str(selected.get("invoice")), key=f"edit_invoice_{selected_id}")
                edit_toll = e8.number_input("Toll / Parking", min_value=0.0, value=money(selected.get("toll")), step=1.0, key=f"edit_toll_{selected_id}")
                edit_fuel = e9.number_input("Petrol / Diesel", min_value=0.0, value=money(selected.get("fuel")), step=1.0, key=f"edit_fuel_{selected_id}")
                edit_lodging = e10.number_input("Lodging / Boarding", min_value=0.0, value=money(selected.get("lodging")), step=1.0, key=f"edit_lodging_{selected_id}")

                e11, e12, e13, e14 = st.columns(4)
                edit_food = e11.number_input("Food / Beverages", min_value=0.0, value=money(selected.get("food")), step=1.0, key=f"edit_food_{selected_id}")
                edit_tel = e12.number_input("Tel / Internet", min_value=0.0, value=money(selected.get("tel")), step=1.0, key=f"edit_tel_{selected_id}")
                edit_courier = e13.number_input("Courier / Stationary", min_value=0.0, value=money(selected.get("courier")), step=1.0, key=f"edit_courier_{selected_id}")
                edit_rikshaw = e14.number_input("Rikshaw / Bus / Ola", min_value=0.0, value=money(selected.get("rikshaw")), step=1.0, key=f"edit_rikshaw_{selected_id}")

                edit_remarks = st.text_area("Remarks / Purpose", value=make_str(selected.get("remarks")), key=f"edit_remarks_{selected_id}")

                u1, u2 = st.columns(2)

                update_clicked = u1.form_submit_button("Update Selected Entry")
                delete_clicked = u2.form_submit_button("Delete Selected Entry")

                if update_clicked:
                    try:
                        updated = {
                            "id": selected_id,
                            "empName": employee["empName"],
                            "designation": employee["designation"],
                            "location": employee["location"],
                            "reportMonthText": employee["reportMonthText"],
                            "date": edit_date.strftime("%Y-%m-%d"),
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
                        updated["total"] = calc_total(updated)

                        if update_entry(selected_id, updated):
                            st.success("Entry updated successfully.")
                            st.rerun()
                        else:
                            st.error("Entry not found.")

                    except Exception as e:
                        st.error(f"Update failed: {e}")
                        with st.expander("Show full error"):
                            st.code(traceback.format_exc())

                if delete_clicked:
                    try:
                        delete_entry(selected_id)
                        st.success("Entry deleted successfully.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Delete failed: {e}")
                        with st.expander("Show full error"):
                            st.code(traceback.format_exc())

    st.markdown("</div>", unsafe_allow_html=True)


if __name__ == "__main__":
    main()
