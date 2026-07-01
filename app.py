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

VEHICLE_OPTIONS = ["", "2 Wheeler", "4 Wheeler", "Public Transport", "Train", "Bus", "Flight", "Other"]

HEADERS = [
    "id", "empName", "designation", "location", "reportMonthText", "date", "from", "to",
    "vehicle", "company", "contact", "invoice", "toll", "fuel", "lodging", "food",
    "tel", "courier", "rikshaw", "remarks", "total", "created_at",
]

AMOUNT_COLUMNS = ["toll", "fuel", "lodging", "food", "tel", "courier", "rikshaw"]

CSS = """<style>
    .stApp { background: #f4f7fb; color: #1f2937; font-family: Arial, sans-serif; }
    .card-box { background: white; border-radius: 14px; padding: 16px; margin-bottom: 16px; box-shadow: 0 4px 14px rgba(15, 23, 42, .08); }
    .card-title { margin: 0 0 14px 0; font-size: 22px; font-weight: 800; color: #111827; }
</style>"""

def make_str(value): return str(value).strip() if value is not None else ""
def money(value):
    try: return float(value) if value not in (None, "") else 0.0
    except: return 0.0

def current_stamp(): return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def calc_total(row): return sum(money(row.get(col, 0)) for col in AMOUNT_COLUMNS)
def default_employee():
    return {"empName": "Mr.Umesh Nikam", "designation": "Executive-Technical Support", "location": "Pune", "reportMonthText": "01st Jun 2026 To 30th Jun 2026"}

def parse_entry_date(value):
    text = make_str(value)
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y"):
        try: return datetime.strptime(text, fmt).date()
        except: pass
    return date.today()

# --- Placeholder functions for Google Sheets (Ensure your existing logic remains) ---
def get_google_worksheet(): pass 
def read_entries(): return pd.DataFrame(columns=HEADERS)
def append_entry(row): pass
def update_entry(id, row): pass
def delete_entry(id): pass
def generate_excel(df, emp): return b""

def main():
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    
    if "employee" not in st.session_state: st.session_state.employee = default_employee()
    if "copy_data" not in st.session_state: st.session_state.copy_data = None

    # Employee Details
    st.markdown('<div class="card-box"><h2 class="card-title">Employee Details</h2>', unsafe_allow_html=True)
    ec1, ec2, ec3, ec4 = st.columns(4)
    emp_name = ec1.text_input("Name", value=st.session_state.employee.get("empName", ""), key="emp_name")
    designation = ec2.text_input("Designation", value=st.session_state.employee.get("designation", ""), key="emp_designation")
    location = ec3.text_input("Location", value=st.session_state.employee.get("location", ""), key="emp_location")
    report_text = ec4.text_input("Month / Period Text", value=st.session_state.employee.get("reportMonthText", ""), key="emp_report_text")
    st.session_state.employee = {"empName": emp_name, "designation": designation, "location": location, "reportMonthText": report_text}
    st.markdown("</div>", unsafe_allow_html=True)

    # --- COPY ENTRY SECTION ---
    df = read_entries()
    filtered = df # Add your existing filtering logic here
    
    st.markdown('<div class="card-box"><h2 class="card-title">Copy Existing Entry</h2>', unsafe_allow_html=True)
    copy_options = {"--- Select an entry ---": None}
    for _, r in filtered.iterrows():
        copy_options[f"{r.get('date')} | {r.get('from')} to {r.get('to')}"] = r.to_dict()
    
    selected_copy = st.selectbox("Select entry to duplicate", list(copy_options.keys()))
    if st.button("Load Selected into Form"):
        st.session_state.copy_data = copy_options[selected_copy]
        st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)

    # --- ADD TRAVEL EXPENSE FORM ---
    cd = st.session_state.copy_data or {}
    with st.form("entry_form", clear_on_submit=True):
        st.markdown('<h2 class="card-title">Add Travel Expense</h2>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        entry_date = c1.date_input("Date", value=date.today())
        from_loc = c2.text_input("From", value=cd.get("from", ""))
        to_loc = c3.text_input("To", value=cd.get("to", ""))
        vehicle_idx = VEHICLE_OPTIONS.index(cd.get("vehicle")) if cd.get("vehicle") in VEHICLE_OPTIONS else 1
        vehicle = c4.selectbox("Vehicle", VEHICLE_OPTIONS, index=vehicle_idx)
        
        # ... (Include your other fields here using value=cd.get("field_name", "")) ...
        
        if st.form_submit_button("Save Entry"):
            # ... (Your existing save logic) ...
            st.session_state.copy_data = None
            st.rerun()

if __name__ == "__main__":
    main()
