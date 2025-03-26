import streamlit as st
import datetime
import pandas as pd
from calendar import monthrange
import random
import io

# --- Helper Functions ---

def generate_all_dates(year, month):
    """Generate a list of all dates for the given month."""
    _, num_days = monthrange(year, month)
    return [datetime.date(year, month, d) for d in range(1, num_days + 1)]

def generate_dates(year, month):
    """Generate all weekdays (Monday to Friday) for the given month."""
    _, num_days = monthrange(year, month)
    start_date = datetime.date(year, month, 1)
    end_date = datetime.date(year, month, num_days)
    return [
        start_date + datetime.timedelta(days=i)
        for i in range((end_date - start_date).days + 1)
        if (start_date + datetime.timedelta(days=i)).weekday() < 5
    ]

# --- Fallback Function for Front Desk Shifts ---
def fallback_front_desk(date):
    """
    Fallback for front desk shifts:
    Only use staff with role "Other" (fallback coverage).
    If multiple candidates have the same lowest shift count, one is chosen at random.
    Returns a tuple (name, staff_obj) or (None, None) if no candidate is available.
    """
    available = [s for s in st.session_state.staff if s["role"] == "Other" and date not in s["holidays"]]
    if available:
        min_count = min(s["shift_count"] for s in available)
        tied = [s for s in available if s["shift_count"] == min_count]
        selected = random.choice(tied)
        return selected["name"], selected
    return None, None

# --- Rota Generation Function ---
def generate_rota(dates):
    rows = []
    for date in dates:
        row = {"Date": date.strftime("%d/%m/%Y"), "Day": date.strftime("%A")}
        # If the date is a University closure, mark all tasks as "CLOSED"
        if date in st.session_state.closure_days:
            for col in ["HS_Front_AM", "HS_Front_PM", "LS_Front_AM", "LS_Front_PM",
                        "LibChat_AM", "LibChat_PM", "Phones_AM1", "Phones_AM2", "Phones_PM1", "Phones_PM2"]:
                row[col] = "CLOSED"
            rows.append(row)
            continue

        weekday = date.weekday()  # Monday=0, Friday=4
        is_friday = (weekday == 4)
        
        # --- Health Sciences Front Desk Assignment ---
        available_hs_am = [s for s in st.session_state.staff 
                           if s["role"] == "Health Sciences Front Desk" 
                           and weekday in s["office_days"]
                           and date not in s["holidays"]
                           and date not in s.get("front_assigned_dates", set())]
        if available_hs_am:
            selected = min(available_hs_am, key=lambda s: s["shift_count"])
            row["HS_Front_AM"] = selected["name"]
            selected.setdefault("front_assigned_dates", set()).add(date)
            selected["shift_count"] += 1
        else:
            name_fb, selected_fb = fallback_front_desk(date)
            if name_fb:
                row["HS_Front_AM"] = name_fb + " (Fallback)"
                selected_fb["shift_count"] += 1
            else:
                row["HS_Front_AM"] = "UNASSIGNED"
        if not is_friday:
            available_hs_pm = [s for s in st.session_state.staff 
                               if s["role"] == "Health Sciences Front Desk" 
                               and weekday in s["office_days"]
                               and date not in s["holidays"]
                               and date not in s.get("front_assigned_dates", set())]
            if available_hs_pm:
                selected = min(available_hs_pm, key=lambda s: s["shift_count"])
                row["HS_Front_PM"] = selected["name"]
                selected.setdefault("front_assigned_dates", set()).add(date)
                selected["shift_count"] += 1
            else:
                name_fb, selected_fb = fallback_front_desk(date)
                if name_fb:
                    row["HS_Front_PM"] = name_fb + " (Fallback)"
                    selected_fb["shift_count"] += 1
                else:
                    row["HS_Front_PM"] = "UNASSIGNED"
        
        # --- Life Sciences Front Desk Assignment ---
        available_ls_am = [s for s in st.session_state.staff 
                           if s["role"] == "Life Sciences Front Desk" 
                           and weekday in s["office_days"]
                           and date not in s["holidays"]
                           and date not in s.get("front_assigned_dates", set())]
        if available_ls_am:
            selected = min(available_ls_am, key=lambda s: s["shift_count"])
            row["LS_Front_AM"] = selected["name"]
            selected.setdefault("front_assigned_dates", set()).add(date)
            selected["shift_count"] += 1
        else:
            name_fb, selected_fb = fallback_front_desk(date)
            if name_fb:
                row["LS_Front_AM"] = name_fb + " (Fallback)"
                selected_fb["shift_count"] += 1
            else:
                row["LS_Front_AM"] = "UNASSIGNED"
        if not is_friday:
            available_ls_pm = [s for s in st.session_state.staff 
                               if s["role"] == "Life Sciences Front Desk" 
                               and weekday in s["office_days"]
                               and date not in s["holidays"]
                               and date not in s.get("front_assigned_dates", set())]
            if available_ls_pm:
                selected = min(available_ls_pm, key=lambda s: s["shift_count"])
                row["LS_Front_PM"] = selected["name"]
                selected.setdefault("front_assigned_dates", set()).add(date)
                selected["shift_count"] += 1
            else:
                name_fb, selected_fb = fallback_front_desk(date)
                if name_fb:
                    row["LS_Front_PM"] = name_fb + " (Fallback)"
                    selected_fb["shift_count"] += 1
                else:
                    row["LS_Front_PM"] = "UNASSIGNED"
        
        # --- Lib Chat Assignment (No on-campus restriction) ---
        available_lib = [s for s in st.session_state.staff if date not in s["holidays"]]
        if available_lib:
            selected = min(available_lib, key=lambda s: s["shift_count"])
            row["LibChat_AM"] = selected["name"]
            selected["shift_count"] += 1
        else:
            row["LibChat_AM"] = "UNASSIGNED"
        if available_lib:
            selected = min(available_lib, key=lambda s: s["shift_count"])
            row["LibChat_PM"] = selected["name"]
            selected["shift_count"] += 1
        else:
            row["LibChat_PM"] = "UNASSIGNED"
        
        # --- Phones Assignment (No on-campus restriction; two slots per shift) ---
        available_phones = [s for s in st.session_state.staff if date not in s["holidays"]]
        if available_phones:
            selected = min(available_phones, key=lambda s: s["shift_count"])
            row["Phones_AM1"] = selected["name"]
            selected["shift_count"] += 1
        else:
            row["Phones_AM1"] = "UNASSIGNED"
        if available_phones:
            selected = min(available_phones, key=lambda s: s["shift_count"])
            row["Phones_AM2"] = selected["name"]
            selected["shift_count"] += 1
        else:
            row["Phones_AM2"] = "UNASSIGNED"
        if available_phones:
            selected = min(available_phones, key=lambda s: s["shift_count"])
            row["Phones_PM1"] = selected["name"]
            selected["shift_count"] += 1
        else:
            row["Phones_PM1"] = "UNASSIGNED"
        if available_phones:
            selected = min(available_phones, key=lambda s: s["shift_count"])
            row["Phones_PM2"] = selected["name"]
            selected["shift_count"] += 1
        else:
            row["Phones_PM2"] = "UNASSIGNED"
        
        rows.append(row)
    return pd.DataFrame(rows)

# --- Streamlit App Layout ---

st.title("CHMLS TPO Rota Generator")

# Initialize session state if not set.
if "staff" not in st.session_state:
    st.session_state.staff = []
if "closure_days" not in st.session_state:
    st.session_state.closure_days = []

# Step 1: Select Target Month
st.header("Select Target Month")
target_date = st.date_input("Select any date in the target month:")
year = target_date.year
month = target_date.month
all_dates = generate_all_dates(year, month)
all_date_strings = [d.strftime("%d/%m/%Y") for d in all_dates]

# Step 2: Add Staff Members
st.header("Add Staff Members")
with st.form("staff_form", clear_on_submit=True):
    name = st.text_input("Staff Name")
    role = st.selectbox("Select Role", options=["Health Sciences Front Desk", "Life Sciences Front Desk", "Other"])
    days_map = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    if role in ["Health Sciences Front Desk", "Life Sciences Front Desk"]:
        office_days = st.multiselect("Select Office Days", options=days_map)
    else:
        office_days = days_map
    holidays_selected = st.multiselect("Select Holidays for this Staff Member", options=all_date_strings)
    submitted = st.form_submit_button("Add Staff Member")
    if submitted and name:
        st.session_state.staff.append({
            "name": name,
            "role": role,
            "office_days": [days_map.index(day) for day in office_days],
            "holidays": [datetime.datetime.strptime(d, "%d/%m/%Y").date() for d in holidays_selected],
            "front_assigned_dates": set(),
            "shift_count": 0
        })
        st.success(f"Added staff member: {name} ({role})")

if st.session_state.staff:
    st.subheader("Staff Members Added")
    for s in st.session_state.staff:
        st.write(f"{s['name']} ({s['role']}) - Office Days: {', '.join([days_map[i] for i in s['office_days']])} - Holidays: {', '.join([d.strftime('%d/%m/%Y') for d in s['holidays']])}")

# Step 3: Add Closure Dates
st.header("Add Closure Dates")
closure_selected = st.multiselect("Select Closure Dates", options=all_date_strings)
st.session_state.closure_days = [datetime.datetime.strptime(d, "%d/%m/%Y").date() for d in closure_selected]
if st.session_state.closure_days:
    st.write("Closure Dates:", ", ".join([d.strftime("%d/%m/%Y") for d in st.session_state.closure_days]))

# Step 4: Generate Rota
if st.button("Generate Rota"):
    working_dates = generate_dates(year, month)
    rota_df = generate_rota(working_dates)
    st.subheader("Rota Table")
    st.dataframe(rota_df)
    
    # --- Excel Export with Color Coding ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        rota_df.to_excel(writer, index=False, sheet_name="Rota")
        workbook = writer.book
        worksheet = writer.sheets["Rota"]
        
        # Create a color mapping for each staff member.
        unique_staff = {s["name"] for s in st.session_state.staff}
        colors = ["#FFC7CE", "#C6EFCE", "#FFEB9C", "#D9E1F2", "#FCE4D6", "#E4DFEC", "#B4C7E7", "#FFD966", "#C5E0B4"]
        staff_formats = {}
        for i, name in enumerate(sorted(unique_staff)):
            color = colors[i % len(colors)]
            staff_formats[name] = workbook.add_format({"bg_color": color})
        
        # Iterate through the data and apply color formatting.
        # The header is in row 0, so data starts at row 1.
        for i in range(1, len(rota_df) + 1):
            for j, col in enumerate(rota_df.columns):
                cell_value = str(rota_df.iloc[i-1, j])
                # If the cell value contains a staff name (even as part of a fallback string), apply that staff's format.
                for name, fmt in staff_formats.items():
                    if name in cell_value:
                        worksheet.write(i, j, cell_value, fmt)
                        break
        # Add a legend below the table.
        legend_row = len(rota_df) + 3
        worksheet.write(legend_row, 0, "Legend:")
        row = legend_row + 1
        for name, fmt in staff_formats.items():
            worksheet.write(row, 0, name, fmt)
            row += 1
    output.seek(0)
    st.download_button("Download Excel File", data=output, file_name="rota.xlsx", mime="application/vnd.ms-excel")
    
    # Shift Summary
    summary = [{"Name": s["name"], "Shift Count": s["shift_count"]} for s in st.session_state.staff]
    summary_df = pd.DataFrame(summary)
    st.subheader("Shift Summary")
    st.dataframe(summary_df)

