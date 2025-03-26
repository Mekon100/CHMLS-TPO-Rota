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

def fallback_dedicated(role, date, weekday):
    """
    Fallback for front desk shifts using dedicated staff even if they are already assigned.
    Returns a tuple (name, staff_obj) chosen fairly (lowest shift_count).
    If none available, returns (None, None).
    """
    candidates = [s for s in st.session_state.staff 
                  if s["role"] == role 
                  and weekday in s["office_days"]
                  and date not in s["holidays"]]
    if candidates:
        selected = min(candidates, key=lambda s: s["shift_count"])
        return selected["name"], selected
    return None, None

# --- Rota Generation Function ---
def generate_rota(dates):
    rows = []
    for date in dates:
        row = {"Date": date.strftime("%d/%m/%Y"), "Day": date.strftime("%A")}
        # If the date is a University closure day, mark all tasks as "CLOSED"
        if date in st.session_state.closure_days:
            for col in ["HS_Front_AM", "HS_Front_PM", "LS_Front_AM", "LS_Front_PM",
                        "LibChat_AM", "LibChat_PM", "Phones_AM1", "Phones_AM2", "Phones_PM1", "Phones_PM2"]:
                row[col] = "CLOSED"
            rows.append(row)
            continue

        weekday = date.weekday()  # Monday=0, Friday=4
        is_friday = (weekday == 4)
        
        # --- Health Sciences Front Desk Assignment ---
        # Try normal assignment (avoid double assignment if possible)
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
            # Fallback: allow dedicated staff (even if already assigned)
            name_fb, selected_fb = fallback_dedicated("Health Sciences Front Desk", date, weekday)
            if name_fb:
                row["HS_Front_AM"] = name_fb + " (Fallback)"
                selected_fb["shift_count"] += 1
            else:
                row["HS_Front_AM"] = "UNASSIGNED"
                
        # For PM shift:
        if is_friday:
            row["HS_Front_PM"] = "CLOSED"
        else:
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
                name_fb, selected_fb = fallback_dedicated("Health Sciences Front Desk", date, weekday)
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
            name_fb, selected_fb = fallback_dedicated("Life Sciences Front Desk", date, weekday)
            if name_fb:
                row["LS_Front_AM"] = name_fb + " (Fallback)"
                selected_fb["shift_count"] += 1
            else:
                row["LS_Front_AM"] = "UNASSIGNED"
                
        if is_friday:
            row["LS_Front_PM"] = "CLOSED"
        else:
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
                name_fb, selected_fb = fallback_dedicated("Life Sciences Front Desk", date, weekday)
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

st.title("Monthly Front Desk Rota Generator – Extended")

# Initialize session state if not already set.
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
        office_days = days_map  # Off-site assumed full availability.
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
        
        # Define additional formats.
        warning_format = workbook.add_format({"bg_color": "#FFC7CE"})  # Red for UNASSIGNED.
        closed_format = workbook.add_format({"bg_color": "#D9D9D9"})   # Gray for CLOSED.
        
        # Apply formatting to cells.
        for i in range(1, len(rota_df) + 1):
            for j, col in enumerate(rota_df.columns):
                cell_value = str(rota_df.iloc[i-1, j])
                fmt = None
                if cell_value == "UNASSIGNED":
                    fmt = warning_format
                elif cell_value == "CLOSED":
                    fmt = closed_format
                else:
                    for name, staff_fmt in staff_formats.items():
                        if name in cell_value:
                            fmt = staff_fmt
                            break
                if fmt:
                    worksheet.write(i, j, cell_value, fmt)
                else:
                    worksheet.write(i, j, cell_value)
        
        # Add a legend below the table.
        legend_row = len(rota_df) + 3
        worksheet.write(legend_row, 0, "Legend:")
        row = legend_row + 1
        for name, fmt in staff_formats.items():
            worksheet.write(row, 0, name, fmt)
            row += 1
        worksheet.write(row, 0, "UNASSIGNED", warning_format)
        row += 1
        worksheet.write(row, 0, "CLOSED", closed_format)
    output.seek(0)
    st.download_button("Download Excel File", data=output, file_name="rota.xlsx", mime="application/vnd.ms-excel")
    
    # --- Shift Summary ---
    summary = [{"Name": s["name"], "Shift Count": s["shift_count"]} for s in st.session_state.staff]
    summary_df = pd.DataFrame(summary)
    st.subheader("Shift Summary")
    st.dataframe(summary_df)


