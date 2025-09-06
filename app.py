import streamlit as st
import openpyxl
import random
from datetime import date
import io
from supabase import create_client, Client
import plotly.express as px
import time

# Initialize Supabase client using Streamlit secrets
supabase_url = st.secrets["SUPABASE_URL"]
supabase_key = st.secrets["SUPABASE_KEY"]
supabase: Client = create_client(supabase_url, supabase_key)

# Pre-populated data
pre_pop_employees = ['AH', 'LS', 'DS', 'KL', 'TH', 'LAO', 'AL', 'HS', 'AG', 'CB']
pre_unavailable = {
    'Monday': ['DS'],
    'Tuesday': ['LAO', 'CB'],
    'Wednesday': ['DS', 'AH', 'CB'],
    'Thursday': ['CB'],
    'Friday': ['CB']
}

st.title("Weekly Employee Schedule Generator for Mammography Unit")

# Input for available employees this week
available_week = st.multiselect(
    "Initials of employees available for the current week",
    options=pre_pop_employees,
    default=pre_pop_employees
)

# Input for unavailable whole week
unavailable_whole_week = st.multiselect(
    "Initials of employees unavailable for the whole current week",
    options=available_week,
    default=[]
)

available_employees = [emp for emp in available_week if emp not in unavailable_whole_week]

# Multiselect for unavailable per day
days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
with st.expander("Assign Unavailable Employees per Day", expanded=True):
    unavailable_per_day = {}
    for day in days:
        default_values = [emp for emp in pre_unavailable.get(day, []) if emp in available_employees]
        unavailable_per_day[day] = st.multiselect(
            f"Initials of employees unavailable on {day}",
            options=available_employees,
            default=default_values
        )

# MDK Overview Bar Graph
with st.expander("MDK Assignments Overview (Bar Graph)"):
    response = supabase.table("mdk_assignments").select("employee").execute()
    assignments = response.data if response.data else []
    mdk_counts = {}
    for assignment in assignments:
        emp = assignment['employee']
        mdk_counts[emp] = mdk_counts.get(emp, 0) + 1
    mdk_counts = {k: v for k, v in mdk_counts.items() if v > 0}
    if mdk_counts:
        employees = list(mdk_counts.keys())
        counts = list(mdk_counts.values())
        fig = px.bar(x=employees, y=counts, labels={'x': 'Employee', 'y': 'MDK Assignments'}, title="MDK Assignments per Employee")
        st.plotly_chart(fig)
    else:
        st.info("No MDK assignments in history yet.")

# Historical Schedules Upload (last 8 weeks)
current_week = date.today().isocalendar()[1]
with st.expander("Historical Schedules (Last 8 Weeks)"):
    bucket_files = supabase.storage.from_("schedules").list()
    file_names = [f['name'] for f in bucket_files] if bucket_files else []

    for i in range(1, 9):
        week = current_week - i
        file_name = f"week_{week}.xlsx"
        st.write(f"Week {week}")
        status = "uploaded" if file_name in file_names else "not uploaded"
        st.write(f"Current file: {status}")
        uploader = st.file_uploader(f"Upload/replace schedule for week {week}", type="xlsx", key=f"hist_{week}")
        if uploader:
            file_content = uploader.getvalue()
            supabase.storage.from_("schedules").upload(file_name, file_content, {"upsert": "true"})
            st.success(f"Uploaded {file_name}")
            time.sleep(1)

            downloaded = supabase.storage.from_("schedules").download(file_name)
            if downloaded:
                wb = openpyxl.load_workbook(io.BytesIO(downloaded))
                sheet = wb["Blad1"] if "Blad1" in wb.sheetnames else wb.active
                mdk_cols = {'Monday': 'D', 'Tuesday': 'H', 'Thursday': 'P'}
                parsed_days = 0
                for day, col in mdk_cols.items():
                    cell_value = sheet.cell(row=3, column=openpyxl.utils.column_index_from_string(col)).value
                    if cell_value and isinstance(cell_value, str) and cell_value.strip() in pre_pop_employees:
                        try:
                            supabase.table("mdk_assignments").upsert({"week": week, "day": day, "employee": cell_value.strip()}).execute()
                            parsed_days += 1
                        except Exception as e:
                            st.error(f"Failed to upsert {day} MDK assignment: {e}")
                st.success(f"Parsed and updated MDK assignments for week {week}. {parsed_days} days added.")


# Use percentages (0-100) for work rates
default_work_rates = {emp: 100 for emp in pre_pop_employees}
response = supabase.table("work_rates").select("*").execute()
db_work_rates = {row['employee']: row['rate'] for row in response.data} if response.data else {}
work_rates = {**default_work_rates, **db_work_rates}

if 'work_rates' not in st.session_state:
    st.session_state['work_rates'] = work_rates

# Collapsible segment for work rates
with st.expander("Employee work rates (adjust as needed)"):
    for emp in pre_pop_employees:
        key = f"rate_{emp}"
        # Ensure the value from session state is an integer for the widget
        value_from_state = int(st.session_state['work_rates'].get(emp, 100))

        st.session_state['work_rates'][emp] = st.number_input(
            f"{emp} work rate (0 to 100%)",
            min_value=0,
            max_value=100,
            value=value_from_state,
            step=5,
            key=key
        )
    
    if st.button("Save Work Rates to Database"):
        try:
            records_to_save = [
                {"employee": emp, "rate": st.session_state['work_rates'][emp]}
                for emp in pre_pop_employees
            ]
            supabase.table("work_rates").upsert(records_to_save).execute()
            st.success("Work rates saved successfully!")
            with st.spinner("Refreshing..."):
                time.sleep(1)
            st.rerun()
        except Exception as e:
            st.error(f"An error occurred while saving to the database: {e}")

work_rates = st.session_state['work_rates']


# Button to generate schedule
if st.button("Generate Schedule"):
    mdk_days = ['Monday', 'Tuesday', 'Thursday']
    mdk_assignments = {}

    assigned_this_week = {emp: 0 for emp in available_employees}
    for day in mdk_days:
        avail_for_day = [emp for emp in available_employees if emp not in unavailable_per_day[day] and work_rates.get(emp, 0) > 0]
        if not avail_for_day:
            st.warning(f"No available employees for MDK/lunch on {day}")
            continue

        scores = {}
        for emp in avail_for_day:
            response = supabase.table("mdk_assignments").select("week", count="exact").eq("employee", emp).execute()
            history_count = response.count if response.count is not None else 0
            rate = work_rates.get(emp, 100)
            this_week_penalty = assigned_this_week[emp] * 10
            scores[emp] = (history_count / rate) + this_week_penalty if rate > 0 else float('inf')

        chosen = min(scores, key=scores.get)
        mdk_assignments[day] = chosen
        assigned_this_week[chosen] += 1

    wb = openpyxl.load_workbook('template.xlsx')
    sheet = wb['Blad1']

    sheet['A1'] = f"v.{current_week}"

    klin_cols = {'Monday': 'B', 'Tuesday': 'F', 'Wednesday': 'J', 'Thursday': 'N', 'Friday': 'R'}
    screen_cols = {'Monday': 'C', 'Tuesday': 'G', 'Wednesday': 'K', 'Thursday': 'O', 'Friday': 'S'}
    mdk_cols = {'Monday': 'D', 'Tuesday': 'H', 'Thursday': 'P'}
    lunchvakt_col = {'Wednesday': 'L'}

    lab_rows = {
        'morning1': {'LAB 3': 4, 'LAB 6': 5, 'LAB 9': 6, 'LAB 10': 7},
        'morning2': {'LAB 3': 9, 'LAB 6': 10, 'LAB 9': 11, 'LAB 10': 12},
        'afternoon1': {'LAB 3': 14, 'LAB 6': 15, 'LAB 9': 16, 'LAB 10': 17}
    }

    labs = ['LAB 3', 'LAB 6', 'LAB 9', 'LAB 10']

    for day in days:
        avail_day = [emp for emp in available_employees if emp not in unavailable_per_day[day]]
        mdk = mdk_assignments.get(day) if day in mdk_days else None
        full_day_mdk_days = ['Tuesday', 'Thursday']
        half_day_mdk_days = ['Monday']

        if day in full_day_mdk_days and mdk in avail_day:
            avail_day.remove(mdk)

        morning_lab_candidates = avail_day.copy()
        if day in mdk_days and mdk in morning_lab_candidates:
            morning_lab_candidates.remove(mdk)
        lab_people_morning = random.sample(morning_lab_candidates, min(4, len(morning_lab_candidates)))

        random.shuffle(labs)
        morning_assign = dict(zip(lab_people_morning, labs))

        morning_remainder = [emp for emp in avail_day if emp not in lab_people_morning and (emp != mdk or day not in mdk_days)]
        sheet[f"{screen_cols[day]}3"] = '/'.join(morning_remainder)

        klin_col = klin_cols[day]
        for p, l in morning_assign.items():
            sheet[f"{klin_col}{lab_rows['morning1'][l]}"] = p
            sheet[f"{klin_col}{lab_rows['morning2'][l]}"] = p

        if day != 'Friday':
            afternoon_lab_candidates = morning_remainder.copy()
            if len(afternoon_lab_candidates) < 4:
                afternoon_lab_candidates.extend([p for p in morning_lab_candidates if p not in afternoon_lab_candidates and (p != mdk or day not in mdk_days)])
            afternoon_lab_candidates = afternoon_lab_candidates[:4]
            lab_people_afternoon = random.sample(afternoon_lab_candidates, min(4, len(afternoon_lab_candidates)))

            afternoon_labs = labs.copy()
            attempts = 0
            while attempts < 100:
                random.shuffle(afternoon_labs)
                afternoon_assign = dict(zip(lab_people_afternoon, afternoon_labs))
                if all(afternoon_assign.get(p, '') != morning_assign.get(p, '') for p in lab_people_afternoon):
                    break
                attempts += 1
            else:
                st.warning(f"Could not find derangement for {day} afternoon LAB assignments. Using random.")
                afternoon_assign = dict(zip(lab_people_afternoon, afternoon_labs))

            for p, l in afternoon_assign.items():
                sheet[f"{klin_col}{lab_rows['afternoon1'][l]}"] = p

            afternoon_remainder = [emp for emp in avail_day if emp not in lab_people_afternoon]
            if day in half_day_mdk_days and mdk and mdk not in afternoon_remainder:
                afternoon_remainder.append(mdk)
            sheet[f"{screen_cols[day]}14"] = '/'.join(afternoon_remainder)

        if day in mdk_days and mdk:
            sheet[f"{mdk_cols[day]}3"] = mdk
        elif day == 'Wednesday' and morning_remainder:
            lunchvakt = random.choice(morning_remainder + lab_people_morning)
            sheet[f"{lunchvakt_col['Wednesday']}3"] = lunchvakt

    # Save new MDK assignments to Supabase
    for day, emp in mdk_assignments.items():
        supabase.table("mdk_assignments").upsert({"week": current_week, "day": day, "employee": emp}).execute()

    # Save the new schedule
    output_file = io.BytesIO()
    wb.save(output_file)
    output_file.seek(0)

    # Provide download
    st.download_button(
        "Download Generated Schedule",
        data=output_file,
        file_name=f"schedule_v{current_week}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("Schedule generated successfully!")