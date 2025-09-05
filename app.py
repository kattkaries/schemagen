import streamlit as st
import openpyxl
import random
import json
import os
from datetime import date

# Load MDK history if exists (may reset on Streamlit Cloud restarts)
history_file = 'mdk_history.json'
if os.path.exists(history_file):
    with open(history_file, 'r') as f:
        mdk_history = json.load(f)
else:
    mdk_history = {}

# Initialize session state for temporary history if not already set
if 'mdk_history' not in st.session_state:
    st.session_state['mdk_history'] = mdk_history

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

# Inputs for unavailable per day
st.write("Note: For drag-and-drop functionality, consider integrating a custom component like 'streamlit-draggable-list'.")
unavailable_per_day = {}
days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
for day in days:
    unavailable_per_day[day] = st.multiselect(
        f"Initials of employees unavailable on {day}",
        options=available_employees,
        default=pre_unavailable.get(day, [])
    )

# Collapsible segment for work rates
with st.expander("Employee work rates (adjust as needed)"):
    if 'work_rates' not in st.session_state:
        st.session_state['work_rates'] = {emp: 1.0 for emp in pre_pop_employees}
    for emp in pre_pop_employees:
        st.session_state['work_rates'][emp] = st.number_input(
            f"{emp} work rate (0.0 to 1.0)",
            min_value=0.0,
            max_value=1.0,
            value=st.session_state['work_rates'][emp],
            step=0.05
        )

work_rates = st.session_state['work_rates']

# Button to generate schedule
if st.button("Generate Schedule"):
    # MDK days (excluding Wednesday)
    mdk_days = ['Monday', 'Tuesday', 'Thursday']
    mdk_assignments = {}

    # Assign MDK with priorities for MDK days
    assigned_this_week = {emp: 0 for emp in available_employees}
    for day in mdk_days:
        avail_for_day = [
            emp for emp in available_employees
            if emp not in unavailable_per_day[day] and work_rates.get(emp, 0) > 0
        ]
        if not avail_for_day:
            st.warning(f"No available employees for MDK/lunch on {day}")
            continue

        scores = {}
        for emp in avail_for_day:
            history_count = st.session_state['mdk_history'].get(emp, 0)
            rate = work_rates.get(emp, 1.0)
            this_week_penalty = assigned_this_week[emp] * 10
            scores[emp] = (history_count / rate) + this_week_penalty if rate > 0 else float('inf')

        chosen = min(scores, key=scores.get)
        mdk_assignments[day] = chosen
        assigned_this_week[chosen] += 1
        st.session_state['mdk_history'][chosen] = st.session_state['mdk_history'].get(chosen, 0) + 1

    # Save MDK history to file (may not persist on Streamlit Cloud)
    try:
        with open(history_file, 'w') as f:
            json.dump(st.session_state['mdk_history'], f)
    except Exception as e:
        st.warning(f"Could not save MDK history to file: {e}. Using session state for this session.")

    # Generate the schedule
    wb = openpyxl.load_workbook('template.xlsx')
    sheet = wb['Blad1']

    # Update week number
    current_week = date.today().isocalendar()[1]
    sheet['A1'] = f"v.{current_week}"

    # Column mappings
    klin_cols = {'Monday': 'B', 'Tuesday': 'F', 'Wednesday': 'J', 'Thursday': 'N', 'Friday': 'R'}
    screen_cols = {'Monday': 'C', 'Tuesday': 'G', 'Wednesday': 'K', 'Thursday': 'O', 'Friday': 'S'}
    mdk_cols = {'Monday': 'D', 'Tuesday': 'H', 'Thursday': 'P'}  # Removed Wednesday from MDK
    lunchvakt_col = {'Wednesday': 'L'}  # Specific column for Wednesday lunchvakt

    # LAB row mappings per time slot (only morning and afternoon1 for Monday-Thursday, only morning for Friday)
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

        # Remove MDK if full day
        if day in full_day_mdk_days and mdk in avail_day:
            avail_day.remove(mdk)

        # Choose up to 4 people for morning LAB, preferring those not in morning Screen/MR
        morning_lab_candidates = avail_day.copy()
        if day in mdk_days and mdk in morning_lab_candidates:
            morning_lab_candidates.remove(mdk)  # Exclude MDK from morning LAB
        lab_people_morning = random.sample(morning_lab_candidates, min(4, len(morning_lab_candidates)))

        # Morning LAB assignment
        random.shuffle(labs)
        morning_assign = dict(zip(lab_people_morning, labs))

        # Morning Screen/MR (row 3), exclude MDK
        morning_remainder = [emp for emp in avail_day if emp not in lab_people_morning and (emp != mdk or day not in mdk_days)]
        screen_str_morning = '/'.join(morning_remainder)
        screen_cell_morning = f"{screen_cols[day]}3"
        sheet[screen_cell_morning] = screen_str_morning

        # Fill morning LAB cells
        klin_col = klin_cols[day]
        for p, l in morning_assign.items():
            row = lab_rows['morning1'][l]
            sheet[f"{klin_col}{row}"] = p
            row = lab_rows['morning2'][l]
            sheet[f"{klin_col}{row}"] = p

        # Afternoon assignments (skip for Friday afternoon)
        if day != 'Friday':
            # Prefer morning Screen/MR people for afternoon LAB
            afternoon_lab_candidates = morning_remainder.copy()
            if len(afternoon_lab_candidates) < 4:
                afternoon_lab_candidates.extend([p for p in morning_lab_candidates if p not in afternoon_lab_candidates and (p != mdk or day not in mdk_days)])
            afternoon_lab_candidates = afternoon_lab_candidates[:4]  # Limit to 4
            lab_people_afternoon = random.sample(afternoon_lab_candidates, min(4, len(afternoon_lab_candidates)))

            # Afternoon LAB assignment (derangement from morning)
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

            # Fill afternoon LAB cells
            for p, l in afternoon_assign.items():
                row = lab_rows['afternoon1'][l]
                sheet[f"{klin_col}{row}"] = p

            # Afternoon Screen/MR (row 14), include MDK for half-day MDK days or any for Wednesday
            afternoon_remainder = [emp for emp in avail_day if emp not in lab_people_afternoon]
            if (day in half_day_mdk_days and mdk and mdk not in afternoon_remainder) or day == 'Wednesday':
                afternoon_remainder.append(mdk) if day in half_day_mdk_days and mdk else None
            screen_str_afternoon = '/'.join(afternoon_remainder)
            screen_cell_afternoon = f"{screen_cols[day]}14"
            sheet[screen_cell_afternoon] = screen_str_afternoon

        # Fill MDK/lunch or lunchvakt
        if day in mdk_days and mdk:
            mdk_cell = f"{mdk_cols[day]}3"
            sheet[mdk_cell] = mdk
        elif day == 'Wednesday' and morning_remainder:  # Assign lunchvakt from morning assignments
            lunchvakt = random.choice(morning_remainder + lab_people_morning)  # Any from LAB or Screen/MR
            lunchvakt_cell = f"{lunchvakt_col['Wednesday']}3"
            sheet[lunchvakt_cell] = lunchvakt

    # Save the new schedule
    output_file = 'generated_schedule.xlsx'
    wb.save(output_file)

    # Provide download
    with open(output_file, 'rb') as f:
        st.download_button(
            "Download Generated Schedule",
            data=f,
            file_name=f"schedule_v{current_week}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.success("Schedule generated successfully!")