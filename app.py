import streamlit as st
import openpyxl
import random
import json
import os
from datetime import date

# Load history if exists
history_file = 'mdk_history.json'
if os.path.exists(history_file):
    with open(history_file, 'r') as f:
        mdk_history = json.load(f)
else:
    mdk_history = {}

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

# Inputs for unavailable per day (multiselect instead of drag-drop for simplicity)
st.write("Note: For drag-and-drop functionality, consider integrating a custom component like 'streamlit-draggable-list' or 'streamlit-elements'.")
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
    # MDK days
    mdk_days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday']
    mdk_assignments = {}

    # Assign MDK with priorities
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
            history_count = mdk_history.get(emp, 0)
            rate = work_rates.get(emp, 1.0)
            this_week_penalty = assigned_this_week[emp] * 10  # Heavy penalty to avoid multiple per week
            scores[emp] = (history_count / rate) + this_week_penalty if rate > 0 else float('inf')

        chosen = min(scores, key=scores.get)
        mdk_assignments[day] = chosen
        assigned_this_week[chosen] += 1

    # Now generate the schedule for each day
    wb = openpyxl.load_workbook('template.xlsx')
    sheet = wb['Blad1']

    # Update week number
    current_week = date.today().isocalendar()[1]
    sheet['A1'] = f"v.{current_week}"

    # Column mappings
    klin_cols = {'Monday': 'B', 'Tuesday': 'F', 'Wednesday': 'J', 'Thursday': 'N', 'Friday': 'R'}
    screen_cols = {'Monday': 'C', 'Tuesday': 'G', 'Wednesday': 'K', 'Thursday': 'O', 'Friday': 'S'}
    mdk_cols = {'Monday': 'D', 'Tuesday': 'H', 'Wednesday': 'L', 'Thursday': 'P'}  # L for lunchvakt on Wednesday

    # LAB row mappings per time slot
    lab_rows = {
        'morning1': {'LAB 3': 4, 'LAB 6': 5, 'LAB 9': 6, 'LAB 10': 7},
        'morning2': {'LAB 3': 9, 'LAB 6': 10, 'LAB 9': 11, 'LAB 10': 12},
        'afternoon1': {'LAB 3': 14, 'LAB 6': 15, 'LAB 9': 16, 'LAB 10': 17},
        'afternoon2': {'LAB 3': 19, 'LAB 6': 20, 'LAB 9': 21, 'LAB 10': 22}
    }

    labs = ['LAB 3', 'LAB 6', 'LAB 9', 'LAB 10']

    for day in days:
        avail_day = [emp for emp in available_employees if emp not in unavailable_per_day[day]]

        mdk = mdk_assignments.get(day)
        full_day_mdk_days = ['Tuesday', 'Thursday']
        half_day_mdk_days = ['Monday', 'Wednesday']

        # Remove MDK if full day
        if day in full_day_mdk_days and mdk in avail_day:
            avail_day.remove(mdk)

        # Choose 4 people for LAB (prefer not to include half-day MDK)
        lab_candidates = [emp for emp in avail_day if not (day in half_day_mdk_days and emp == mdk)]
        if len(lab_candidates) < 4:
            lab_candidates = avail_day  # Include if necessary

        lab_people = random.sample(lab_candidates, min(4, len(lab_candidates)))

        # Morning assignment
        random.shuffle(labs)
        morning_assign = dict(zip(lab_people, labs))

        # Afternoon assignment (derangement)
        afternoon_labs = labs.copy()
        attempts = 0
        while attempts < 100:
            random.shuffle(afternoon_labs)
            afternoon_assign = dict(zip(lab_people, afternoon_labs))
            if all(afternoon_assign.get(p, '') != morning_assign.get(p, '') for p in lab_people):
                break
            attempts += 1
        else:
            st.warning(f"Could not find derangement for {day} LAB assignments. Using random (may have overlaps).")
            afternoon_assign = dict(zip(lab_people, afternoon_labs))

        # Fill LAB cells
        klin_col = klin_cols[day]
        for p, l in morning_assign.items():
            row = lab_rows['morning1'][l]
            sheet[f"{klin_col}{row}"] = p
            row = lab_rows['morning2'][l]
            sheet[f"{klin_col}{row}"] = p

        for p, l in afternoon_assign.items():
            row = lab_rows['afternoon1'][l]
            sheet[f"{klin_col}{row}"] = p
            row = lab_rows['afternoon2'][l]
            sheet[f"{klin_col}{row}"] = p

        # Remainder for Screen/MR
        remainder = [emp for emp in avail_day if emp not in lab_people]
        if day in half_day_mdk_days and mdk and mdk not in remainder:
            remainder.append(mdk)  # Add half-day MDK to Screen/MR

        screen_str = '/'.join(remainder)
        screen_cell = f"{screen_cols[day]}3"
        sheet[screen_cell] = screen_str

        # Fill MDK/lunch
        if day in mdk_days and mdk:
            mdk_cell = f"{mdk_cols[day]}3"
            sheet[mdk_cell] = mdk

    # Save the new schedule
    output_file = 'generated_schedule.xlsx'
    wb.save(output_file)

    # Update history
    for emp in mdk_assignments.values():
        mdk_history[emp] = mdk_history.get(emp, 0) + 1
    with open(history_file, 'w') as f:
        json.dump(mdk_history, f)

    # Provide download
    with open(output_file, 'rb') as f:
        st.download_button(
            "Download Generated Schedule",
            data=f,
            file_name=f"schedule_v{current_week}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.success("Schedule generated successfully!")