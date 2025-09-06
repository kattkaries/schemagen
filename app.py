import streamlit as st
import openpyxl
import random
from datetime import date
import io
from supabase import create_client, Client
import plotly.express as px
import time

# --- SWEDISH TRANSLATION SETUP ---
days_sv = {
    'Monday': 'Måndag', 'Tuesday': 'Tisdag', 'Wednesday': 'Onsdag', 
    'Thursday': 'Torsdag', 'Friday': 'Fredag'
}

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
    'Friday': ['CB', 'AL']
}

st.title("Schemagenerator för världens bästa enhet!")

# Input for available employees this week
available_week = st.multiselect(
    "Initialer för samtliga medarbetare",
    options=pre_pop_employees,
    default=pre_pop_employees
)

# Input for unavailable whole week
unavailable_whole_week = st.multiselect(
    "Initialer för medarbetare som är otillgängliga hela veckan",
    options=available_week,
    default=[]
)

available_employees = [emp for emp in available_week if emp not in unavailable_whole_week]

# Multiselect for unavailable per day
days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
with st.expander("Ange otillgänglighet per dag", expanded=True):
    unavailable_per_day = {}
    for day in days:
        default_values = [emp for emp in pre_unavailable.get(day, []) if emp in available_employees]
        unavailable_per_day[day] = st.multiselect(
            f"Initialer för otillgängliga medarbetare på {days_sv[day]}",
            options=available_employees,
            default=default_values
        )

# MDK Overview Bar Graph
with st.expander("MDK-fördelning de senaste månaderna (stapeldiagram)"):
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
        fig = px.bar(x=employees, y=counts, 
                     labels={'x': 'Medarbetare', 'y': 'Antal MDK'}, 
                     title="MDK-fördelning de senaste 2 månaderna")
        st.plotly_chart(fig)
    else:
        st.info("Inga MDK-uppdrag i historiken ännu.")

# Historical Schedules Upload (last 8 weeks)
current_week = date.today().isocalendar()[1]
with st.expander("Historiska scheman (Senaste 8 veckorna)"):
    bucket_files = supabase.storage.from_("schedules").list()
    file_names = [f['name'] for f in bucket_files] if bucket_files else []

    for i in range(1, 9):
        week = current_week - i
        file_name = f"week_{week}.xlsx"
        st.write(f"Vecka {week}")
        status = "uppladdat" if file_name in file_names else "ej uppladdat"
        st.write(f"Status för fil: {status}")
        uploader = st.file_uploader(f"Ladda upp/ersätt schema för vecka {week}", type="xlsx", key=f"hist_{week}")
        if uploader:
            file_content = uploader.getvalue()
            supabase.storage.from_("schedules").upload(file_name, file_content, {"upsert": "true"})
            st.success(f"Laddade upp {file_name}")
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
                            st.error(f"Misslyckades med att spara MDK-uppdrag för {days_sv[day]}: {e}")
                st.success(f"Läste in och uppdaterade MDK-uppdrag för vecka {week}. {parsed_days} dagar tillagda.")

# Use percentages (0-100) for work rates
default_work_rates = {emp: 100 for emp in pre_pop_employees}
response = supabase.table("work_rates").select("*").execute()
db_work_rates = {row['employee']: row['rate'] for row in response.data} if response.data else {}
work_rates = {**default_work_rates, **db_work_rates}

if 'work_rates' not in st.session_state:
    st.session_state['work_rates'] = work_rates

# Collapsible segment for work rates
with st.expander("Klinisk arbetstid per medarbetare (justera vid behov)"):
    for emp in pre_pop_employees:
        key = f"rate_{emp}"
        value_from_state = int(st.session_state['work_rates'].get(emp, 100))

        st.session_state['work_rates'][emp] = st.number_input(
            f"{emp} arbetstid (0 till 100%)",
            min_value=0,
            max_value=100,
            value=value_from_state,
            step=5,
            key=key
        )
    
    if st.button("Spara arbetstid till databasen"):
        try:
            records_to_save = [
                {"employee": emp, "rate": st.session_state['work_rates'][emp]}
                for emp in pre_pop_employees
            ]
            supabase.table("work_rates").upsert(records_to_save).execute()
            st.success("Arbetstid sparad!")
            with st.spinner("Uppdaterar..."):
                time.sleep(1)
            st.rerun()
        except Exception as e:
            st.error(f"Ett fel uppstod vid sparning till databasen: {e}")

work_rates = st.session_state['work_rates']

# Button to generate schedule
if st.button("Generera Schema"):
    with st.spinner("Genererar schema, vänligen vänta..."):
        mdk_days = ['Monday', 'Tuesday', 'Thursday']
        mdk_assignments = {}

        assigned_this_week = {emp: 0 for emp in available_employees}
        for day in mdk_days:
            avail_for_day = [emp for emp in available_employees if emp not in unavailable_per_day.get(day, []) and work_rates.get(emp, 0) > 0]
            if not avail_for_day:
                st.warning(f"Inga tillgängliga medarbetare för MDK/lunch på {days_sv[day]}")
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
            avail_day = [emp for emp in available_employees if emp not in unavailable_per_day.get(day, [])]
            mdk = mdk_assignments.get(day)
            
            # Simplified logic for who is available for labs
            lab_candidates = avail_day[:]
            if mdk in lab_candidates and day in ['Tuesday', 'Thursday']:  # Full-day MDK
                lab_candidates.remove(mdk)

            # Morning assignment
            morning_candidates = lab_candidates[:]
            if mdk in morning_candidates and day in ['Monday']:  # Half-day MDK
                morning_candidates.remove(mdk)

            lab_people_morning = random.sample(morning_candidates, k=min(len(morning_candidates), 4))
            random.shuffle(labs)
            morning_assign = dict(zip(lab_people_morning, labs))
            morning_remainder = [emp for emp in avail_day if emp not in lab_people_morning]
            # Track morning Screen/MR assignments
            morning_screen_mr = morning_remainder[:]

            sheet[f"{screen_cols[day]}3"] = '/'.join(sorted([emp for emp in morning_remainder if (emp != mdk or day not in mdk_days)]))
            klin_col = klin_cols[day]
            for p, l in morning_assign.items():
                sheet[f"{klin_col}{lab_rows['morning1'][l]}"] = p
                sheet[f"{klin_col}{lab_rows['morning2'][l]}"] = p

            # Afternoon assignment (not Friday)
            if day != 'Friday':
                afternoon_candidates = lab_candidates[:]
                # Remove MDK if present for full-day MDK days
                if mdk in afternoon_candidates and day in ['Tuesday', 'Thursday']:
                    afternoon_candidates.remove(mdk)

                # Prioritize variety: LAB for morning Screen/MR, Screen/MR for morning LAB
                lab_people_afternoon = []
                screen_mr_afternoon = []
                available_for_afternoon = afternoon_candidates[:]

                # Assign LAB roles first, preferring those on Screen/MR in the morning
                lab_slots = min(4, len(available_for_afternoon))
                lab_candidates_afternoon = [emp for emp in available_for_afternoon if emp in morning_screen_mr] or available_for_afternoon
                if len(lab_candidates_afternoon) >= lab_slots:
                    lab_people_afternoon = random.sample(lab_candidates_afternoon, lab_slots)
                else:
                    lab_people_afternoon = lab_candidates_afternoon[:]
                    remaining_slots = lab_slots - len(lab_people_afternoon)
                    if remaining_slots > 0:
                        other_candidates = [emp for emp in available_for_afternoon if emp not in lab_people_afternoon]
                        lab_people_afternoon.extend(random.sample(other_candidates, min(remaining_slots, len(other_candidates))))

                # Assign Screen/MR roles, preferring those on LAB in the morning, excluding LAB assignees
                screen_mr_candidates = [emp for emp in available_for_afternoon if emp not in lab_people_afternoon]
                if morning_assign:
                    morning_lab_employees = [emp for emp in morning_assign.keys() if emp in screen_mr_candidates]
                    if morning_lab_employees:
                        screen_mr_afternoon.extend(random.sample(morning_lab_employees, min(1, len(morning_lab_employees))))
                        screen_mr_candidates = [emp for emp in screen_mr_candidates if emp not in screen_mr_afternoon]
                if len(screen_mr_afternoon) < 1 and screen_mr_candidates:
                    screen_mr_afternoon.append(random.choice(screen_mr_candidates))

                afternoon_labs = labs[:]
                # Simple derangement attempt
                for _ in range(10):  # Try to shuffle to avoid same lab
                    random.shuffle(afternoon_labs)
                    afternoon_assign = dict(zip(lab_people_afternoon, afternoon_labs))
                    if all(afternoon_assign.get(p) != morning_assign.get(p) for p in afternoon_assign if p in morning_assign):
                        break
                
                for p, l in afternoon_assign.items():
                    sheet[f"{klin_col}{lab_rows['afternoon1'][l]}"] = p

                # Recalculate afternoon_remainder excluding all assigned employees
                afternoon_remainder = [emp for emp in avail_day if emp not in lab_people_afternoon and emp not in screen_mr_afternoon]
                if day in ['Monday'] and mdk and mdk not in afternoon_remainder:
                    afternoon_remainder.append(mdk)
                sheet[f"{screen_cols[day]}14"] = '/'.join(sorted(afternoon_remainder))

            # MDK and Lunch Guard assignment
            if mdk:
                sheet[f"{mdk_cols[day]}3"] = mdk
            elif day == 'Wednesday' and avail_day:
                sheet[f"{lunchvakt_col['Wednesday']}3"] = random.choice(avail_day)

        # Save new MDK assignments to Supabase
        for day, emp in mdk_assignments.items():
            supabase.table("mdk_assignments").upsert({"week": current_week, "day": day, "employee": emp}).execute()

        # Save the new schedule to a byte stream
        output_file = io.BytesIO()
        wb.save(output_file)
        
        st.success("Schemat har genererats!")
        # --- MANUAL DOWNLOAD ---
        output_file.seek(0)
        st.download_button(
            label="Ladda ner schemat",
            data=output_file,
            file_name=f"schema_v{current_week}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )