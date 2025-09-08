import streamlit as st
import openpyxl
import random
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo
import io
from supabase import create_client, Client
import plotly.express as px

# --- CSS for multiselect styling (scoped with ID, avoids brittle selectors) ---
st.markdown("""
<style>
#ms-emp [data-baseweb="tag"]{
    background:#2E8B57;border-radius:.5rem;
}
#ms-emp [data-baseweb="tag"] span,
#ms-emp [data-baseweb="tag"] span[role="button"]{
    color:white !important;
}
</style>
""", unsafe_allow_html=True)

# --- Constants ---
DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
DAYS_SV = {
    'Monday': 'Måndag', 'Tuesday': 'Tisdag', 'Wednesday': 'Onsdag',
    'Thursday': 'Torsdag', 'Friday': 'Fredag'
}

# --- Session state init ---
if 'confirm_delete' not in st.session_state:
    st.session_state.confirm_delete = False

# --- Supabase client (cached) ---
@st.cache_resource
def get_supabase() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

supabase = get_supabase()

# --- Helpers ---
def last_n_iso_weeks(n: int, tz="Europe/Stockholm"):
    """Return list of (year, week) tuples for last n full ISO weeks."""
    now = datetime.now(ZoneInfo(tz))
    monday = now - timedelta(days=now.weekday())  # start of current week
    weeks = []
    for i in range(1, n + 1):
        d = monday - timedelta(weeks=i)
        y, w, _ = d.isocalendar()
        weeks.append((y, w))
    return weeks

def slash_join(names):
    """Join list of employee initials with '/', safe for empty lists."""
    return "/".join(sorted(names)) if names else ""

# --- Cached queries ---
@st.cache_data(ttl=120)
def list_schedule_objects():
    return supabase.storage.from_("schedules").list() or []

@st.cache_data(ttl=120)
def fetch_all_mdk():
    return supabase.table("mdk_assignments").select("employee").execute().data or []

# --- Pre-populated data ---
pre_pop_employees = sorted(['AH', 'LS', 'DS', 'KL', 'TH', 'LAO', 'AL', 'HS', 'AG', 'CB'])
pre_unavailable = {
    'Monday': ['DS'],
    'Tuesday': ['LAO', 'CB'],
    'Wednesday': ['DS', 'AH', 'CB'],
    'Thursday': ['CB'],
    'Friday': ['CB', 'AL']
}

# --- Title ---
st.title("Schemagenerator för världens bästa enhet!")

# --- Employee selection ---
st.markdown('<div id="ms-emp">', unsafe_allow_html=True)
available_week = st.multiselect(
    "Initialer för samtliga medarbetare",
    options=pre_pop_employees,
    default=pre_pop_employees,
    key="ms_emp"
)
st.markdown('</div>', unsafe_allow_html=True)

# Input for unavailable whole week
unavailable_whole_week = st.multiselect(
    "Initialer för medarbetare som är otillgängliga hela veckan",
    options=available_week,
    default=[]
)
available_employees = [emp for emp in available_week if emp not in unavailable_whole_week]

if not available_employees:
    st.error("Alla medarbetare är markerade som otillgängliga denna vecka.")
    st.stop()

# --- Daily unavailable ---
with st.expander("Ange otillgänglighet per dag", expanded=True):
    unavailable_per_day = {}
    for day in DAYS:
        default_values = [emp for emp in pre_unavailable.get(day, []) if emp in available_employees]
        unavailable_per_day[day] = st.multiselect(
            f"Initialer för otillgängliga medarbetare på {DAYS_SV[day]}",
            options=available_employees,
            default=default_values
        )

# --- MDK Overview ---
with st.expander("MDK-fördelning de senaste månaderna (stapeldiagram)"):
    assignments = fetch_all_mdk()
    mdk_counts = {}
    for a in assignments:
        emp = a['employee']
        mdk_counts[emp] = mdk_counts.get(emp, 0) + 1
    mdk_counts = {k: v for k, v in mdk_counts.items() if v > 0}

    if mdk_counts:
        # Sort employees by MDK count (descending)
        sorted_items = sorted(mdk_counts.items(), key=lambda x: x[1], reverse=True)
        employees, counts = zip(*sorted_items)

        fig = px.bar(
            x=employees,
            y=counts,
            labels={'x': 'Medarbetare', 'y': 'Antal MDK'},
            title="MDK-fördelning de senaste 2 månaderna",
            color=counts,
            color_continuous_scale=["green", "yellow", "red"],  # low=green, high=red
        )
        fig.update_coloraxes(showscale=False)  # hide legend if you want a clean look
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Inga MDK-uppdrag i historiken ännu.")


# --- Historical schedules ---
iso_year, current_week, _ = date.today().isocalendar()
with st.expander("Historiska scheman (Senaste 8 veckorna)"):
    bucket_files = list_schedule_objects()
    file_names = [f['name'] for f in bucket_files]

    for year, week in last_n_iso_weeks(8):
        file_name = f"week_{year}_{week}.xlsx"
        st.write(f"Vecka {week} ({year})")
        status = "uppladdat" if file_name in file_names else "ej uppladdat"
        st.write(f"Status för fil: {status}")

        uploader = st.file_uploader(
            f"Ladda upp/ersätt schema för vecka {week} ({year})",
            type="xlsx",
            key=f"hist_{year}_{week}"
        )
        if uploader:
            file_content = uploader.getvalue()
            with st.status("Laddar upp...", expanded=True) as s:
                supabase.storage.from_("schedules").upload(file_name, file_content, {"upsert": "true"})
                s.update(label=f"{file_name} uppladdad", state="complete")

            downloaded = supabase.storage.from_("schedules").download(file_name)
            if downloaded:
                wb = openpyxl.load_workbook(io.BytesIO(downloaded))
                sheet = wb["Blad1"] if "Blad1" in wb.sheetnames else wb.active
                mdk_cols = {'Monday': 'D', 'Tuesday': 'H', 'Thursday': 'P'}
                parsed_days = 0
                for day, col in mdk_cols.items():
                    cell_value = sheet.cell(
                        row=3,
                        column=openpyxl.utils.column_index_from_string(col)
                    ).value
                    if cell_value and isinstance(cell_value, str) and cell_value.strip() in pre_pop_employees:
                        try:
                            supabase.table("mdk_assignments").upsert({
                                "iso_year": year, "week": week,
                                "day": day, "employee": cell_value.strip()
                            }).execute()
                            parsed_days += 1
                        except Exception as e:
                            st.error(f"Misslyckades för {DAYS_SV[day]}: {e}")
                st.success(f"Uppdaterade MDK-uppdrag för vecka {week}. {parsed_days} dagar tillagda.")

    # --- Clear history ---
    st.markdown("---")
    if st.button("Rensa all MDK-historik"):
        st.session_state.confirm_delete = True

    if st.session_state.confirm_delete:
        st.warning("Är du säker på att du vill radera all MDK-historik? Detta kan inte ångras.")
        col1, col2, _ = st.columns([1.5, 1, 4])
        with col1:
            if st.button("Ja, radera all historik", type="primary"):
                try:
                    supabase.table("mdk_assignments").delete().neq("week", -1).execute()
                    st.success("All MDK-historik har raderats.")
                    st.session_state.confirm_delete = False
                    st.rerun()
                except Exception as e:
                    st.error(f"Ett fel uppstod vid radering: {e}")
        with col2:
            if st.button("Avbryt"):
                st.session_state.confirm_delete = False
                st.rerun()

# --- Work rates ---
default_work_rates = {emp: 100 for emp in pre_pop_employees}
response = supabase.table("work_rates").select("*").execute()
db_work_rates = {row['employee']: row['rate'] for row in response.data} if response.data else {}
work_rates = {**default_work_rates, **db_work_rates}

if 'work_rates' not in st.session_state:
    st.session_state['work_rates'] = work_rates

with st.expander("Klinisk arbetstid per medarbetare (justera vid behov)"):
    for emp in pre_pop_employees:
        key = f"rate_{emp}"
        value_from_state = int(st.session_state['work_rates'].get(emp, 100))
        st.session_state['work_rates'][emp] = st.number_input(
            f"{emp} arbetstid (0 till 100%)",
            min_value=0, max_value=100,
            value=value_from_state, step=5, key=key
        )
    
    if st.button("Spara arbetstid till databasen"):
        try:
            records = [{"employee": emp, "rate": st.session_state['work_rates'][emp]} for emp in pre_pop_employees]
            supabase.table("work_rates").upsert(records).execute()
            st.success("Arbetstid sparad!")
            st.rerun()
        except Exception as e:
            st.error(f"Ett fel uppstod: {e}")

work_rates = st.session_state['work_rates']

# --- Generate schedule ---
if st.button("Generera Schema"):
    st.subheader("Genererat Schema")

    # Fetch MDK history once (aggregated)
    rows = supabase.table("mdk_assignments").select("employee").execute().data or []
    mdk_history = {}
    for r in rows:
        e = r["employee"]
        mdk_history[e] = mdk_history.get(e, 0) + 1

    # Setup workbook
    wb = openpyxl.load_workbook("template.xlsx")
    sheet = wb.active

    assignments = {}
    mdk_cols = {"Monday": "D", "Tuesday": "H", "Thursday": "P"}

    for day in DAYS:
        available_today = [
            emp for emp in available_employees
            if emp not in unavailable_per_day[day]
        ]
        if not available_today:
            st.error(f"Inga tillgängliga medarbetare för {DAYS_SV[day]}")
            continue

        # --- MDK assignment ---
        if day in mdk_cols:
            best_emp = None
            best_score = float("inf")

            for emp in available_today:
                history_count = mdk_history.get(emp, 0)
                rate = work_rates.get(emp, 100)
                score = history_count / rate

                # penalty if missed earlier days
                penalty = sum(
                    1 for prev_day in DAYS
                    if prev_day in assignments and emp not in assignments[prev_day]
                )
                score += penalty

                if score < best_score:
                    best_emp = emp
                    best_score = score

            if best_emp:
                assignments.setdefault(day, []).append(best_emp)
                sheet[f"{mdk_cols[day]}3"] = best_emp
                mdk_history[best_emp] = mdk_history.get(best_emp, 0) + 1

                try:
                    supabase.table("mdk_assignments").upsert({
                        "iso_year": iso_year, "week": current_week,
                        "day": day, "employee": best_emp
                    }).execute()
                except Exception as e:
                    st.error(f"Kunde inte spara MDK-uppdrag för {best_emp} på {DAYS_SV[day]}: {e}")

        # --- Random assignment for other tasks ---
        others = [emp for emp in available_today if emp not in assignments.get(day, [])]
        random.shuffle(others)
        assignments.setdefault(day, []).extend(others)

        # Fill Excel row
        row_num = {"Monday": 3, "Tuesday": 6, "Wednesday": 9, "Thursday": 12, "Friday": 15}[day]
        for col, emp in enumerate(assignments[day], start=2):
            sheet.cell(row=row_num, column=col, value=emp)

    # Save Excel to buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    # Upload to Supabase storage
    file_name = f"week_{iso_year}_{current_week}.xlsx"
    supabase.storage.from_("schedules").upload(file_name, buffer.getvalue(), {"upsert": "true"})

    st.success("Schema genererat och uppladdat!")
    st.download_button("Ladda ner schemat som Excel-fil", buffer, file_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
