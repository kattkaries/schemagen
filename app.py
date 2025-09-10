import streamlit as st
import openpyxl
import random
from datetime import date
import io
from supabase import create_client, Client
import plotly.express as px
import time
from collections import Counter
import math

# --- PAGE CONFIGURATION ---
# Set page to centered layout for better readability on wide screens.
st.set_page_config(
    page_title="Schemagenerator",
    page_icon="📅",
    layout="centered"
)

# --- CSS FOR STYLING THE MULTISELECT WIDGET ---
# This CSS targets the selection "tags" in the first multiselect widget on the page.
st.markdown("""
<style>
    /* Target the container for the first multiselect widget */
    div[data-testid="stMultiSelect"]:first-of-type {
        /* This selector is for scoping purposes; no specific style is needed here. */
    }

    /* Style the selected "tags" within the first multiselect widget */
    div[data-testid="stMultiSelect"]:first-of-type [data-baseweb="tag"] {
        background-color: #2E8B57; /* SeaGreen */
        border-radius: 0.5rem;   /* Rounded corners */
    }

    /* Improve contrast by making the text and 'x' button white */
    div[data-testid="stMultiSelect"]:first-of-type [data-baseweb="tag"] span,
    div[data-testid="stMultiSelect"]:first-of-type [data-baseweb="tag"] span[role="button"] {
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# --- SWEDISH TRANSLATION SETUP ---
# Dictionary to map English day names to Swedish for the UI.
SWEDISH_DAYS = {
    'Monday': 'Måndag', 'Tuesday': 'Tisdag', 'Wednesday': 'Onsdag',
    'Thursday': 'Torsdag', 'Friday': 'Fredag'
}

# --- SESSION STATE INITIALIZATION ---
# Initialize a flag for the multi-step deletion confirmation process.
if 'confirm_delete' not in st.session_state:
    st.session_state.confirm_delete = False

# --- SUPABASE CLIENT INITIALIZATION ---
# Securely connect to the Supabase backend using secrets.
try:
    supabase_url = st.secrets["SUPABASE_URL"]
    supabase_key = st.secrets["SUPABASE_KEY"]
    supabase: Client = create_client(supabase_url, supabase_key)
    st.success("✅ Databasanslutning etablerad.")
except Exception as e:
    st.error(f"❌ Kunde inte ansluta till databasen: {e}. Appen körs i begränsat läge.")
    supabase = None

# --- DATA FETCHING AND CACHING ---
@st.cache_data(ttl=600)  # Cache for 10 minutes to reduce database calls
def fetch_all_data():
    """
    Fetches all necessary data from the database in a single call.
    Returns MDK assignments and employee work rates.
    """
    if supabase:
        try:
            mdk_response = supabase.table("mdk_assignments").select("employee, week, day").execute()
            work_rate_response = supabase.table("work_rates").select("employee, rate").execute()
            return (
                mdk_response.data if mdk_response.data else [],
                work_rate_response.data if work_rate_response.data else []
            )
        except Exception as e:
            st.warning(f"⚠️ Kunde inte hämta data: {e}. Använder standardvärden.")
            return [], []
    return [], []

# Fetch data once per session run.
all_mdk_assignments, db_work_rates_list = fetch_all_data()

# --- PRE-POPULATED DATA & CONSTANTS ---
PRE_POP_EMPLOYEES = ['AH', 'LS', 'DS', 'KL', 'TH', 'LAO', 'AL', 'HS', 'AG', 'CB', 'NC']
PRE_UNAVAILABLE = {
    'Monday': ['DS', 'HS', 'LS'],
    'Tuesday': ['LAO', 'CB', 'HS', 'LS'],
    'Wednesday': ['DS', 'AH', 'CB', 'KL'],
    'Thursday': ['CB', 'KL', 'NC'],
    'Friday': ['CB', 'AL', 'KL']
}
DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

# --- UI: TITLE AND INSTRUCTIONS ---
st.title("📅 Schemagenerator för världens bästa enhet!")

st.info(
    "**💡 Användarinstruktioner:**\n\n"
    "1. **Ange frånvaro:** Markera medarbetare som är frånvarande hela veckan eller specifika dagar.\n"
    "2. **Granska & justera:** Kontrollera MDK-historiken och arbetstiden. Spara ändringar i arbetstid.\n"
    "3. **Generera:** Klicka på '✨ Generera Schema' för att skapa ett schema. Högre arbetstid ökar chans för MDK och Screen/MR.\n"
    "4. **Ladda hem:** Klicka på '📥 Ladda ner schemat' för att spara Excel-filen."
)

# --- UI: EMPLOYEE AVAILABILITY SETUP ---
available_week = st.multiselect(
    "🙋 Initialer för samtliga medarbetare denna vecka",
    options=PRE_POP_EMPLOYEES,
    default=PRE_POP_EMPLOYEES,
    help="Välj alla medarbetare som är schemaläggningsbara denna vecka."
)

unavailable_whole_week = st.multiselect(
    "🏖️ Initialer för medarbetare som är otillgängliga hela veckan",
    options=available_week,
    default=[],
    help="Medarbetare här listas inte för någon dag."
)

# Filter out employees who are unavailable for the entire week.
available_employees = [emp for emp in available_week if emp not in unavailable_whole_week]

# UI for setting unavailability per day.
with st.expander("📅 Ange otillgänglighet per dag", expanded=True):
    unavailable_per_day = {}
    for day in DAYS:
        default_values = [emp for emp in PRE_UNAVAILABLE.get(day, []) if emp in available_employees]
        unavailable_per_day[day] = st.multiselect(
            f"Frånvarande på {SWEDISH_DAYS[day]}",
            options=available_employees,
            default=default_values,
            help=f"Välj medarbetare som inte kan schemaläggas på {SWEDISH_DAYS[day]}."
        )

# --- UI: MDK DISTRIBUTION CHART ---
with st.expander("📊 MDK-fördelning (historik)"):
    if all_mdk_assignments:
        # Calculate MDK counts from cached data.
        mdk_counts = Counter(assignment['employee'] for assignment in all_mdk_assignments)
        sorted_items = sorted(mdk_counts.items(), key=lambda x: x[1], reverse=True)
        employees, counts = zip(*sorted_items) if mdk_counts else ([], [])
        
        if mdk_counts:
            fig = px.bar(
                x=employees,
                y=counts,
                labels={'x': 'Medarbetare', 'y': 'Antal MDK'},
                title="MDK-fördelning (baserat på sparad historik)",
                color=counts,
                color_continuous_scale="blugrn",
            )
            fig.update_layout(xaxis={'categoryorder':'total descending'})
            fig.update_coloraxes(showscale=False)
            fig.update_yaxes(dtick=1)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("ℹ️ Inga MDK-uppdrag finns i historiken ännu.")
    else:
        st.info("ℹ️ Inga MDK-uppdrag finns i historiken ännu.")

# --- UI: HISTORICAL SCHEDULES ---
current_week = date.today().isocalendar()[1]
with st.expander("🗄️ Historiska scheman (Senaste 8 veckorna)"):
    file_names = set()
    if supabase:
        try:
            bucket_files = supabase.storage.from_("schedules").list()
            file_names = {f['name'] for f in bucket_files} if bucket_files else set()
        except Exception as e:
            st.warning(f"⚠️ Kunde inte hämta fillistan: {e}")

    progress_bar = st.progress(0)
    for i in range(1, 9):
        progress_bar.progress(i / 8)
        week = current_week - i
        file_name = f"week_{week}.xlsx"
        status_emoji = "✅" if file_name in file_names else "❌"
        st.write(f"**Vecka {week}:** {status_emoji} {'Uppladdat' if file_name in file_names else 'Ej uppladdat'}")
        
        uploader = st.file_uploader(
            f"Ladda upp/ersätt schema för vecka {week}", 
            type="xlsx", 
            key=f"hist_{week}",
            help="Ladda upp XLSX-fil för att uppdatera historik och MDK-data."
        )
        if uploader:
            with st.spinner(f"Laddar upp och bearbetar {file_name}..."):
                try:
                    file_content = uploader.getvalue()
                    supabase.storage.from_("schedules").upload(file_name, file_content, {"upsert": "true"})
                    st.success(f"✅ Laddade upp {file_name}")
                    time.sleep(0.5)

                    downloaded = supabase.storage.from_("schedules").download(file_name)
                    if downloaded:
                        wb = openpyxl.load_workbook(io.BytesIO(downloaded))
                        sheet = wb["Blad1"] if "Blad1" in wb.sheetnames else wb.active
                        mdk_cols = {'Monday': 'D', 'Tuesday': 'H', 'Thursday': 'P'}
                        parsed_mdk = []
                        for day, col in mdk_cols.items():
                            cell_value = sheet.cell(row=3, column=openpyxl.utils.column_index_from_string(col)).value
                            if cell_value and isinstance(cell_value, str) and cell_value.strip() in PRE_POP_EMPLOYEES:
                                parsed_mdk.append({"week": week, "day": day, "employee": cell_value.strip()})
                        
                        if parsed_mdk:
                            supabase.table("mdk_assignments").upsert(parsed_mdk).execute()
                            st.success(f"✅ Läste in MDK-uppdrag för vecka {week} ({len(parsed_mdk)} dagar).")
                            st.cache_data.clear()
                        else:
                            st.info(f"ℹ️ Inga giltiga MDK-initialer hittades för vecka {week}.")
                    else:
                        st.error(f"❌ Kunde inte ladda ner {file_name}.")
                except Exception as e:
                    st.error(f"❌ Fel vid hantering av {file_name}: {e}")

    progress_bar.empty()

    # --- UI: CLEAR MDK HISTORY ---
    st.markdown("---")
    st.error("⚠️ Radering av historik kan inte ångras.")
    if st.button("🗑️ Rensa all MDK-historik"):
        st.session_state.confirm_delete = True

    if st.session_state.confirm_delete:
        st.warning("**Är du helt säker på att du vill radera ALL MDK-historik?**")
        col1, col2, _ = st.columns([1.5, 1, 4])
        with col1:
            if st.button("Ja, radera all historik", type="primary"):
                if supabase:
                    try:
                        supabase.table("mdk_assignments").delete().neq("week", -1).execute()
                        st.success("✅ All MDK-historik har raderats.")
                        st.session_state.confirm_delete = False
                        st.cache_data.clear()
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Fel vid radering: {e}")
                else:
                    st.error("❌ Ingen databasanslutning.")
        with col2:
            if st.button("Avbryt"):
                st.session_state.confirm_delete = False
                st.rerun()

# --- UI: WORK RATES ---
db_work_rates = {row['employee']: row['rate'] for row in db_work_rates_list}
default_work_rates = {emp: 100 for emp in PRE_POP_EMPLOYEES}
work_rates = {**default_work_rates, **db_work_rates}

if 'work_rates' not in st.session_state:
    st.session_state['work_rates'] = work_rates.copy()

with st.expander("💼 Klinisk arbetstid per medarbetare (%)"):
    col1, col2 = st.columns(2)
    sorted_employees = sorted(PRE_POP_EMPLOYEES)
    midpoint = math.ceil(len(sorted_employees) / 2)

    for i, emp in enumerate(sorted_employees):
        target_col = col1 if i < midpoint else col2
        key = f"rate_{emp}"
        value_from_state = int(st.session_state['work_rates'].get(emp, 100))
        
        st.session_state['work_rates'][emp] = target_col.number_input(
            f"{emp} arbetstid",
            min_value=0, max_value=100,
            value=value_from_state,
            step=5, key=key,
            help=f"Ange klinisk arbetstid för {emp}. Högre värde ökar chans för MDK och Screen/MR."
        )
    
    if st.button("💾 Spara arbetstid till databasen"):
        if supabase:
            try:
                records_to_save = [
                    {"employee": emp, "rate": st.session_state['work_rates'][emp]}
                    for emp in PRE_POP_EMPLOYEES
                ]
                supabase.table("work_rates").upsert(records_to_save).execute()
                st.success("✅ Arbetstid sparad!")
                st.cache_data.clear()
                time.sleep(0.5)
                st.rerun()
            except Exception as e:
                st.error(f"❌ Fel vid sparning: {e}")
        else:
            st.warning("⚠️ Ingen databasanslutning. Förändringar sparas lokalt.")

work_rates = st.session_state['work_rates']

# --- UI: GENERATE SCHEDULE BUTTON ---
if st.button("✨ Generera Schema", type="primary"):
    if not available_employees:
        st.error("❌ Inga tillgängliga medarbetare. Välj minst en.")
    else:
        with st.spinner("🔄 Tänker, slumpar och skapar... ett ögonblick..."):
            try:
                # --- MDK ASSIGNMENT LOGIC ---
                mdk_days = ['Monday', 'Tuesday', 'Thursday']
                mdk_assignments = {}
                mdk_history_counts = Counter(a['employee'] for a in all_mdk_assignments)
                assigned_this_week = Counter()

                for day in mdk_days:
                    avail_for_day = [
                        emp for emp in available_employees
                        if emp not in unavailable_per_day.get(day, []) and work_rates.get(emp, 0) > 0
                    ]
                    if not avail_for_day:
                        st.warning(f"⚠️ Inga tillgängliga medarbetare för MDK på {SWEDISH_DAYS[day]}")
                        continue

                    scores = {}
                    for emp in avail_for_day:
                        history_count = mdk_history_counts.get(emp, 0)
                        rate_factor = work_rates.get(emp, 100) / 100.0
                        this_week_penalty = assigned_this_week[emp] * 10
                        scores[emp] = (history_count / rate_factor if rate_factor > 0 else float('inf')) + this_week_penalty

                    chosen = min(scores, key=scores.get)
                    mdk_assignments[day] = chosen
                    assigned_this_week[chosen] += 1

                # --- SCHEDULE POPULATION LOGIC ---
                try:
                    wb = openpyxl.load_workbook('template.xlsx')
                    sheet = wb['Blad1']
                except FileNotFoundError:
                    st.error("❌ `template.xlsx` hittades inte. Se till att filen finns.")
                    st.stop()
                except Exception as e:
                    st.error(f"❌ Fel vid laddning av template: {e}")
                    st.stop()

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
                labs = list(lab_rows['morning1'].keys())

                # Track Screen/MR assignments for fairness
                screen_mr_counts = Counter()

                for day in DAYS:
                    avail_day = [emp for emp in available_employees if emp not in unavailable_per_day.get(day, [])]
                    mdk = mdk_assignments.get(day)
                    
                    # Lab candidates (prioritize those with fewer Screen/MR shifts)
                    lab_candidates = [emp for emp in avail_day if not (mdk == emp and day in ['Tuesday', 'Thursday'])]
                    morning_candidates = [emp for emp in lab_candidates if not (mdk == emp and day == 'Monday')]

                    # Morning lab assignment (unchanged: prioritize fewer Screen/MR shifts)
                    num_lab_slots = min(len(morning_candidates), 4)
                    lab_people_morning = sorted(
                        morning_candidates,
                        key=lambda emp: screen_mr_counts.get(emp, 0)
                    )[:num_lab_slots]
                    random.shuffle(labs)
                    morning_assign = dict(zip(lab_people_morning, labs))

                    # Morning Screen/MR: Weight by work rates
                    morning_remainder = [emp for emp in avail_day if emp not in lab_people_morning and (emp != mdk or day not in mdk_days)]
                    if morning_remainder:
                        weights = [100 / (work_rates.get(emp, 100) / 100.0 or 0.1) + screen_mr_counts.get(emp, 0) * 5 for emp in morning_remainder]
                        morning_screen_mr = random.choices(morning_remainder, weights=[1/w if w > 0 else 0.001 for w in weights], k=len(morning_remainder))
                    else:
                        morning_screen_mr = []
                    sheet[f"{screen_cols[day]}3"] = '/'.join(sorted(morning_screen_mr))

                    klin_col = klin_cols[day]
                    for p, l in morning_assign.items():
                        sheet[f"{klin_col}{lab_rows['morning1'][l]}"] = p
                        sheet[f"{klin_col}{lab_rows['morning2'][l]}"] = p

                    # Afternoon assignment (not Friday)
                    afternoon_screen_mr_pool = []
                    if day != 'Friday':
                        available_for_afternoon = lab_candidates[:]
                        lab_slots = min(4, len(available_for_afternoon))
                        
                        # Afternoon lab: Prefer morning Screen/MR employees
                        preferred_candidates = [emp for emp in morning_screen_mr if emp in available_for_afternoon]
                        other_candidates = [emp for emp in available_for_afternoon if emp not in preferred_candidates]
                        combined_candidates = preferred_candidates + other_candidates
                        lab_people_afternoon = combined_candidates[:lab_slots]
                        
                        afternoon_labs = labs[:]
                        for _ in range(10):
                            random.shuffle(afternoon_labs)
                            afternoon_assign = dict(zip(lab_people_afternoon, afternoon_labs))
                            if all(afternoon_assign.get(p) != morning_assign.get(p) for p in afternoon_assign if p in morning_assign):
                                break
                        
                        for p, l in afternoon_assign.items():
                            sheet[f"{klin_col}{lab_rows['afternoon1'][l]}"] = p

                        # Afternoon Screen/MR: Weight by work rates
                        afternoon_screen_mr_pool = [emp for emp in available_for_afternoon if emp not in lab_people_afternoon]
                        if afternoon_screen_mr_pool:
                            weights = [100 / (work_rates.get(emp, 100) / 100.0 or 0.1) + screen_mr_counts.get(emp, 0) * 5 for emp in afternoon_screen_mr_pool]
                            afternoon_screen_mr_pool = random.choices(afternoon_screen_mr_pool, weights=[1/w if w > 0 else 0.001 for w in weights], k=len(afternoon_screen_mr_pool))
                        sheet[f"{screen_cols[day]}14"] = '/'.join(sorted(afternoon_screen_mr_pool))

                    # MDK and Lunch Guard
                    if mdk and day in mdk_cols:
                        sheet[f"{mdk_cols[day]}3"] = mdk
                    elif day == 'Wednesday' and avail_day:
                        lunch_candidates = [p for p in avail_day if p not in lab_people_morning] or avail_day
                        sheet[f"{lunchvakt_col['Wednesday']}3"] = random.choice(lunch_candidates)

                    # Update Screen/MR counts
                    all_screeners_today = set(morning_screen_mr) | set(afternoon_screen_mr_pool)
                    for emp in all_screeners_today:
                        screen_mr_counts[emp] += 1

                # --- SAVE & DOWNLOAD ---
                if supabase and mdk_assignments:
                    try:
                        new_mdk_records = [
                            {"week": current_week, "day": day, "employee": emp}
                            for day, emp in mdk_assignments.items()
                        ]
                        supabase.table("mdk_assignments").upsert(new_mdk_records).execute()
                        st.success("✅ MDK-uppdrag sparade.")
                        st.cache_data.clear()
                    except Exception as e:
                        st.warning(f"⚠️ Kunde inte spara MDK-uppdrag: {e}. Schemat genererades ändå.")

                output_file = io.BytesIO()
                wb.save(output_file)
                output_file.seek(0)
                
                st.success("🎉 Schemat har genererats!")
                st.download_button(
                    label="📥 Ladda ner schemat (.xlsx)",
                    data=output_file,
                    file_name=f"veckoschema_v{current_week}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Klicka för att ladda ner den genererade Excel-filen."
                )
            except Exception as e:
                st.error(f"❌ Fel vid generering: {e}. Kontrollera indata och försök igen.")