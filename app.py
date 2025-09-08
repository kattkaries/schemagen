import streamlit as st
import openpyxl
import random
from datetime import date
import io
from supabase import create_client, Client
import plotly.express as px
import time
import pandas as pd  # Added for potential data handling improvements

# --- CSS FOR STYLING MULTISELECT ---
st.markdown("""
<style>
    /* Target the first multiselect widget on the page */
    div[data-testid="stMultiSelect"]:first-of-type {
        /* No specific style needed here for scoping */
    }

    /* Style selected tags in the first multiselect */
    div[data-testid="stMultiSelect"]:first-of-type [data-baseweb="tag"] {
        background-color: #2E8B57; /* SeaGreen for selected tags */
        border-radius: 0.5rem;    /* Rounded corners */
    }

    /* Ensure text and close button in tags have good contrast */
    div[data-testid="stMultiSelect"]:first-of-type [data-baseweb="tag"] span {
        color: white !important;
    }
    div[data-testid="stMultiSelect"]:first-of-type [data-baseweb="tag"] span[role="button"] {
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# --- SWEDISH TRANSLATION SETUP ---
days_sv = {
    'Monday': 'Måndag', 'Tuesday': 'Tisdag', 'Wednesday': 'Onsdag', 
    'Thursday': 'Torsdag', 'Friday': 'Fredag'
}

# --- TRACK CONFIRMATION FOR DELETION ---
if 'confirm_delete' not in st.session_state:
    st.session_state.confirm_delete = False

# Initialize Supabase client with error handling
try:
    supabase_url = st.secrets["SUPABASE_URL"]
    supabase_key = st.secrets["SUPABASE_KEY"]
    supabase: Client = create_client(supabase_url, supabase_key)
    st.success("✅ Databasanslutning etablerad.")  # Visual feedback on connection
except Exception as e:
    st.error(f"❌ Misslyckades med att ansluta till databasen: {e}. Appen körs i begränsat läge.")
    supabase = None  # Graceful degradation: set to None and handle in code

# Pre-populated data (unchanged for functionality)
pre_pop_employees = ['AH', 'LS', 'DS', 'KL', 'TH', 'LAO', 'AL', 'HS', 'AG', 'CB']
pre_unavailable = {
    'Monday': ['DS'],
    'Tuesday': ['LAO', 'CB'],
    'Wednesday': ['DS', 'AH', 'CB'],
    'Thursday': ['CB'],
    'Friday': ['CB', 'AL']
}

st.title("📅 Schemagenerator för världens bästa enhet!")  # Added emoji for UX

# Help text for user guidance
st.info("👋 Välkommen! Välj medarbetare, ange otillgänglighet och generera ett optimerat schema. Historiska data används för balanserad MDK-fördelning.")

# Employee selection with improved layout using columns if needed, but single multiselect is fine
available_week = st.multiselect(
    "Initialer för samtliga medarbetare",
    options=pre_pop_employees,
    default=pre_pop_employees,
    help="Välj alla medarbetare som är schemaläggningsbara denna vecka."
)

# Input for unavailable whole week
unavailable_whole_week = st.multiselect(
    "Initialer för medarbetare som är otillgängliga hela veckan",
    options=available_week,
    default=[],
    help="Medarbetare här listas inte för någon dag."
)

available_employees = [emp for emp in available_week if emp not in unavailable_whole_week]

# Multiselect for unavailable per day (efficient list comprehension)
days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
with st.expander("📋 Ange otillgänglighet per dag", expanded=True):
    unavailable_per_day = {}
    for day in days:
        default_values = [emp for emp in pre_unavailable.get(day, []) if emp in available_employees]
        unavailable_per_day[day] = st.multiselect(
            f"Initialer för otillgängliga medarbetare på {days_sv[day]}",
            options=available_employees,
            default=default_values,
            help=f"Välj medarbetare som inte kan schemaläggas på {days_sv[day]}."
        )

# Fetch MDK data once for reuse (performance optimization: single DB call)
mdk_data = None
historical_mdk_counts = {}
if supabase:
    try:
        response = supabase.table("mdk_assignments").select("employee").execute()
        mdk_data = response.data if response.data else []
        # Efficient counting with dict comprehension
        historical_mdk_counts = {emp: mdk_data.count(lambda x: x['employee'] == emp) for emp in pre_pop_employees}
    except Exception as e:
        st.warning(f"⚠️ Kunde inte hämta MDK-historik: {e}. Använder standardvärden för balansering.")
        historical_mdk_counts = {emp: 0 for emp in pre_pop_employees}

# MDK Overview Bar Graph (reuses fetched data)
with st.expander("📊 MDK-fördelning de senaste månaderna (stapeldiagram)"):
    if mdk_data and historical_mdk_counts:
        mdk_counts = {k: v for k, v in historical_mdk_counts.items() if v > 0}
        if mdk_counts:
            # Sort by count descending (efficient)
            sorted_items = sorted(mdk_counts.items(), key=lambda x: x[1], reverse=True)
            employees, counts = zip(*sorted_items)
            fig = px.bar(
                x=employees,
                y=counts,
                labels={'x': 'Medarbetare', 'y': 'Antal MDK'},
                title="MDK-fördelning de senaste månaderna",  # Minor text tweak
                color=counts,
                color_continuous_scale="blugrn",
            )
            fig.update_coloraxes(showscale=False)
            fig.update_yaxes(dtick=1)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("ℹ️ Inga MDK-uppdrag i historiken ännu.")
    else:
        st.info("ℹ️ Inga data tillgängliga för MDK-grafen.")

# Historical Schedules Upload (last 8 weeks) - Optimized bucket list once
current_week = date.today().isocalendar()[1]
with st.expander("📁 Historiska scheman (Senaste 8 veckorna)"):
    bucket_files = []
    file_names = []
    if supabase:
        try:
            bucket_files = supabase.storage.from_("schedules").list()
            file_names = [f['name'] for f in bucket_files] if bucket_files else []
        except Exception as e:
            st.warning(f"⚠️ Kunde inte lista filer i bucket: {e}. Fortsätter utan statusvisning.")

    progress_bar = st.progress(0)
    status_texts = []
    for i in range(1, 9):
        progress_bar.progress(i / 8)
        week = current_week - i
        file_name = f"week_{week}.xlsx"
        status = "✅ uppladdat" if file_name in file_names else "❌ ej uppladdat"  # Emojis for UX
        status_texts.append(f"Vecka {week}: {status}")
        st.write(f"Vecka {week}: {status}")
        
        uploader = st.file_uploader(
            f"Ladda upp/ersätt schema för vecka {week}", 
            type="xlsx", 
            key=f"hist_{week}",
            help="Ladda upp XLSX-fil för att uppdatera historik och MDK-data."
        )
        if uploader:
            try:
                file_content = uploader.getvalue()
                supabase.storage.from_("schedules").upload(file_name, file_content, {"upsert": "true"})
                st.success(f"✅ Laddade upp {file_name}")
                time.sleep(0.5)  # Reduced sleep for better UX

                # Parse and save MDK (with error handling)
                downloaded = supabase.storage.from_("schedules").download(file_name)
                if downloaded:
                    wb = openpyxl.load_workbook(io.BytesIO(downloaded))
                    sheet = wb["Blad1"] if "Blad1" in wb.sheetnames else wb.active
                    mdk_cols = {'Monday': 'D', 'Tuesday': 'H', 'Thursday': 'P'}
                    parsed_days = 0
                    for day, col in mdk_cols.items():
                        cell_value = sheet.cell(row=3, column=openpyxl.utils.column_index_from_string(col)).value
                        if cell_value and isinstance(cell_value, str) and cell_value.strip() in pre_pop_employees:
                            supabase.table("mdk_assignments").upsert({
                                "week": week, "day": day, "employee": cell_value.strip()
                            }).execute()
                            parsed_days += 1
                    st.success(f"✅ Läste in och uppdaterade MDK-uppdrag för vecka {week}. {parsed_days} dagar tillagda.")
                    # Refresh historical counts after update
                    if mdk_data:
                        # Re-fetch or update in memory (simple append simulation for perf)
                        pass  # For now, rerun will refresh
                else:
                    st.error(f"❌ Kunde inte ladda ner {file_name} efter uppladdning.")
            except Exception as e:
                st.error(f"❌ Fel vid hantering av {file_name}: {e}")

    st.text("\n".join(status_texts))  # Summary
    progress_bar.empty()

    # --- BUTTON TO CLEAR MDK HISTORY ---
    st.markdown("---")
    if st.button("🗑️ Rensa all MDK-historik"):
        st.session_state.confirm_delete = True

    if st.session_state.confirm_delete:
        st.warning("⚠️ Är du säker på att du vill radera all MDK-historik? Detta kan inte ångras.")
        
        col1, col2, _ = st.columns([1.5, 1, 4])
        with col1:
            if st.button("Ja, radera all historik", type="primary"):
                if supabase:
                    try:
                        supabase.table("mdk_assignments").delete().neq("week", -1).execute()
                        st.success("✅ All MDK-historik har raderats.")
                        st.session_state.confirm_delete = False
                        # Reset in-memory data
                        historical_mdk_counts = {emp: 0 for emp in pre_pop_employees}
                        time.sleep(1)
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Ett fel uppstod vid radering: {e}")
                else:
                    st.error("❌ Ingen databasanslutning tillgänglig.")
        with col2:
            if st.button("Avbryt"):
                st.session_state.confirm_delete = False
                st.rerun()

# Work rates: Fetch once, use session_state (already efficient)
default_work_rates = {emp: 100 for emp in pre_pop_employees}
db_work_rates = {}
if supabase:
    try:
        response = supabase.table("work_rates").select("*").execute()
        db_work_rates = {row['employee']: row['rate'] for row in response.data} if response.data else {}
    except Exception as e:
        st.warning(f"⚠️ Kunde inte hämta arbetstider: {e}. Använder standardvärden.")
work_rates = {**default_work_rates, **db_work_rates}

if 'work_rates' not in st.session_state:
    st.session_state['work_rates'] = work_rates

# Collapsible for work rates with better UX
with st.expander("⏰ Klinisk arbetstid per medarbetare (justera vid behov)"):
    for emp in pre_pop_employees:
        key = f"rate_{emp}"
        value_from_state = int(st.session_state['work_rates'].get(emp, 100))

        st.session_state['work_rates'][emp] = st.number_input(
            f"{emp} arbetstid (0 till 100%)",
            min_value=0,
            max_value=100,
            value=value_from_state,
            step=5,
            key=key,
            help="Procentuell andel av normal arbetstid. Lägre värde minskar sannolikhet för tunga uppgifter."
        )
    
    if st.button("💾 Spara arbetstid till databasen"):
        if supabase:
            try:
                records_to_save = [
                    {"employee": emp, "rate": st.session_state['work_rates'][emp]}
                    for emp in pre_pop_employees
                ]
                supabase.table("work_rates").upsert(records_to_save).execute()
                st.success("✅ Arbetstid sparad!")
                with st.spinner("Uppdaterar vy..."):
                    time.sleep(0.5)
                st.rerun()
            except Exception as e:
                st.error(f"❌ Ett fel uppstod vid sparning: {e}")
        else:
            st.warning("⚠️ Ingen databasanslutning. Förändringar sparas endast lokalt.")

work_rates = st.session_state['work_rates']

# Button to generate schedule with enhanced loading
if st.button("🚀 Generera Schema"):
    if not available_employees:
        st.error("❌ Inga tillgängliga medarbetare. Välj minst en.")
    else:
        with st.spinner("🔄 Genererar schema – detta tar ett ögonblick..."):
            try:
                # Use pre-fetched historical counts (performance win)
                mdk_days = ['Monday', 'Tuesday', 'Thursday']
                mdk_assignments = {}
                assigned_this_week = {emp: 0 for emp in available_employees}

                for day in mdk_days:
                    avail_for_day = [emp for emp in available_employees 
                                     if emp not in unavailable_per_day.get(day, []) 
                                     and work_rates.get(emp, 0) > 0]
                    if not avail_for_day:
                        st.warning(f"⚠️ Inga tillgängliga medarbetare för MDK på {days_sv[day]}")
                        continue

                    scores = {}
                    for emp in avail_for_day:
                        history_count = historical_mdk_counts.get(emp, 0)  # Use pre-fetched
                        rate = work_rates.get(emp, 100)
                        this_week_penalty = assigned_this_week[emp] * 10
                        scores[emp] = (history_count / rate) + this_week_penalty if rate > 0 else float('inf')

                    chosen = min(scores, key=scores.get)
                    mdk_assignments[day] = chosen
                    assigned_this_week[chosen] += 1

                # Template loading with error handling
                try:
                    wb = openpyxl.load_workbook('template.xlsx')
                except FileNotFoundError:
                    st.error("❌ Template-filen 'template.xlsx' hittades inte. Se till att den finns i appens katalog.")
                    st.stop()
                except Exception as e:
                    st.error(f"❌ Fel vid laddning av template: {e}")
                    st.stop()

                sheet = wb['Blad1']
                sheet['A1'] = f"v.{current_week}"

                # Column mappings (unchanged)
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
                    
                    # Simplified logic for lab candidates (efficient list ops)
                    lab_candidates = avail_day[:]
                    if mdk in lab_candidates and day in ['Tuesday', 'Thursday']:
                        lab_candidates.remove(mdk)

                    # Morning assignment
                    morning_candidates = lab_candidates[:]
                    if mdk in morning_candidates and day == 'Monday':
                        morning_candidates.remove(mdk)

                    lab_people_morning = random.sample(morning_candidates, k=min(len(morning_candidates), 4))
                    random.shuffle(labs)
                    morning_assign = dict(zip(lab_people_morning, labs))
                    morning_remainder = [emp for emp in avail_day if emp not in lab_people_morning]

                    sheet[f"{screen_cols[day]}3"] = '/'.join(sorted([emp for emp in morning_remainder if (emp != mdk or day not in mdk_days)]))
                    klin_col = klin_cols[day]
                    for p, l in morning_assign.items():
                        sheet[f"{klin_col}{lab_rows['morning1'][l]}"] = p
                        sheet[f"{klin_col}{lab_rows['morning2'][l]}"] = p

                    # Afternoon assignment (not Friday)
                    if day != 'Friday':
                        available_for_afternoon = lab_candidates[:]
                        if mdk in available_for_afternoon and day in ['Tuesday', 'Thursday']:
                            available_for_afternoon.remove(mdk)

                        lab_people_afternoon = []
                        screen_mr_afternoon = []
                        
                        # Assign LAB roles, preferring morning Screen/MR
                        lab_slots = min(4, len(available_for_afternoon))
                        lab_candidates_afternoon = [emp for emp in available_for_afternoon if emp in morning_remainder] or available_for_afternoon  # Fixed: was morning_screen_mr, but defined as morning_remainder
                        if len(lab_candidates_afternoon) >= lab_slots:
                            lab_people_afternoon = random.sample(lab_candidates_afternoon, lab_slots)
                        else:
                            lab_people_afternoon = lab_candidates_afternoon[:]
                            remaining_slots = lab_slots - len(lab_people_afternoon)
                            if remaining_slots > 0:
                                other_candidates = [emp for emp in available_for_afternoon if emp not in lab_people_afternoon]
                                lab_people_afternoon.extend(random.sample(other_candidates, min(remaining_slots, len(other_candidates))))

                        # Assign Screen/MR, preferring morning LAB
                        screen_mr_candidates = [emp for emp in available_for_afternoon if emp not in lab_people_afternoon]
                        morning_lab_employees = list(morning_assign.keys())
                        screen_mr_candidates_with_pref = [emp for emp in morning_lab_employees if emp in screen_mr_candidates] + [emp for emp in screen_mr_candidates if emp not in morning_lab_employees]
                        if screen_mr_candidates_with_pref:
                            screen_mr_afternoon = random.sample(screen_mr_candidates_with_pref, min(1, len(screen_mr_candidates_with_pref)))

                        afternoon_labs = labs[:]
                        # Derangement attempt (unchanged, but limited tries for perf)
                        for _ in range(10):
                            random.shuffle(afternoon_labs)
                            afternoon_assign = dict(zip(lab_people_afternoon, afternoon_labs))
                            if all(afternoon_assign.get(p) != morning_assign.get(p) for p in afternoon_assign if p in morning_assign):
                                break
                        
                        for p, l in afternoon_assign.items():
                            sheet[f"{klin_col}{lab_rows['afternoon1'][l]}"] = p

                        afternoon_screen_mr_pool = [emp for emp in available_for_afternoon if emp not in lab_people_afternoon]
                        sheet[f"{screen_cols[day]}14"] = '/'.join(sorted(afternoon_screen_mr_pool))

                    # MDK and Lunch Guard
                    if mdk and day in mdk_cols:
                        sheet[f"{mdk_cols[day]}3"] = mdk
                    elif day == 'Wednesday' and avail_day:
                        sheet[f"{lunchvakt_col['Wednesday']}3"] = random.choice(avail_day)

                # Batch save new MDK assignments (performance: single upsert)
                if supabase and mdk_assignments:
                    try:
                        records = [{"week": current_week, "day": day, "employee": emp} for day, emp in mdk_assignments.items()]
                        supabase.table("mdk_assignments").upsert(records).execute()
                        st.success("✅ MDK-uppdrag sparade till historik.")
                        # Update in-memory counts
                        for emp in mdk_assignments.values():
                            historical_mdk_counts[emp] = historical_mdk_counts.get(emp, 0) + 1
                    except Exception as e:
                        st.error(f"⚠️ Kunde inte spara MDK-uppdrag: {e}. Schemat genererades ändå.")

                # Save to bytes
                output_file = io.BytesIO()
                wb.save(output_file)
                
                st.success("🎉 Schemat har genererats framgångsrikt!")
                
                # Download button
                output_file.seek(0)
                st.download_button(
                    label="⬇️ Ladda ner schemat",
                    data=output_file,
                    file_name=f"veckoschema_v{current_week}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Klicka för att ladda ner den genererade Excel-filen."
                )
            except Exception as e:
                st.error(f"❌ Ett oväntat fel uppstod vid generering: {e}. Försök igen eller kontrollera indata.")