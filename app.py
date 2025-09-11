import io
import math
import random
import time
from collections import Counter
from datetime import date

import openpyxl
import plotly.express as px
import streamlit as st
from supabase import Client, create_client

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Schemagenerator",
    page_icon="📅",
    layout="centered",
)

# --- (Optional) CSS placeholder ---
st.markdown(
    """
    <style>
    /* Add custom CSS here if needed */
    </style>
    """,
    unsafe_allow_html=True,
)

# --- SWEDISH TRANSLATION SETUP ---
SWEDISH_DAYS = {
    "Monday": "Måndag",
    "Tuesday": "Tisdag",
    "Wednesday": "Onsdag",
    "Thursday": "Torsdag",
    "Friday": "Fredag",
}

# --- SESSION STATE FLAGS ---
if "confirm_delete" not in st.session_state:
    st.session_state.confirm_delete = False

# --- SUPABASE CLIENT ---
try:
    supabase_url = st.secrets["SUPABASE_URL"]
    supabase_key = st.secrets["SUPABASE_KEY"]
    supabase: Client = create_client(supabase_url, supabase_key)
except Exception:
    st.error("Could not connect to the database. Please check your Supabase credentials.")
    st.stop()


# --- DATA FETCHING AND CACHING ---
@st.cache_data(ttl=600)
def fetch_all_data():
    """
    Fetches MDK assignments and employee work rates.
    """
    mdk_response = supabase.table("mdk_assignments").select("employee, week, day").execute()
    work_rate_response = supabase.table("work_rates").select("employee, rate").execute()

    mdk_data = mdk_response.data if mdk_response.data else []
    work_rate_data = work_rate_response.data if work_rate_response.data else []

    return mdk_data, work_rate_data


# Cached data (refresh via st.cache_data.clear() when mutated)
all_mdk_assignments, db_work_rates_list = fetch_all_data()

# --- PRE-POPULATED DATA & CONSTANTS ---
PRE_POP_EMPLOYEES = ["AH", "LS", "DS", "KL", "TH", "LAO", "AL", "HS", "AG", "CB", "NC"]
PRE_UNAVAILABLE = {
    "Monday": ["DS", "HS", "LS"],
    "Tuesday": ["LAO", "CB", "HS", "LS"],
    "Wednesday": ["DS", "AH", "CB", "KL"],
    "Thursday": ["CB", "KL", "NC"],
    "Friday": ["CB", "AL", "KL"],
}
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

# --- Screen/MR configuration ---
# Number of Screen/MR positions per block (morning & afternoon)
SCREEN_MR_PER_BLOCK = 1
# Soft weekly cap (prefer no more than this many Screen/MR per person per week)
SCREEN_MR_WEEKLY_CAP = 1
# --- HELPERS: Weighted selection with weekly cap ---
def _unique_weighted_choices(candidates, weight_lookup, k):
    """
    Choose k unique items from candidates using their weights (probabilities ∝ work rate).
    Falls back to near-uniform if all weights are zero.
    """
    if k <= 0 or not candidates:
        return []
    picks = []
    pool = list(candidates)
    while pool and len(picks) < k:
        weights = [max(0.001, float(weight_lookup.get(c, 0))) for c in pool]
        total = sum(weights)
        if total <= 0:
            # all zero or invalid → uniform-like
            probs = None
        else:
            probs = [w / total for w in weights]
        chosen = random.choices(pool, weights=probs, k=1)[0]
        picks.append(chosen)
        pool.remove(chosen)
    return picks


def weighted_sample_with_cap(candidates, weight_lookup, k, weekly_counts, cap):
    """
    Prefer people under the weekly cap; if not enough candidates under-cap, fill from over-cap pool.
    """
    if k <= 0:
        return []
    if not candidates:
        return []
    under_cap = [c for c in candidates if weekly_counts.get(c, 0) < cap]
    if len(under_cap) >= k:
        return _unique_weighted_choices(under_cap, weight_lookup, k)
    picks = under_cap[:]  # take all under-cap
    remaining_k = k - len(picks)
    # fill the rest from all remaining (over-cap or not previously chosen)
    over_cap_pool = [c for c in candidates if c not in picks]
    if remaining_k > 0 and over_cap_pool:
        picks.extend(_unique_weighted_choices(over_cap_pool, weight_lookup, remaining_k))
    return picks


# --- UI: TITLE AND INSTRUCTIONS ---
st.title("📅 Schemagenerator för världens bästa enhet!")
st.info(
    "**💡 Användarinstruktioner:**\n\n"
    "1. **Ange frånvaro:** Markera medarbetare som är frånvarande hela veckan eller specifika dagar.\n"
    "2. **Granska & justera:** Kontrollera MDK-historiken och arbetstiden. Spara eventuella ändringar i arbetstid.\n"
    "3. **Generera:** Klicka på '✨ Generera Schema' för att skapa ett nytt veckoschema.\n"
    "4. **Ladda hem:** Klicka på '📥 Ladda ner schemat' för ta hem det."
)

# --- UI: EMPLOYEE AVAILABILITY ---
available_week = st.multiselect(
    "🙋 Initialer för samtliga medarbetare denna vecka",
    options=PRE_POP_EMPLOYEES,
    default=PRE_POP_EMPLOYEES,
)

unavailable_whole_week = st.multiselect(
    "🏖️ Initialer för medarbetare som är otillgängliga hela veckan",
    options=available_week,
    default=[],
)

# Filter out employees who are unavailable for the entire week.
available_employees = [emp for emp in available_week if emp not in unavailable_whole_week]

# Per-day unavailability
with st.expander("📅 Ange otillgänglighet per dag", expanded=True):
    unavailable_per_day = {}
    for day in DAYS:
        default_values = [emp for emp in PRE_UNAVAILABLE.get(day, []) if emp in available_employees]
        unavailable_per_day[day] = st.multiselect(
            f"Frånvarande på {SWEDISH_DAYS[day]}",
            options=available_employees,
            default=default_values,
            key=f"unavail_{day}",
        )

# --- UI: MDK DISTRIBUTION CHART (historik) ---
with st.expander("📊 MDK-fördelning (historik)"):
    if all_mdk_assignments:
        mdk_counts = Counter(assignment["employee"] for assignment in all_mdk_assignments)
        sorted_items = sorted(mdk_counts.items(), key=lambda x: x[1], reverse=True)
        employees, counts = zip(*sorted_items)
        fig = px.bar(
            x=employees,
            y=counts,
            labels={"x": "Medarbetare", "y": "Antal MDK"},
            title="MDK-fördelning (baserat på sparad historik)",
            color=counts,
            color_continuous_scale="blugrn",
        )
        fig.update_layout(xaxis={"categoryorder": "total descending"})
        fig.update_coloraxes(showscale=False)
        fig.update_yaxes(dtick=1)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Inga MDK-uppdrag finns i historiken ännu.")

# --- UI: HISTORICAL SCHEDULES (Senaste 8 veckorna) ---
current_week = date.today().isocalendar()[1]
with st.expander("📝 Historiska scheman (Senaste 8 veckorna)"):
    try:
        bucket_files = supabase.storage.from_("schedules").list()
        file_names = {f["name"] for f in bucket_files} if bucket_files else set()
    except Exception as e:
        file_names = set()
        st.warning(f"Kunde inte hämta fillistan från lagringen: {e}")

    for i in range(1, 9):
        week = current_week - i
        file_name = f"week_{week}.xlsx"
        status_emoji = "✅" if file_name in file_names else "❌"
        st.write(f"**Vecka {week}:** {status_emoji} {'Uppladdat' if file_name in file_names else 'Ej uppladdat'}")

        uploader = st.file_uploader(
            f"Ladda upp/ersätt schema för vecka {week}",
            type="xlsx",
            key=f"hist_{week}",
        )

        if uploader:
            with st.spinner(f"Laddar upp och bearbetar {file_name}..."):
                file_content = uploader.getvalue()
                supabase.storage.from_("schedules").upload(file_name, file_content, {"upsert": "true"})
                st.success(f"Laddade upp {file_name}")
                time.sleep(1)

                # Parse the uploaded file to extract MDK assignments
                downloaded = supabase.storage.from_("schedules").download(file_name)
                if downloaded:
                    wb = openpyxl.load_workbook(io.BytesIO(downloaded))
                    sheet = wb["Blad1"] if "Blad1" in wb.sheetnames else wb.active

                    # Columns used in the template for MDK
                    mdk_cols = {"Monday": "D", "Tuesday": "H", "Thursday": "P"}
                    parsed_mdk = []
                    for day, col in mdk_cols.items():
                        col_idx = openpyxl.utils.column_index_from_string(col)
                        cell_value = sheet.cell(row=3, column=col_idx).value
                        if cell_value and isinstance(cell_value, str) and cell_value.strip() in PRE_POP_EMPLOYEES:
                            parsed_mdk.append({"week": week, "day": day, "employee": cell_value.strip()})

                    if parsed_mdk:
                        try:
                            supabase.table("mdk_assignments").upsert(parsed_mdk).execute()
                            st.success(f"Läste in och uppdaterade MDK-uppdrag för vecka {week} ({len(parsed_mdk)} dagar).")
                            st.cache_data.clear()
                        except Exception as e:
                            st.error(f"Misslyckades med att spara MDK-uppdrag: {e}")
                    else:
                        st.info(f"Inga giltiga MDK-initialer hittades i filen för vecka {week}.")

# --- UI: CLEAR MDK HISTORY ---
st.markdown("---")
st.error("Radering av historik kan inte ångras.")
if st.button("🗑️ Rensa all MDK-historik"):
    st.session_state.confirm_delete = True

if st.session_state.confirm_delete:
    st.warning("**Är du helt säker på att du vill radera ALL MDK-historik?**")
    col1, col2, _ = st.columns([1.5, 1, 4])
    with col1:
        if st.button("Ja, radera all historik", type="primary"):
            try:
                supabase.table("mdk_assignments").delete().neq("week", -1).execute()
                st.success("All MDK-historik har raderats.")
                st.session_state.confirm_delete = False
                st.cache_data.clear()
                time.sleep(2)
                st.rerun()
            except Exception as e:
                st.error(f"Ett fel uppstod vid radering: {e}")
    with col2:
        if st.button("Avbryt"):
            st.session_state.confirm_delete = False
            st.rerun()

# --- UI: WORK RATES ---
db_work_rates = {row["employee"]: row["rate"] for row in db_work_rates_list}
default_work_rates = {emp: 100 for emp in PRE_POP_EMPLOYEES}
work_rates_initial = {**default_work_rates, **db_work_rates}

if "work_rates" not in st.session_state:
    st.session_state["work_rates"] = work_rates_initial.copy()

with st.expander("💼 Klinisk arbetstid per medarbetare (%)"):
    col1, col2 = st.columns(2)
    sorted_employees = sorted(PRE_POP_EMPLOYEES)
    midpoint = math.ceil(len(sorted_employees) / 2)

    for i, emp in enumerate(sorted_employees):
        target_col = col1 if i < midpoint else col2
        key = f"rate_{emp}"
        value_from_state = int(st.session_state["work_rates"].get(emp, 100))
        st.session_state["work_rates"][emp] = target_col.number_input(
            f"{emp} arbetstid",
            min_value=0,
            max_value=100,
            value=value_from_state,
            step=5,
            key=key,
            help=f"Ange den procentuella kliniska arbetstiden för {emp}.",
        )

    if st.button("💾 Spara arbetstid till databasen"):
        try:
            records_to_save = [
                {"employee": emp, "rate": st.session_state["work_rates"][emp]}
                for emp in PRE_POP_EMPLOYEES
            ]
            supabase.table("work_rates").upsert(records_to_save).execute()
            st.success("Arbetstid sparad!")
            st.cache_data.clear()
            time.sleep(1)
            st.rerun()
        except Exception as e:
            st.error(f"Ett fel uppstod vid sparning: {e}")

# Use the latest work rates from session state for generation.
work_rates = st.session_state["work_rates"]

# --- UI: GENERATE SCHEDULE ---
if st.button("✨ Generera Schema", type="primary"):
    with st.spinner("Tänker, slumpar och skapar... ett ögonblick..."):
        # --- MDK ASSIGNMENT LOGIC ---
        mdk_days = ["Monday", "Tuesday", "Thursday"]
        mdk_assignments = {}

        # Pre-calculate historical MDK counts
        mdk_history_counts = Counter(a["employee"] for a in all_mdk_assignments)
        assigned_this_week = Counter()

        for day in mdk_days:
            # Who can take MDK on this day
            avail_for_day = [
                emp
                for emp in available_employees
                if emp not in unavailable_per_day.get(day, []) and work_rates.get(emp, 0) > 0
            ]
            if not avail_for_day:
                st.warning(f"Inga tillgängliga medarbetare för MDK/lunch på {SWEDISH_DAYS[day]}")
                continue

            # Score: lower is better (less historical MDK, higher rate helps, strong penalty if already assigned this week)
            scores = {}
            for emp in avail_for_day:
                history_count = mdk_history_counts.get(emp, 0)
                rate_factor = work_rates.get(emp, 100) / 100.0
                this_week_penalty = assigned_this_week[emp] * 10
                score = (history_count / rate_factor if rate_factor > 0 else float("inf")) + this_week_penalty
                scores[emp] = score

            chosen = min(scores, key=scores.get)
            mdk_assignments[day] = chosen
            assigned_this_week[chosen] += 1

        # --- SCHEDULE POPULATION LOGIC ---
        try:
            wb = openpyxl.load_workbook("template.xlsx")
            sheet = wb["Blad1"]
        except FileNotFoundError:
            st.error("`template.xlsx` hittades inte. Se till att filen ligger i samma mapp som appen.")
            st.stop()

        sheet["A1"] = f"v.{current_week}"

        # Columns in template
        klin_cols = {"Monday": "B", "Tuesday": "F", "Wednesday": "J", "Thursday": "N", "Friday": "R"}
        screen_cols = {"Monday": "C", "Tuesday": "G", "Wednesday": "K", "Thursday": "O", "Friday": "S"}
        mdk_cols = {"Monday": "D", "Tuesday": "H", "Thursday": "P"}
        lunchvakt_col = {"Wednesday": "L"}

        lab_rows = {
            "morning1": {"LAB 3": 4, "LAB 6": 5, "LAB 9": 6, "LAB 10": 7},
            "morning2": {"LAB 3": 9, "LAB 6": 10, "LAB 9": 11, "LAB 10": 12},
            "afternoon1": {"LAB 3": 14, "LAB 6": 15, "LAB 9": 16, "LAB 10": 17},
        }
        labs = list(lab_rows["morning1"].keys())

        # Counters for fairness & weekly cap
        screen_mr_counts = Counter()       # for intra-week balancing lab vs. screen
        screen_mr_week_counts = Counter()  # for enforcing ~once/week cap on Screen/MR

        for day in DAYS:
            # Who is available today
            avail_day = [
                emp
                for emp in available_employees
                if emp not in unavailable_per_day.get(day, [])
            ]
            mdk = mdk_assignments.get(day)

            # Determine MDK behavior
            is_full_day_mdk = (day in ["Tuesday", "Thursday"]) and bool(mdk)
            is_monday_mdk_morning = (day == "Monday") and bool(mdk)

            # --- Morning Screen/MR selection (weighted + weekly cap) ---
            pool_morning_for_screen = [
                emp for emp in avail_day
                if work_rates.get(emp, 0) > 0
                and not (is_full_day_mdk and emp == mdk)               # exclude Tue/Thu MDK (full day)
                and not (is_monday_mdk_morning and emp == mdk)         # exclude Mon MDK (morning only)
            ]

            screen_mr_morning = weighted_sample_with_cap(
                candidates=pool_morning_for_screen,
                weight_lookup=work_rates,
                k=SCREEN_MR_PER_BLOCK,
                weekly_counts=screen_mr_week_counts,
                cap=SCREEN_MR_WEEKLY_CAP,
            )

            # Update counters
            for s in screen_mr_morning:
                screen_mr_week_counts[s] += 1
                screen_mr_counts[s] += 1

            # --- Morning labs: prefer those who already screened to rotate them into lab ---
            lab_candidates_morning = [emp for emp in avail_day if emp not in screen_mr_morning]

            # Respect MDK exclusions for labs
            if mdk in lab_candidates_morning and day in ["Tuesday", "Thursday"]:
                lab_candidates_morning.remove(mdk)  # full-day MDK out of labs
            if mdk in lab_candidates_morning and day == "Monday":
                lab_candidates_morning.remove(mdk)  # Mon MDK out of morning labs

            # Prioritize based on prior Screen/MR counts (more Screen/MR ⇒ more priority to lab slots)
            lab_priority_candidates = sorted(
                lab_candidates_morning,
                key=lambda emp: screen_mr_counts.get(emp, 0),
                reverse=True,
            )

            num_lab_slots = min(len(lab_priority_candidates), 4)
            lab_people_morning = lab_priority_candidates[:num_lab_slots]

            # Assign morning labs to template
            random.shuffle(labs)
            morning_assign = dict(zip(lab_people_morning, labs))

            klin_col = klin_cols[day]
            for p, l in morning_assign.items():
                sheet[f"{klin_col}{lab_rows['morning1'][l]}"] = p
                sheet[f"{klin_col}{lab_rows['morning2'][l]}"] = p

            # Write morning Screen/MR assignee(s)
            sheet[f"{screen_cols[day]}3"] = "/".join(screen_mr_morning) if screen_mr_morning else ""

            # --- Afternoon (not Friday) ---
            if day != "Friday":
                # Everyone available (except Tue/Thu MDK is full-day)
                available_for_afternoon = [
                    emp for emp in avail_day
                    if not (is_full_day_mdk and emp == mdk)
                ]
                lab_slots = min(4, len(available_for_afternoon))

                # Prefer morning Screen/MR folks for afternoon lab slots to rotate them
                preferred_candidates = [emp for emp in screen_mr_morning if emp in available_for_afternoon]
                other_candidates = [emp for emp in available_for_afternoon if emp not in preferred_candidates]
                combined_candidates = preferred_candidates + other_candidates
                lab_people_afternoon = combined_candidates[:lab_slots]

                # Try to avoid same lab morning vs afternoon for the same person (derangement attempt)
                afternoon_labs = labs[:]
                for _ in range(10):
                    random.shuffle(afternoon_labs)
                    afternoon_assign = dict(zip(lab_people_afternoon, afternoon_labs))
                    if all(
                        afternoon_assign.get(p) != morning_assign.get(p)
                        for p in afternoon_assign
                        if p in morning_assign
                    ):
                        break

                # Write afternoon labs
                for p, l in afternoon_assign.items():
                    sheet[f"{klin_col}{lab_rows['afternoon1'][l]}"] = p

                # Afternoon Screen/MR (weighted + weekly cap) from people not in afternoon labs
                afternoon_screen_pool = [
                    emp for emp in available_for_afternoon
                    if emp not in lab_people_afternoon
                    and not (is_full_day_mdk and emp == mdk)
                    and work_rates.get(emp, 0) > 0
                ]
                screen_mr_afternoon = weighted_sample_with_cap(
                    candidates=afternoon_screen_pool,
                    weight_lookup=work_rates,
                    k=SCREEN_MR_PER_BLOCK,
                    weekly_counts=screen_mr_week_counts,
                    cap=SCREEN_MR_WEEKLY_CAP,
                )

                for s in screen_mr_afternoon:
                    screen_mr_week_counts[s] += 1
                    screen_mr_counts[s] += 1

                sheet[f"{screen_cols[day]}14"] = "/".join(screen_mr_afternoon) if screen_mr_afternoon else ""

            # --- MDK & Lunch Guard (Wednesday only) ---
            if mdk:
                if day in mdk_cols:
                    sheet[f"{mdk_cols[day]}3"] = mdk
            elif day == "Wednesday" and avail_day:
                lunch_candidates = [p for p in avail_day if p not in lab_people_morning] or avail_day
                sheet[f"{lunchvakt_col['Wednesday']}3"] = random.choice(lunch_candidates)

        # --- SAVE & DOWNLOAD ---
        new_mdk_records = [{"week": current_week, "day": d, "employee": e} for d, e in mdk_assignments.items()]
        if new_mdk_records:
            supabase.table("mdk_assignments").upsert(new_mdk_records).execute()
            st.cache_data.clear()

        output_file = io.BytesIO()
        wb.save(output_file)
        output_file.seek(0)

        st.success("✅ Schemat har genererats!")
        st.download_button(
            label="📥 Ladda ner schemat (.xlsx)",
            data=output_file,
            file_name=f"veckoschema_v{current_week}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
