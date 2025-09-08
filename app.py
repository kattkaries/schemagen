import streamlit as st
import openpyxl
import random
from datetime import date
import io
from supabase import create_client, Client
import plotly.express as px
import time
from typing import Dict, List, Optional

# --- CONFIGURATION ---
class Config:
    DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
    DAYS_SV = {
        'Monday': 'Måndag', 'Tuesday': 'Tisdag', 'Wednesday': 'Onsdag', 
        'Thursday': 'Torsdag', 'Friday': 'Fredag'
    }
    MDK_DAYS = ['Monday', 'Tuesday', 'Thursday']
    LABS = ['LAB 3', 'LAB 6', 'LAB 9', 'LAB 10']
    PRE_POP_EMPLOYEES = ['AH', 'LS', 'DS', 'KL', 'TH', 'LAO', 'AL', 'HS', 'AG', 'CB']
    PRE_UNAVAILABLE = {
        'Monday': ['DS'],
        'Tuesday': ['LAO', 'CB'],
        'Wednesday': ['DS', 'AH', 'CB'],
        'Thursday': ['CB'],
        'Friday': ['CB', 'AL']
    }

# --- STYLING ---
def apply_custom_css():
    st.markdown("""
    <style>
        /* Styling for first multiselect widget */
        div[data-testid="stMultiSelect"]:first-of-type [data-baseweb="tag"] {
            background-color: #2E8B57;
            border-radius: 0.5rem;
        }
        div[data-testid="stMultiSelect"]:first-of-type [data-baseweb="tag"] span {
            color: white !important;
        }
        div[data-testid="stMultiSelect"]:first-of-type [data-baseweb="tag"] span[role="button"] {
            color: white !important;
        }
        
        /* Loading spinner improvements */
        .stSpinner > div > div {
            border-color: #2E8B57 !important;
        }
    </style>
    """, unsafe_allow_html=True)

# --- DATABASE OPERATIONS ---
class DatabaseManager:
    def __init__(self, supabase_client: Client):
        self.supabase = supabase_client
    
    def get_mdk_history(self) -> Dict[str, int]:
        """Get MDK assignment counts for all employees."""
        try:
            response = self.supabase.table("mdk_assignments").select("employee").execute()
            assignments = response.data or []
            mdk_counts = {}
            for assignment in assignments:
                emp = assignment['employee']
                mdk_counts[emp] = mdk_counts.get(emp, 0) + 1
            return {k: v for k, v in mdk_counts.items() if v > 0}
        except Exception as e:
            st.error(f"Fel vid hämtning av MDK-historik: {e}")
            return {}
    
    def get_employee_mdk_count(self, employee: str) -> int:
        """Get MDK count for specific employee."""
        try:
            response = self.supabase.table("mdk_assignments").select("week", count="exact").eq("employee", employee).execute()
            return response.count or 0
        except Exception:
            return 0
    
    def save_mdk_assignments(self, assignments: Dict[str, str], week: int):
        """Save MDK assignments to database."""
        try:
            for day, emp in assignments.items():
                self.supabase.table("mdk_assignments").upsert({
                    "week": week, 
                    "day": day, 
                    "employee": emp
                }).execute()
        except Exception as e:
            st.error(f"Fel vid sparning av MDK-uppdrag: {e}")
    
    def clear_mdk_history(self):
        """Clear all MDK history."""
        try:
            self.supabase.table("mdk_assignments").delete().neq("week", -1).execute()
            return True
        except Exception as e:
            st.error(f"Fel vid radering av MDK-historik: {e}")
            return False
    
    def get_work_rates(self) -> Dict[str, int]:
        """Get work rates for all employees."""
        try:
            response = self.supabase.table("work_rates").select("*").execute()
            return {row['employee']: row['rate'] for row in (response.data or [])}
        except Exception as e:
            st.error(f"Fel vid hämtning av arbetstid: {e}")
            return {}
    
    def save_work_rates(self, work_rates: Dict[str, int]):
        """Save work rates to database."""
        try:
            records = [{"employee": emp, "rate": rate} for emp, rate in work_rates.items()]
            self.supabase.table("work_rates").upsert(records).execute()
            return True
        except Exception as e:
            st.error(f"Fel vid sparning av arbetstid: {e}")
            return False

# --- SCHEDULE GENERATION ---
class ScheduleGenerator:
    def __init__(self, db_manager: DatabaseManager):
        self.db_manager = db_manager
        
        # Excel column mappings
        self.klin_cols = {'Monday': 'B', 'Tuesday': 'F', 'Wednesday': 'J', 'Thursday': 'N', 'Friday': 'R'}
        self.screen_cols = {'Monday': 'C', 'Tuesday': 'G', 'Wednesday': 'K', 'Thursday': 'O', 'Friday': 'S'}
        self.mdk_cols = {'Monday': 'D', 'Tuesday': 'H', 'Thursday': 'P'}
        self.lunchvakt_col = {'Wednesday': 'L'}
        self.lab_rows = {
            'morning1': {'LAB 3': 4, 'LAB 6': 5, 'LAB 9': 6, 'LAB 10': 7},
            'morning2': {'LAB 3': 9, 'LAB 6': 10, 'LAB 9': 11, 'LAB 10': 12},
            'afternoon1': {'LAB 3': 14, 'LAB 6': 15, 'LAB 9': 16, 'LAB 10': 17}
        }
    
    def calculate_mdk_scores(self, available_employees: List[str], work_rates: Dict[str, int], 
                           assigned_this_week: Dict[str, int]) -> Dict[str, float]:
        """Calculate MDK assignment scores for available employees."""
        scores = {}
        for emp in available_employees:
            history_count = self.db_manager.get_employee_mdk_count(emp)
            rate = work_rates.get(emp, 100)
            this_week_penalty = assigned_this_week.get(emp, 0) * 10
            
            if rate > 0:
                scores[emp] = (history_count / rate) + this_week_penalty
            else:
                scores[emp] = float('inf')
        
        return scores
    
    def assign_mdk_roles(self, available_employees: List[str], unavailable_per_day: Dict[str, List[str]], 
                        work_rates: Dict[str, int]) -> Dict[str, str]:
        """Assign MDK roles for the week."""
        mdk_assignments = {}
        assigned_this_week = {emp: 0 for emp in available_employees}
        
        for day in Config.MDK_DAYS:
            avail_for_day = [
                emp for emp in available_employees 
                if emp not in unavailable_per_day.get(day, []) and work_rates.get(emp, 0) > 0
            ]
            
            if not avail_for_day:
                st.warning(f"Inga tillgängliga medarbetare för MDK på {Config.DAYS_SV[day]}")
                continue
            
            scores = self.calculate_mdk_scores(avail_for_day, work_rates, assigned_this_week)
            chosen = min(scores, key=scores.get)
            mdk_assignments[day] = chosen
            assigned_this_week[chosen] += 1
        
        return mdk_assignments
    
    def assign_lab_roles(self, available_for_labs: List[str], prefer_from: List[str] = None) -> Dict[str, str]:
        """Assign lab roles, optionally preferring from a specific group."""
        if not available_for_labs:
            return {}
        
        labs = Config.LABS.copy()
        random.shuffle(labs)
        
        # Prefer from specific group if provided
        if prefer_from:
            preferred = [emp for emp in available_for_labs if emp in prefer_from]
            if len(preferred) >= min(4, len(available_for_labs)):
                lab_people = random.sample(preferred, min(4, len(preferred)))
            else:
                lab_people = preferred + random.sample(
                    [emp for emp in available_for_labs if emp not in preferred],
                    min(4 - len(preferred), len(available_for_labs) - len(preferred))
                )
        else:
            lab_people = random.sample(available_for_labs, min(len(available_for_labs), 4))
        
        return dict(zip(lab_people, labs[:len(lab_people)]))
    

    
    def generate_schedule(self, available_employees: List[str], unavailable_per_day: Dict[str, List[str]], 
                         work_rates: Dict[str, int], current_week: int) -> io.BytesIO:
        """Generate the complete schedule."""
        
        # Assign MDK roles first
        mdk_assignments = self.assign_mdk_roles(available_employees, unavailable_per_day, work_rates)
        
        # Load template and set week number
        wb = openpyxl.load_workbook('template.xlsx')
        sheet = wb['Blad1']
        sheet['A1'] = f"v.{current_week}"
        
        # Process each day
        for day in Config.DAYS:
            avail_day = [emp for emp in available_employees if emp not in unavailable_per_day.get(day, [])]
            mdk = mdk_assignments.get(day)
            
            # Determine lab candidates
            lab_candidates = avail_day.copy()
            if mdk in lab_candidates and day in ['Tuesday', 'Thursday']:  # Full-day MDK
                lab_candidates.remove(mdk)
            
            # Morning assignments
            morning_candidates = lab_candidates.copy()
            if mdk in morning_candidates and day == 'Monday':  # Half-day MDK
                morning_candidates.remove(mdk)
            
            morning_assign = self.assign_lab_roles(morning_candidates)
            morning_remainder = [emp for emp in avail_day if emp not in morning_assign]
            
            # Fill morning Screen/MR
            screen_mr_morning = [emp for emp in morning_remainder if (emp != mdk or day not in Config.MDK_DAYS)]
            sheet[f"{self.screen_cols[day]}3"] = '/'.join(sorted(screen_mr_morning))
            
            # Fill morning lab positions
            klin_col = self.klin_cols[day]
            for person, lab in morning_assign.items():
                for time_slot in ['morning1', 'morning2']:
                    sheet[f"{klin_col}{self.lab_rows[time_slot][lab]}"] = person
            
            # Afternoon assignments (not Friday)
            if day != 'Friday':
                available_for_afternoon = lab_candidates.copy()
                if mdk in available_for_afternoon and day in ['Tuesday', 'Thursday']:
                    available_for_afternoon.remove(mdk)
                
                # Prefer those on Screen/MR in morning for afternoon labs
                morning_screen_mr = [emp for emp in morning_remainder if emp in available_for_afternoon]
                afternoon_assign = self.assign_lab_roles(available_for_afternoon, morning_screen_mr)
                
                # Try to avoid same lab assignment
                afternoon_assign = self.avoid_same_lab_assignment(morning_assign, list(afternoon_assign.keys()))
                
                # Fill afternoon lab positions
                for person, lab in afternoon_assign.items():
                    sheet[f"{klin_col}{self.lab_rows['afternoon1'][lab]}"] = person
                
                # Fill afternoon Screen/MR
                afternoon_screen_mr = [emp for emp in available_for_afternoon if emp not in afternoon_assign]
                sheet[f"{self.screen_cols[day]}14"] = '/'.join(sorted(afternoon_screen_mr))
            
            # Fill MDK and Lunch Guard
            if mdk:
                sheet[f"{self.mdk_cols[day]}3"] = mdk
            elif day == 'Wednesday' and avail_day:
                sheet[f"{self.lunchvakt_col['Wednesday']}3"] = random.choice(avail_day)
        
        # Save MDK assignments to database
        self.db_manager.save_mdk_assignments(mdk_assignments, current_week)
        
        # Return Excel file as BytesIO
        output_file = io.BytesIO()
        wb.save(output_file)
        output_file.seek(0)
        return output_file

# --- FILE OPERATIONS ---
class FileManager:
    def __init__(self, supabase_client: Client):
        self.supabase = supabase_client
    
    def get_uploaded_files(self) -> List[str]:
        """Get list of uploaded schedule files."""
        try:
            bucket_files = self.supabase.storage.from_("schedules").list()
            return [f['name'] for f in bucket_files] if bucket_files else []
        except Exception:
            return []
    
    def upload_file(self, file_name: str, file_content: bytes) -> bool:
        """Upload file to storage."""
        try:
            self.supabase.storage.from_("schedules").upload(file_name, file_content, {"upsert": "true"})
            return True
        except Exception as e:
            st.error(f"Fel vid uppladdning av {file_name}: {e}")
            return False
    
    def process_uploaded_schedule(self, file_name: str, week: int) -> int:
        """Process uploaded schedule and extract MDK assignments."""
        try:
            downloaded = self.supabase.storage.from_("schedules").download(file_name)
            if not downloaded:
                return 0
            
            wb = openpyxl.load_workbook(io.BytesIO(downloaded))
            sheet = wb["Blad1"] if "Blad1" in wb.sheetnames else wb.active
            
            mdk_cols = {'Monday': 'D', 'Tuesday': 'H', 'Thursday': 'P'}
            parsed_days = 0
            
            for day, col in mdk_cols.items():
                cell_value = sheet.cell(row=3, column=openpyxl.utils.column_index_from_string(col)).value
                if cell_value and isinstance(cell_value, str) and cell_value.strip() in Config.PRE_POP_EMPLOYEES:
                    try:
                        self.supabase.table("mdk_assignments").upsert({
                            "week": week, 
                            "day": day, 
                            "employee": cell_value.strip()
                        }).execute()
                        parsed_days += 1
                    except Exception as e:
                        st.error(f"Misslyckades med att spara MDK-uppdrag för {Config.DAYS_SV[day]}: {e}")
            
            return parsed_days
        except Exception as e:
            st.error(f"Fel vid bearbetning av {file_name}: {e}")
            return 0

# --- MAIN APPLICATION ---
def main():
    # Apply custom styling
    apply_custom_css()
    
    # Initialize session state
    if 'confirm_delete' not in st.session_state:
        st.session_state.confirm_delete = False
    
    # Initialize Supabase client
    try:
        supabase_url = st.secrets["SUPABASE_URL"]
        supabase_key = st.secrets["SUPABASE_KEY"]
        supabase = create_client(supabase_url, supabase_key)
    except Exception as e:
        st.error(f"Fel vid anslutning till databas: {e}")
        st.stop()
    
    # Initialize managers
    db_manager = DatabaseManager(supabase)
    file_manager = FileManager(supabase)
    schedule_generator = ScheduleGenerator(db_manager)
    
    st.title("Schemagenerator för världens bästa enhet!")
    
    # Employee availability input
    available_week = st.multiselect(
        "Initialer för samtliga medarbetare",
        options=Config.PRE_POP_EMPLOYEES,
        default=Config.PRE_POP_EMPLOYEES
    )
    
    unavailable_whole_week = st.multiselect(
        "Initialer för medarbetare som är otillgängliga hela veckan",
        options=available_week,
        default=[]
    )
    
    available_employees = [emp for emp in available_week if emp not in unavailable_whole_week]
    
    # Daily unavailability input
    with st.expander("Ange otillgänglighet per dag", expanded=True):
        unavailable_per_day = {}
        for day in Config.DAYS:
            default_values = [
                emp for emp in Config.PRE_UNAVAILABLE.get(day, []) 
                if emp in available_employees
            ]
            unavailable_per_day[day] = st.multiselect(
                f"Initialer för otillgängliga medarbetare på {Config.DAYS_SV[day]}",
                options=available_employees,
                default=default_values
            )
    
    # MDK Overview
    with st.expander("MDK-fördelning de senaste månaderna (stapeldiagram)"):
        mdk_counts = db_manager.get_mdk_history()
        
        if mdk_counts:
            sorted_items = sorted(mdk_counts.items(), key=lambda x: x[1], reverse=True)
            employees, counts = zip(*sorted_items)
            
            fig = px.bar(
                x=employees,
                y=counts,
                labels={'x': 'Medarbetare', 'y': 'Antal MDK'},
                title="MDK-fördelning de senaste 2 månaderna",
                color=counts,
                color_continuous_scale="blugrn",
            )
            fig.update_coloraxes(showscale=False)
            fig.update_yaxes(dtick=1)
            
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Inga MDK-uppdrag i historiken ännu.")
    
    # Historical schedules upload
    current_week = date.today().isocalendar()[1]
    with st.expander("Historiska scheman (Senaste 8 veckorna)"):
        file_names = file_manager.get_uploaded_files()
        
        for i in range(1, 9):
            week = current_week - i
            file_name = f"week_{week}.xlsx"
            
            col1, col2 = st.columns([2, 3])
            with col1:
                st.write(f"**Vecka {week}**")
                status = "✅ Uppladdat" if file_name in file_names else "❌ Ej uppladdat"
                st.write(f"Status: {status}")
            
            with col2:
                uploader = st.file_uploader(
                    f"Ladda upp/ersätt schema för vecka {week}", 
                    type="xlsx", 
                    key=f"hist_{week}"
                )
                
                if uploader:
                    file_content = uploader.getvalue()
                    if file_manager.upload_file(file_name, file_content):
                        st.success(f"✅ Laddade upp {file_name}")
                        time.sleep(1)
                        
                        parsed_days = file_manager.process_uploaded_schedule(file_name, week)
                        st.success(f"📊 Läste in MDK-uppdrag för {parsed_days} dagar")
        
        # MDK history deletion
        st.markdown("---")
        if st.button("🗑️ Rensa all MDK-historik", help="Raderar all historisk MDK-data"):
            st.session_state.confirm_delete = True
        
        if st.session_state.confirm_delete:
            st.warning("⚠️ Är du säker på att du vill radera all MDK-historik? Detta kan inte ångras.")
            
            col1, col2, _ = st.columns([1.5, 1, 4])
            with col1:
                if st.button("✅ Ja, radera all historik", type="primary"):
                    if db_manager.clear_mdk_history():
                        st.success("🗑️ All MDK-historik har raderats.")
                        st.session_state.confirm_delete = False
                        time.sleep(2)
                        st.rerun()
            with col2:
                if st.button("❌ Avbryt"):
                    st.session_state.confirm_delete = False
                    st.rerun()
    
    # Work rates management
    if 'work_rates' not in st.session_state:
        default_rates = {emp: 100 for emp in Config.PRE_POP_EMPLOYEES}
        db_rates = db_manager.get_work_rates()
        st.session_state['work_rates'] = {**default_rates, **db_rates}
    
    with st.expander("Klinisk arbetstid per medarbetare (justera vid behov)"):
        col1, col2 = st.columns(2)
        for i, emp in enumerate(Config.PRE_POP_EMPLOYEES):
            col = col1 if i % 2 == 0 else col2
            with col:
                st.session_state['work_rates'][emp] = st.number_input(
                    f"{emp} arbetstid (0-100%)",
                    min_value=0,
                    max_value=100,
                    value=int(st.session_state['work_rates'].get(emp, 100)),
                    step=5,
                    key=f"rate_{emp}"
                )
        
        if st.button("💾 Spara arbetstid till databasen"):
            if db_manager.save_work_rates(st.session_state['work_rates']):
                st.success("✅ Arbetstid sparad!")
                with st.spinner("Uppdaterar..."):
                    time.sleep(1)
                st.rerun()
    
    # Schedule generation
    st.markdown("---")
    if st.button("🎯 Generera Schema", type="primary"):
        with st.spinner("Genererar schema, vänligen vänta..."):
            try:
                schedule_file = schedule_generator.generate_schedule(
                    available_employees, 
                    unavailable_per_day, 
                    st.session_state['work_rates'], 
                    current_week
                )
                
                st.success("✅ Schemat har genererats!")
                st.download_button(
                    label="📥 Ladda ner schemat",
                    data=schedule_file,
                    file_name=f"veckoschema_v{current_week}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"❌ Fel vid generering av schema: {e}")

if __name__ == "__main__":
    main()