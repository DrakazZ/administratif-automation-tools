import pandas as pd
from datetime import datetime, timedelta
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass
import logging
import os
import sys

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Get executable directory - FIXED for PyInstaller
def get_exe_directory():
    """Obtenir le répertoire de l'exécutable - fonctionne avec PyInstaller"""
    if getattr(sys, 'frozen', False):
        # Si l'application est compilée avec PyInstaller
        return os.path.dirname(sys.executable)
    else:
        # Si on exécute le script Python directement
        return os.path.dirname(os.path.abspath(__file__))

@dataclass
class Room:
    name: str
    capacity: int  # Number of groups it can hold simultaneously
    
    def can_hold(self, num_groups: int) -> bool:
        return num_groups <= self.capacity

@dataclass 
class Subject:
    name: str
    duration: int  # in minutes
    
@dataclass
class Group:
    name: str
    subjects: List[Subject]
    
    @property
    def num_subjects(self) -> int:
        return len(self.subjects)

@dataclass
class Teacher:
    name: str
    subjects_taught: List[str]  # Subjects this teacher teaches (cannot supervise)
    
    def can_supervise(self, subject_name: str) -> bool:
        return subject_name not in self.subjects_taught

@dataclass
class TimeSlot:
    date: datetime
    start_time: datetime
    end_time: datetime
    duration: int  # in minutes
    
    def __str__(self):
        return f"{self.date.strftime('%Y-%m-%d')} {self.start_time.strftime('%H:%M')}-{self.end_time.strftime('%H:%M')}"

@dataclass
class ExamAssignment:
    group: Group
    subject: Subject
    time_slot: TimeSlot
    room: Room
    supervisor: Teacher

class ExamScheduler:
    def __init__(self):
        self.rooms: List[Room] = []
        self.groups: List[Group] = []
        self.teachers: List[Teacher] = []
        self.assignments: List[ExamAssignment] = []
        self.time_slots: List[TimeSlot] = []
        
        # Get executable directory for file outputs
        self.exe_dir = get_exe_directory()
        
        # Tracking availability
        self.room_availability: Dict[str, Dict[str, int]] = {}  # room_name -> slot_str -> available_capacity
        self.teacher_availability: Dict[str, List[str]] = {}    # teacher_name -> list of available slot_strs
        self.group_daily_count: Dict[str, Dict[str, int]] = {} # group_name -> date_str -> exam_count
        
    def load_excel_data(self, filepath: str) -> bool:
        """Load data from Excel file with three sheets: Salles, Groupes, Enseignants"""
        try:
            # Read all sheets - Updated for French names
            sheet_names = ['Salles', 'Groupes', 'Enseignants']
            try:
                excel_data = pd.read_excel(filepath, sheet_name=sheet_names)
            except ValueError:
                # Fallback to English names if French not found
                sheet_names = ['Rooms', 'Groups', 'Teachers']
                excel_data = pd.read_excel(filepath, sheet_name=sheet_names)
            
            # Load Rooms/Salles
            self._load_rooms(excel_data[sheet_names[0]], sheet_names[0])
            logger.info(f"Chargé {len(self.rooms)} salles")
            
            # Load Groups/Groupes
            self._load_groups(excel_data[sheet_names[1]], sheet_names[1])
            logger.info(f"Chargé {len(self.groups)} groupes")
            
            # Load Teachers/Enseignants
            self._load_teachers(excel_data[sheet_names[2]], sheet_names[2])
            logger.info(f"Chargé {len(self.teachers)} enseignants")
            
            return True
            
        except Exception as e:
            logger.error(f"Erreur lors du chargement des données Excel: {str(e)}")
            return False
    
    def _load_rooms(self, rooms_df: pd.DataFrame, sheet_name: str):
        """Load rooms from DataFrame - supports both French and English"""
        self.rooms = []
        name_col = 'Nom de la Salle' if sheet_name == 'Salles' else 'Room Name'
        capacity_col = 'Capacité' if sheet_name == 'Salles' else 'Capacity'
        
        for _, row in rooms_df.iterrows():
            room = Room(
                name=str(row[name_col]).strip(),
                capacity=int(row[capacity_col])
            )
            self.rooms.append(room)
    
    def _load_groups(self, groups_df: pd.DataFrame, sheet_name: str):
        """Load groups and their subjects from DataFrame - supports both French and English"""
        self.groups = []
        is_french = sheet_name == 'Groupes'
        
        name_col = 'Nom du Groupe' if is_french else 'Group Name'
        num_col = 'Nb Matières' if is_french else 'Num Subjects'
        subject_prefix = 'Matière' if is_french else 'Subject'
        duration_prefix = 'Durée' if is_french else 'Duration'
        
        for _, row in groups_df.iterrows():
            group_name = str(row[name_col]).strip()
            num_subjects = int(row[num_col])
            
            subjects = []
            for i in range(1, num_subjects + 1):
                subject_col = f'{subject_prefix} {i}'
                duration_col = f'{duration_prefix} {i}'
                
                # Handle French format with "(min)" suffix
                if is_french and f'{duration_prefix} {i} (min)' in row:
                    duration_col = f'{duration_prefix} {i} (min)'
                
                if subject_col in row and duration_col in row:
                    subject_name = str(row[subject_col]).strip()
                    duration = int(row[duration_col])
                    
                    subjects.append(Subject(name=subject_name, duration=duration))
            
            group = Group(name=group_name, subjects=subjects)
            self.groups.append(group)
    
    def _load_teachers(self, teachers_df: pd.DataFrame, sheet_name: str):
        """Load teachers and their subjects from DataFrame - supports both French and English"""
        self.teachers = []
        is_french = sheet_name == 'Enseignants'
        
        name_col = 'Nom de l\'Enseignant' if is_french else 'Teacher Name'
        num_col = 'Nb Matières' if is_french else 'Num Subjects'
        subject_prefix = 'Matière' if is_french else 'Subject'
        
        for _, row in teachers_df.iterrows():
            teacher_name = str(row[name_col]).strip()
            num_subjects = int(row[num_col])
            
            subjects_taught = []
            for i in range(1, num_subjects + 1):
                subject_col = f'{subject_prefix} {i}'
                if subject_col in row and pd.notna(row[subject_col]):
                    subject_name = str(row[subject_col]).strip()
                    subjects_taught.append(subject_name)
            
            teacher = Teacher(name=teacher_name, subjects_taught=subjects_taught)
            self.teachers.append(teacher)
    
    def generate_time_slots(self, start_date: datetime, end_date: datetime, 
                          daily_start: str, daily_end: str, break_minutes: int) -> bool:
        """Generate all possible time slots for the exam period"""
        try:
            # Parse daily times
            start_time = datetime.strptime(daily_start, '%H:%M').time()
            end_time = datetime.strptime(daily_end, '%H:%M').time()
            
            self.time_slots = []
            current_date = start_date
            
            while current_date <= end_date:
                # Skip weekends (optional - you said might be overkill)
                if current_date.weekday() < 5:  # 0-6, where 0 is Monday
                    daily_slots = self._generate_daily_slots(
                        current_date, start_time, end_time, break_minutes
                    )
                    self.time_slots.extend(daily_slots)
                
                current_date += timedelta(days=1)
            
            logger.info(f"Généré {len(self.time_slots)} créneaux horaires")
            return True
            
        except Exception as e:
            logger.error(f"Erreur lors de la génération des créneaux: {str(e)}")
            return False
    
    def _generate_daily_slots(self, date: datetime, start_time, end_time, break_minutes: int) -> List[TimeSlot]:
        """Generate time slots for a single day"""
        slots = []
        
        # Get all unique durations from all subjects
        durations = set()
        for group in self.groups:
            for subject in group.subjects:
                durations.add(subject.duration)
        
        # Sort durations to try longer ones first (better slot utilization)
        durations = sorted(durations, reverse=True)
        
        # Create slots for each duration type
        for duration in durations:
            current_time = datetime.combine(date, start_time)
            end_datetime = datetime.combine(date, end_time)
            
            while current_time + timedelta(minutes=duration) <= end_datetime:
                slot_end = current_time + timedelta(minutes=duration)
                
                slot = TimeSlot(
                    date=date,
                    start_time=current_time,
                    end_time=slot_end,
                    duration=duration
                )
                slots.append(slot)
                
                # Move to next slot with break
                current_time = slot_end + timedelta(minutes=break_minutes)
        
        return slots
    
    def initialize_availability_tracking(self):
        """Initialize availability matrices for rooms, teachers, and groups"""
        # Initialize room availability
        self.room_availability = {}
        for room in self.rooms:
            self.room_availability[room.name] = {}
            for slot in self.time_slots:
                slot_key = str(slot)
                self.room_availability[room.name][slot_key] = room.capacity
        
        # Initialize teacher availability  
        self.teacher_availability = {}
        for teacher in self.teachers:
            self.teacher_availability[teacher.name] = []
            for slot in self.time_slots:
                self.teacher_availability[teacher.name].append(str(slot))
        
        # Initialize group daily exam count
        self.group_daily_count = {}
        for group in self.groups:
            self.group_daily_count[group.name] = {}
            unique_dates = set(slot.date.strftime('%Y-%m-%d') for slot in self.time_slots)
            for date_str in unique_dates:
                self.group_daily_count[group.name][date_str] = 0
        
        logger.info("Suivi de disponibilité initialisé")
    
    def print_summary(self):
        """Print a summary of loaded data"""
        print("\n" + "="*50)
        print("PLANIFICATEUR D'EXAMENS - RÉSUMÉ DES DONNÉES")
        print("="*50)
        
        print(f"\n📚 GROUPES ({len(self.groups)}):")
        for group in self.groups:
            print(f"  • {group.name}: {group.num_subjects} matières")
            for subject in group.subjects:
                print(f"    - {subject.name} ({subject.duration}min)")
        
        print(f"\n🏫 SALLES ({len(self.rooms)}):")
        for room in self.rooms:
            print(f"  • {room.name}: capacité {room.capacity} groupe(s)")
        
        print(f"\n👨‍🏫 ENSEIGNANTS ({len(self.teachers)}):")
        for teacher in self.teachers:
            subjects_str = ", ".join(teacher.subjects_taught)
            print(f"  • {teacher.name}: enseigne [{subjects_str}]")
        
        print(f"\n⏰ CRÉNEAUX HORAIRES ({len(self.time_slots)}):")
        if self.time_slots:
            dates = set(slot.date.strftime('%Y-%m-%d') for slot in self.time_slots)
            print(f"  • {len(dates)} jours d'examens")
            print(f"  • Exemples de créneaux:")
            for slot in self.time_slots[:3]:
                print(f"    - {slot}")
            if len(self.time_slots) > 3:
                print(f"    ... et {len(self.time_slots) - 3} autres")

    def schedule_all_exams(self, max_exams_per_day: int = 2) -> Dict[str, List[str]]:
        """
        Main scheduling algorithm - assigns all exams to time slots
        Priority: Groups with more exams first (your strategy)
        Returns: Dictionary of conflicts/errors if any
        """
        logger.info("🚀 Début de la planification des examens...")
        
        # Clear previous assignments
        self.assignments = []
        conflicts = {"unassigned": [], "warnings": []}
        
        # Sort groups by number of subjects (descending) - your priority rule
        sorted_groups = sorted(self.groups, key=lambda g: g.num_subjects, reverse=True)
        
        logger.info(f"📊 Ordre de planification: {[f'{g.name}({g.num_subjects} examens)' for g in sorted_groups]}")
        
        # Schedule each group's exams
        for group in sorted_groups:
            group_conflicts = self._schedule_group_exams(group, max_exams_per_day)
            conflicts["unassigned"].extend(group_conflicts)
        
        # Generate summary
        total_exams = sum(len(g.subjects) for g in self.groups)
        scheduled_exams = len(self.assignments)
        
        logger.info(f"📈 Planification terminée: {scheduled_exams}/{total_exams} examens assignés")
        
        if conflicts["unassigned"]:
            logger.warning(f"⚠️  {len(conflicts['unassigned'])} examens n'ont pas pu être assignés automatiquement")
        
        return conflicts
    
    def _schedule_group_exams(self, group: Group, max_exams_per_day: int) -> List[str]:
        """Schedule all exams for a single group"""
        conflicts = []
        
        # Sort subjects by duration (longer first - harder to place)
        sorted_subjects = sorted(group.subjects, key=lambda s: s.duration, reverse=True)
        
        for subject in sorted_subjects:
            success = self._assign_exam(group, subject, max_exams_per_day)
            if not success:
                conflict_msg = f"❌ {group.name} - {subject.name} ({subject.duration}min): Impossible de planifier l'examen"
                conflicts.append(conflict_msg)
                logger.warning(conflict_msg)
        
        return conflicts
    
    def _assign_exam(self, group: Group, subject: Subject, max_exams_per_day: int) -> bool:
        """Try to assign a single exam to an available slot"""
        
        # Filter slots by duration
        compatible_slots = [slot for slot in self.time_slots if slot.duration == subject.duration]
        
        # Sort slots chronologically (front-load strategy)
        compatible_slots.sort(key=lambda s: (s.date, s.start_time))
        
        for slot in compatible_slots:
            slot_key = str(slot)
            date_key = slot.date.strftime('%Y-%m-%d')
            
            # Check constraints
            if not self._check_group_availability(group, slot, date_key, max_exams_per_day):
                continue
                
            room = self._find_available_room(slot_key)
            if not room:
                continue
                
            teacher = self._find_available_teacher(subject, slot_key)
            if not teacher:
                continue
            
            # All constraints satisfied - make assignment
            assignment = ExamAssignment(group, subject, slot, room, teacher)
            self.assignments.append(assignment)
            
            # Update availability tracking
            self._update_availability_after_assignment(assignment, slot_key, date_key)
            
            logger.info(f"✅ Assigné: {group.name} - {subject.name} → {slot} dans {room.name} avec {teacher.name}")
            return True
        
        return False
    
    def _check_group_availability(self, group: Group, slot: TimeSlot, date_key: str, max_exams_per_day: int) -> bool:
        """Check if group can take exam in this slot"""
        current_count = self.group_daily_count[group.name][date_key]
        return current_count < max_exams_per_day
    
    def _find_available_room(self, slot_key: str) -> Optional[Room]:
        """Find a room with available capacity for this slot"""
        # Prefer single-capacity rooms first (save multi-capacity for later)
        available_rooms = []
        
        for room in self.rooms:
            if self.room_availability[room.name][slot_key] > 0:
                available_rooms.append((room, room.capacity))
        
        if not available_rooms:
            return None
        
        # Sort by capacity (prefer smaller rooms first - save big ones for busy periods)
        available_rooms.sort(key=lambda x: x[1])
        return available_rooms[0][0]
    
    def _find_available_teacher(self, subject: Subject, slot_key: str) -> Optional[Teacher]:
        """Find an available teacher who can supervise this subject"""
        for teacher in self.teachers:
            if (slot_key in self.teacher_availability[teacher.name] and 
                teacher.can_supervise(subject.name)):
                return teacher
        return None
    
    def _update_availability_after_assignment(self, assignment: ExamAssignment, slot_key: str, date_key: str):
        """Update all availability matrices after making an assignment"""
        # Update room availability
        self.room_availability[assignment.room.name][slot_key] -= 1
        
        # Update teacher availability
        if slot_key in self.teacher_availability[assignment.supervisor.name]:
            self.teacher_availability[assignment.supervisor.name].remove(slot_key)
        
        # Update group daily count
        self.group_daily_count[assignment.group.name][date_key] += 1

# ================== SECTION 4: EXCEL EXPORT (French Headers) - FIXED PATHS ==================

    def export_group_calendars(self, filename: str = None) -> bool:
        """Export individual group calendars to Excel - FIXED to use exe directory"""
        try:
            if filename is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"Calendriers_Groupes_{timestamp}.xlsx"
            
            # FIXED: Use exe directory instead of script directory
            filepath = os.path.join(self.exe_dir, filename)
            
            # Create a workbook with separate sheet for each group
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                
                for group in self.groups:
                    group_assignments = [a for a in self.assignments if a.group.name == group.name]
                    
                    if group_assignments:
                        # Sort by date and time
                        group_assignments.sort(key=lambda a: (a.time_slot.date, a.time_slot.start_time))
                        
                        # Create DataFrame with French headers
                        data = []
                        for assignment in group_assignments:
                            data.append({
                                'Date': assignment.time_slot.date.strftime('%d/%m/%Y'),
                                'Début': assignment.time_slot.start_time.strftime('%H:%M'),
                                'Fin': assignment.time_slot.end_time.strftime('%H:%M'),
                                'Matière': assignment.subject.name,
                                'Salle': assignment.room.name,
                                'Surveillant': assignment.supervisor.name
                            })
                        
                        df = pd.DataFrame(data)
                        
                        # Clean sheet name for Excel
                        sheet_name = group.name.replace('/', '_')[:31]  # Excel sheet name limit
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        
                        logger.info(f"✅ Calendrier exporté pour {group.name}")
            
            logger.info(f"📁 Calendriers des groupes exportés vers: {filepath}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Erreur lors de l'export des calendriers: {str(e)}")
            return False
    
    def export_global_plan(self, filename: str = None) -> bool:
        """Export global exam plan for administration - FIXED to use exe directory"""
        try:
            if not self.assignments:
                logger.warning("⚠️  Aucune assignation à exporter")
                return False
            
            if filename is None:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"Planning_Global_Examens_{timestamp}.xlsx"
            
            # FIXED: Use exe directory instead of script directory
            filepath = os.path.join(self.exe_dir, filename)
            
            # Sort all assignments chronologically  
            sorted_assignments = sorted(self.assignments, 
                                      key=lambda a: (a.time_slot.date, a.time_slot.start_time))
            
            # Create DataFrame with French headers
            data = []
            for assignment in sorted_assignments:
                data.append({
                    'Date': assignment.time_slot.date.strftime('%d/%m/%Y'),
                    'Heure': f"{assignment.time_slot.start_time.strftime('%H:%M')}–{assignment.time_slot.end_time.strftime('%H:%M')}",
                    'Groupe': assignment.group.name,
                    'Matière': assignment.subject.name,
                    'Salle': assignment.room.name,
                    'Surveillant': assignment.supervisor.name,
                    'Durée': f"{assignment.subject.duration} min"
                })
            
            df = pd.DataFrame(data)
            
            # Export to Excel
            df.to_excel(filepath, sheet_name='Planning Global', index=False)
            
            logger.info(f"📁 Planning global exporté vers: {filepath}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Erreur lors de l'export du planning global: {str(e)}")
            return False

    def get_schedule_statistics(self) -> Dict:
        """Get scheduling statistics for UI display"""
        if not self.assignments:
            return {"total_exams": 0, "scheduled": 0, "success_rate": 0}
        
        total_exams = sum(len(group.subjects) for group in self.groups)
        scheduled = len(self.assignments)
        success_rate = (scheduled / total_exams) * 100 if total_exams > 0 else 0
        
        # Daily distribution
        daily_counts = {}
        for assignment in self.assignments:
            date_str = assignment.time_slot.date.strftime('%d/%m/%Y')
            daily_counts[date_str] = daily_counts.get(date_str, 0) + 1
        
        return {
            "total_exams": total_exams,
            "scheduled": scheduled, 
            "unscheduled": total_exams - scheduled,
            "success_rate": round(success_rate, 1),
            "daily_distribution": daily_counts,
            "exam_days": len(daily_counts)
        }

# ================== SECTION 5: FRENCH QT GUI - FIXED PATHS ==================

try:
    from PyQt5.QtWidgets import (QApplication, QMainWindow, QVBoxLayout, QHBoxLayout, 
                               QWidget, QPushButton, QLabel, QLineEdit, QDateEdit, 
                               QSpinBox, QTextEdit, QFileDialog, QMessageBox, 
                               QProgressBar, QGroupBox, QGridLayout, QTimeEdit)
    from PyQt5.QtCore import Qt, QDate, QTime, QThread, pyqtSignal
    from PyQt5.QtGui import QFont
    PYQT_AVAILABLE = True
except ImportError:
    PYQT_AVAILABLE = False
    print("PyQt5 non disponible. Interface en ligne de commande uniquement.")

if PYQT_AVAILABLE:
    class SchedulingThread(QThread):
        """Background thread for scheduling to keep UI responsive"""
        progress = pyqtSignal(str)  # Progress messages
        finished = pyqtSignal(dict)  # Results
        
        def __init__(self, scheduler, start_date, end_date, daily_start, daily_end, break_minutes, max_exams_per_day):
            super().__init__()
            self.scheduler = scheduler
            self.start_date = start_date
            self.end_date = end_date
            self.daily_start = daily_start
            self.daily_end = daily_end
            self.break_minutes = break_minutes
            self.max_exams_per_day = max_exams_per_day
        
        def run(self):
            try:
                self.progress.emit("⏰ Génération des créneaux horaires...")
                success = self.scheduler.generate_time_slots(
                    self.start_date, self.end_date, self.daily_start, self.daily_end, self.break_minutes
                )
                
                if not success:
                    self.finished.emit({"error": "Échec de la génération des créneaux horaires"})
                    return
                
                self.progress.emit("🔄 Initialisation du système de suivi...")
                self.scheduler.initialize_availability_tracking()
                
                self.progress.emit("📋 Planification des examens...")
                conflicts = self.scheduler.schedule_all_exams(self.max_exams_per_day)
                
                stats = self.scheduler.get_schedule_statistics()
                
                self.finished.emit({"success": True, "conflicts": conflicts, "stats": stats})
                
            except Exception as e:
                self.finished.emit({"error": f"Erreur de planification: {str(e)}"})

    class ExamSchedulerGUI(QMainWindow):
        def __init__(self):
            super().__init__()
            self.scheduler = ExamScheduler()
            self.scheduling_thread = None
            self.setup_ui()
            
        def setup_ui(self):
            """إعداد واجهة المستخدم العربية"""
            self.setWindowTitle("نظام جدولة الامتحانات الآلي")
            self.setGeometry(100, 100, 800, 600)
            
            # Set Arabic font
            font = QFont("Arial", 10)
            self.setFont(font)
            
            # Central widget
            central_widget = QWidget()
            self.setCentralWidget(central_widget)
            
            # Main layout
            layout = QVBoxLayout(central_widget)
            
            # File input section
            file_group = QGroupBox("📁 ملف البيانات")
            file_layout = QHBoxLayout(file_group)
            
            self.file_path_label = QLabel("لم يتم اختيار ملف")
            self.browse_button = QPushButton("تحديد ملف Excel")
            self.browse_button.clicked.connect(self.browse_file)
            
            file_layout.addWidget(QLabel("ملف البيانات:"))
            file_layout.addWidget(self.file_path_label, 1)
            file_layout.addWidget(self.browse_button)
            
            layout.addWidget(file_group)
            
            # Schedule parameters section
            params_group = QGroupBox("⚙️ إعدادات الجدولة")
            params_layout = QGridLayout(params_group)
            
            # Date range
            params_layout.addWidget(QLabel("تاريخ البداية:"), 0, 0)
            self.start_date = QDateEdit()
            self.start_date.setDate(QDate.currentDate())
            params_layout.addWidget(self.start_date, 0, 1)
            
            params_layout.addWidget(QLabel("تاريخ النهاية:"), 0, 2)
            self.end_date = QDateEdit()
            self.end_date.setDate(QDate.currentDate().addDays(7))
            params_layout.addWidget(self.end_date, 0, 3)
            
            # Daily schedule
            params_layout.addWidget(QLabel("بداية اليوم:"), 1, 0)
            self.daily_start = QTimeEdit()
            self.daily_start.setTime(QTime(8, 30))
            params_layout.addWidget(self.daily_start, 1, 1)
            
            params_layout.addWidget(QLabel("نهاية اليوم:"), 1, 2)
            self.daily_end = QTimeEdit()
            self.daily_end.setTime(QTime(17, 0))
            params_layout.addWidget(self.daily_end, 1, 3)
            
            # Break and limits
            params_layout.addWidget(QLabel("الاستراحة (دقيقة):"), 2, 0)
            self.break_minutes = QSpinBox()
            self.break_minutes.setRange(15, 60)
            self.break_minutes.setValue(30)
            params_layout.addWidget(self.break_minutes, 2, 1)
            
            params_layout.addWidget(QLabel("حد الامتحانات يومياً:"), 2, 2)
            self.max_exams_per_day = QSpinBox()
            self.max_exams_per_day.setRange(1, 5)
            self.max_exams_per_day.setValue(2)
            params_layout.addWidget(self.max_exams_per_day, 2, 3)
            
            layout.addWidget(params_group)
            
            # Action buttons
            buttons_layout = QHBoxLayout()
            
            self.schedule_button = QPushButton("🚀 توليد الجدول")
            self.schedule_button.clicked.connect(self.start_scheduling)
            self.schedule_button.setEnabled(False)
            
            self.export_groups_button = QPushButton("📊 اصدار جداول المجموعات")
            self.export_groups_button.clicked.connect(self.export_group_calendars)
            self.export_groups_button.setEnabled(False)
            
            self.export_global_button = QPushButton("📋 اصدار الجدول العام")
            self.export_global_button.clicked.connect(self.export_global_plan)
            self.export_global_button.setEnabled(False)
            
            buttons_layout.addWidget(self.schedule_button)
            buttons_layout.addWidget(self.export_groups_button)
            buttons_layout.addWidget(self.export_global_button)
            
            layout.addLayout(buttons_layout)
            
            # Progress bar
            self.progress_bar = QProgressBar()
            self.progress_bar.setVisible(False)
            layout.addWidget(self.progress_bar)
            
            # Results text area
            results_group = QGroupBox("📈 النتائج والإحصائيات")
            results_layout = QVBoxLayout(results_group)
            
            self.results_text = QTextEdit()
            self.results_text.setReadOnly(True)
            self.results_text.setMaximumHeight(200)
            results_layout.addWidget(self.results_text)
            
            layout.addWidget(results_group)
        
        def browse_file(self):
            """Sélectionner un fichier Excel"""
            file_path, _ = QFileDialog.getOpenFileName(
                self, "Sélectionner le fichier de données", "", "Fichiers Excel (*.xlsx *.xls)"
            )
            
            if file_path:
                self.file_path_label.setText(os.path.basename(file_path))
                
                # Try to load the file
                success = self.scheduler.load_excel_data(file_path)
                if success:
                    self.schedule_button.setEnabled(True)
                    self.results_text.append("✅ Données chargées avec succès!")
                    self.scheduler.print_summary()
                else:
                    QMessageBox.warning(self, "Erreur", "Échec du chargement des données. Vérifiez le format du fichier.")
        
        def start_scheduling(self):
            """بدء عملية الجدولة"""
            if self.scheduling_thread and self.scheduling_thread.isRunning():
                return
            
            # Prepare parameters
            start_date = self.start_date.date().toPyDate()
            end_date = self.end_date.date().toPyDate()
            daily_start = self.daily_start.time().toString("HH:mm")
            daily_end = self.daily_end.time().toString("HH:mm")
            break_mins = self.break_minutes.value()
            max_exams = self.max_exams_per_day.value()
            
            # Disable buttons and show progress
            self.schedule_button.setEnabled(False)
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)  # Indeterminate
            
            # Start background thread
            self.scheduling_thread = SchedulingThread(
                self.scheduler, start_date, end_date, daily_start, daily_end, break_mins, max_exams
            )
            self.scheduling_thread.progress.connect(self.update_progress)
            self.scheduling_thread.finished.connect(self.scheduling_finished)
            self.scheduling_thread.start()
        
        def update_progress(self, message):
            """تحديث رسالة التقدم"""
            self.results_text.append(message)
        
        def scheduling_finished(self, result):
            """انتهاء عملية الجدولة"""
            self.progress_bar.setVisible(False)
            self.schedule_button.setEnabled(True)
            
            if "error" in result:
                QMessageBox.critical(self, "خطأ", result["error"])
                return
            
            if result["success"]:
                stats = result["stats"]
                conflicts = result["conflicts"]
                
                # Display results
                results_text = f"""
    ✅ تمت الجدولة بنجاح!

    📊 الإحصائيات:
    • العدد الإجمالي للامتحانات: {stats['total_exams']}
    • الامتحانات المجدولة: {stats['scheduled']}
    • الامتحانات غير المجدولة: {stats['unscheduled']}
    • نسبة النجاح: {stats['success_rate']}%
    • عدد أيام الامتحانات: {stats['exam_days']}

    📅 توزيع الامتحانات:
    """
                for date, count in stats['daily_distribution'].items():
                    results_text += f"• {date}: {count} امتحان\n"
                
                if conflicts["unassigned"]:
                    results_text += f"\n⚠️ الامتحانات غير المجدولة ({len(conflicts['unassigned'])}):\n"
                    for conflict in conflicts["unassigned"][:5]:  # Show first 5
                        results_text += f"• {conflict}\n"
                    if len(conflicts["unassigned"]) > 5:
                        results_text += f"... و {len(conflicts['unassigned']) - 5} امتحان آخر\n"
                
                self.results_text.setText(results_text)
                
                # Enable export buttons
                if stats['scheduled'] > 0:
                    self.export_groups_button.setEnabled(True)
                    self.export_global_button.setEnabled(True)
        
        def export_group_calendars(self):
            """Exporter les calendriers des groupes - FIXED PATH"""
            success = self.scheduler.export_group_calendars()
            if success:
                # Show the actual file path
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"Calendriers_Groupes_{timestamp}.xlsx"
                filepath = os.path.join(self.scheduler.exe_dir, filename)
                QMessageBox.information(self, "نجح", f"تم اصدار جداول المجموعات في:\n{filepath}")
            else:
                QMessageBox.warning(self, "خطأ", "فشل في اصدار جداول المجموعات")
        
        def export_global_plan(self):
            """Exporter le planning global - FIXED PATH"""
            success = self.scheduler.export_global_plan()
            if success:
                # Show the actual file path
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"Planning_Global_Examens_{timestamp}.xlsx"
                filepath = os.path.join(self.scheduler.exe_dir, filename)
                QMessageBox.information(self, "نجح", f"تم اصدار الجدول العام في:\n{filepath}")
            else:
                QMessageBox.warning(self, "خطأ", "فشل في اصدار الجدول العام")

# ================== SECTION 6: MAIN ENTRY POINT ==================

def main():
    if PYQT_AVAILABLE:
        app = QApplication(sys.argv)
        
        # Set application properties
        app.setApplicationName("نظام جدولة الامتحانات")
        app.setApplicationVersion("1.0")
        
        # Create and show main window
        window = ExamSchedulerGUI()
        window.show()
        
        sys.exit(app.exec_())
    else:
        print("Interface graphique non disponible. Utilisation en ligne de commande.")
        run_tests()

def run_tests():
    """Exécuter les tests en ligne de commande"""
    scheduler = ExamScheduler()
    
    # Test with sample data
    print("🚀 Test des Composants Principaux du Planificateur")
    
    # Create sample data for testing
    scheduler.rooms = [
        Room("101", 1),
        Room("102", 1), 
        Room("103", 2)  # Can hold 2 groups
    ]
    
    scheduler.groups = [
        Group("1ère AP1", [
            Subject("Histoire", 90),
            Subject("A.A. Mod", 120)
        ]),
        Group("2ème AP", [
            Subject("Mathématiques", 120),
            Subject("Informatique", 90),
            Subject("Physique", 90)
        ])
    ]
    
    scheduler.teachers = [
        Teacher("Allagui", ["Histoire"]),
        Teacher("TakTak", ["Esthétique"]),
        Teacher("Ghassallah", ["A.A. Mod"]),
        Teacher("Benali", []),  # Can supervise anything
        Teacher("Khadija", ["Mathématiques"])
    ]
    
    # Generate time slots and schedule
    start_date = datetime(2026, 1, 6)  # Monday
    end_date = datetime(2026, 1, 10)   # Friday
    
    success = scheduler.generate_time_slots(start_date, end_date, "08:30", "17:00", 30)
    if success:
        scheduler.initialize_availability_tracking()
        conflicts = scheduler.schedule_all_exams(2)
        
        scheduler.print_summary()
        print(f"\n📊 Résultats de la Planification:")
        stats = scheduler.get_schedule_statistics()
        for key, value in stats.items():
            print(f"  {key}: {value}")
        
        # Show assignments
        print(f"\n📋 ASSIGNATIONS:")
        for assignment in scheduler.assignments:
            print(f"  {assignment.group.name} - {assignment.subject.name} → {assignment.time_slot} dans {assignment.room.name} avec {assignment.supervisor.name}")
        
        # Test exports - FIXED to show actual paths
        print(f"\n💾 Test des exports...")
        print(f"Répertoire de sortie: {scheduler.exe_dir}")
        
        group_success = scheduler.export_group_calendars("Test_Calendriers_Groupes.xlsx")
        global_success = scheduler.export_global_plan("Test_Planning_Global.xlsx")
        
        if group_success and global_success:
            print("✅ Fichiers d'export créés avec succès!")
            print(f"📁 Les fichiers sont sauvegardés dans: {scheduler.exe_dir}")
        else:
            print("❌ Erreur lors de la création des fichiers d'export")

# Example usage and testing
if __name__ == "__main__":
    import sys
    
    # Check command line arguments
    if len(sys.argv) > 1 and sys.argv[1] == "test":
        # Run tests
        run_tests()
    else:
        # Run GUI by default
        main()