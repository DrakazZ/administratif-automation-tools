# leave_tracker_app.py
import sys
import os
import sqlite3
from datetime import datetime
from openpyxl import Workbook
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QComboBox, QMessageBox, QTabWidget, QSpinBox
)

DB_PATH = os.path.join(os.path.dirname(__file__), 'leave_data.db')

LEAVE_TYPES = {
    'سنوية': {'yearly_add': 45, 'reset': False},
    'استثنائية': {'yearly_add': 6, 'reset': True},
    'مرضية': {'yearly_add': 0, 'reset': False},
    'تعوضية': {'yearly_add': 0, 'reset': False},
    'اضافية': {'yearly_add': 0, 'reset': False},
}

def get_column_index(cursor, column_name):
    for i, col in enumerate(cursor.description):
        if col[0] == column_name:
            return i
    raise ValueError(f"Column '{column_name}' not found in results")

class LeaveApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("منضومة العطل")
        self.resize(600, 400)
        self.create_db()

        self.tabs = QTabWidget()
        self.tabs.addTab(self.add_employee_tab(), "اضافة موظف")
        self.tabs.addTab(self.manage_leave_tab(), "ادارة العطل")
        self.tabs.addTab(self.summary_tab(), "مخلص العطل")
        self.tabs.addTab(self.reset_tab(),"اعادة ضبط سنوي")

        layout = QVBoxLayout()
        layout.addWidget(self.tabs)
        self.setLayout(layout)

    def create_db(self):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        fields = ", ".join([
            f"{t}_taken INTEGER DEFAULT 0, {t}_left INTEGER DEFAULT {LEAVE_TYPES[t]['yearly_add']}"
            for t in LEAVE_TYPES
        ])
        cursor.execute(f"""
            CREATE TABLE IF NOT EXISTS employees (
                reg_number TEXT PRIMARY KEY,
                name TEXT,
                surname TEXT,
                {fields}
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS leave_history (
                reg_number TEXT,
                year INTEGER,
                leave_type TEXT,
                taken INTEGER,
                PRIMARY KEY (reg_number, year, leave_type)
            )
        """)

        conn.commit()
        conn.close()

    def add_employee_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        self.name_input = QLineEdit()
        self.surname_input = QLineEdit()
        self.reg_input = QLineEdit()
        layout.addWidget(QLabel("الاسم"))
        layout.addWidget(self.name_input)
        layout.addWidget(QLabel("اللقب"))
        layout.addWidget(self.surname_input)
        layout.addWidget(QLabel("المعرف الوحيد"))
        layout.addWidget(self.reg_input)

        self.initial_left_inputs = {}
        for t in LEAVE_TYPES:
            input_field = QSpinBox()
            input_field.setRange(0, 365)
            self.initial_left_inputs[t] = input_field
            layout.addWidget(QLabel(f"  عدد ايام العطل {t} المتبقية"))
            layout.addWidget(input_field)

        add_btn = QPushButton("اضافة موظف")
        add_btn.clicked.connect(self.add_employee)
        layout.addWidget(add_btn)

        tab.setLayout(layout)
        return tab

    def manage_leave_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        self.reg_manage = QLineEdit()
        self.days_input = QSpinBox()
        self.days_input.setRange(1, 365)
        self.leave_type = QComboBox()
        self.leave_type.addItems(LEAVE_TYPES.keys())

        layout.addWidget(QLabel("المعرف الوحيد"))
        layout.addWidget(self.reg_manage)
        layout.addWidget(QLabel("عدد ايام العطل"))
        layout.addWidget(self.days_input)
        layout.addWidget(QLabel("نوع العطل"))
        layout.addWidget(self.leave_type)

        btn_layout = QHBoxLayout()
        take_btn = QPushButton("اخذ ايام عطل")
        take_btn.clicked.connect(self.take_days)
        add_btn = QPushButton("اضافة ايام عطل")
        add_btn.clicked.connect(self.add_days)
        btn_layout.addWidget(take_btn)
        btn_layout.addWidget(add_btn)

        layout.addLayout(btn_layout)
        tab.setLayout(layout)
        return tab

    def summary_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        self.reg_summary = QLineEdit()
        self.year_input = QSpinBox()
        self.year_input.setRange(2000, 2100)
        self.year_input.setValue(datetime.now().year)

        self.summary_label = QLabel("")

        layout.addWidget(QLabel("المعرف الوحيد"))
        layout.addWidget(self.reg_summary)
        layout.addWidget(QLabel("السنة"))
        layout.addWidget(self.year_input)

        summary_btn = QPushButton("عرض الملخص")
        summary_btn.clicked.connect(self.show_summary)
        layout.addWidget(summary_btn)

        # ✅ Export personal summary to Excel
        export_btn = QPushButton("تحميل ملخص الموظف (Excel)")
        export_btn.clicked.connect(self.export_selected_summary)
        layout.addWidget(export_btn)

        # ✅ Export all summaries to Excel
        export_all_btn = QPushButton("تحميل ملخص كل الموظفين (Excel)")
        export_all_btn.clicked.connect(self.export_all_summaries)
        layout.addWidget(export_all_btn)

        layout.addWidget(self.summary_label)
        tab.setLayout(layout)
        return tab


    def reset_tab(self):
        tab = QWidget()
        layout = QVBoxLayout()

        reset_btn = QPushButton("اعادة ضبط سنوي")
        reset_btn.clicked.connect(self.perform_reset)
        layout.addWidget(reset_btn)

        tab.setLayout(layout)
        return tab

    def add_employee(self):
        name = self.name_input.text()
        surname = self.surname_input.text()
        reg = self.reg_input.text()

        if not (name and surname and reg):
            QMessageBox.warning(self, "معلومات خاطئة", "الرجاء ادخال المعلومات المطلوبة")
            return

        values = {
            f"{t}_left": self.initial_left_inputs[t].value() for t in LEAVE_TYPES
        }

        columns = ", ".join(["reg_number", "name", "surname"] + list(values.keys()))
        placeholders = ", ".join(["?"] * len(values))
        sql = f"INSERT INTO employees ({columns}) VALUES (?, ?, ?, {placeholders})"

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        try:
            cursor.execute(sql, (reg, name, surname, *values.values()))
            conn.commit()
            QMessageBox.information(self, "نجاح", "تم اضافة الموظف بنجاح")
        except sqlite3.IntegrityError:
            QMessageBox.warning(self, "خطا", "المعرف الوحيد مستخدم من قبل")
        conn.close()

    def take_days(self):
        self.modify_days(-1)

    def add_days(self):
        self.modify_days(1)


    def modify_days(self, direction):
        reg = self.reg_manage.text()
        days = self.days_input.value()
        leave_type = self.leave_type.currentText()

        if not reg or not leave_type or days <= 0:
            QMessageBox.warning(self, "خطأ", "الرجاء ملء كل الحقول بشكل صحيح")
            return

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM employees WHERE reg_number = ?", (reg,))
        row = cursor.fetchone()

        if not row:
            QMessageBox.warning(self, "خطأ", "الموظف غير موجود")
            return

        taken_col = f"{leave_type}_taken"
        left_col = f"{leave_type}_left"
        col_names = [col[0] for col in cursor.description]

        try:
            left_index = col_names.index(left_col)
            taken_index = col_names.index(taken_col)
        except ValueError:
            QMessageBox.warning(self, "خطأ", "نوع العطلة غير موجود في قاعدة البيانات")
            return

        current_left = row[left_index]

        if direction == -1:
            # ✅ Taking leave — restrict if not enough days
            if current_left < days:
                QMessageBox.warning(self, "أيام غير كافية", "لا يوجد أيام كافية لهذه العطلة")
                return

            # Update employee leave balance
            cursor.execute(f"""
                UPDATE employees SET
                    {taken_col} = {taken_col} + ?,
                    {left_col} = {left_col} - ?
                WHERE reg_number = ?
            """, (days, days, reg))

            # Update or insert into leave_history
            year = datetime.now().year  # import datetime at the top if needed

            cursor.execute("""
                SELECT taken FROM leave_history
                WHERE reg_number = ? AND year = ? AND leave_type = ?
            """, (reg, year, leave_type))

            existing = cursor.fetchone()

            if existing:
                cursor.execute("""
                    UPDATE leave_history
                    SET taken = taken + ?
                    WHERE reg_number = ? AND year = ? AND leave_type = ?
                """, (days, reg, year, leave_type))
            else:
                cursor.execute("""
                    INSERT INTO leave_history (reg_number, year, leave_type, taken)
                    VALUES (?, ?, ?, ?)
                """, (reg, year, leave_type, days))


        elif direction == 1:
            # ✅ Adding days — allow resulting negative or high values (no restriction)
            cursor.execute(f"""
                UPDATE employees SET
                    {left_col} = {left_col} + ?
                WHERE reg_number = ?
            """, (days, reg))

        conn.commit()
        conn.close()
        QMessageBox.information(self, "تحديث", "تم تحديث بيانات العطلة بنجاح")

    def show_summary(self):
        reg = self.reg_summary.text()
        year = self.year_input.value()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        # Fetch employee data first
        cursor.execute("SELECT * FROM employees WHERE reg_number = ?", (reg,))
        row = cursor.fetchone()
        if not row:
            QMessageBox.warning(self, "خطأ", "الموظف غير موجود")
            return

        employee_cols = [col[0] for col in cursor.description]
        name, surname = row[1], row[2]
        summary = f"الاسم: {name} {surname}\nالسنة: {year}\n\n"

        # Now fetch history AFTER employee data
        cursor.execute("SELECT leave_type, taken FROM leave_history WHERE reg_number = ? AND year = ?", (reg, year))
        history = dict(cursor.fetchall())

        for t in LEAVE_TYPES:
            taken = history.get(t, 0)
            try:
                left_index = employee_cols.index(f"{t}_left")
                left = row[left_index]
            except ValueError:
                left = "؟"

            summary += f"{t}: أخذ = {taken}, متبقي = {left}\n"

        self.summary_label.setText(summary)
        conn.close()



    def perform_reset(self):
        year = datetime.now().year
        today = datetime.now()
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        RESET_ALLOWED_DAY = 10

        # Check if today is before Jan 10th
        if today.month == 1 and today.day <= RESET_ALLOWED_DAY:
            QMessageBox.warning(self, "تحذير", f"الرجاء الانتظار حتى {RESET_ALLOWED_DAY} يناير للقيام بإعادة الضبط السنوي.")
            conn.close()
            return

        # Optional: If reset already done this year, block or warn (from previous logic)
        cursor.execute("SELECT COUNT(*) FROM leave_history WHERE year = ?", (year,))
        already_reset = cursor.fetchone()[0] > 0
        if already_reset:
            reply = QMessageBox.question(self, "تحذير", 
                f"تمت إعادة الضبط لهذا العام {year} مسبقاً. هل تريد المتابعة؟", 
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                conn.close()
                return

        reply = QMessageBox.question(self, "تأكيد", f"هل أنت متأكد أنك تريد إعادة ضبط السنة {year}؟",
                             QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply != QMessageBox.Yes:
            conn.close()
            return

        # 🛡️ Prevent duplicate reset for the same year
        cursor.execute("SELECT COUNT(*) FROM leave_history WHERE year = ?", (year,))
        already_reset = cursor.fetchone()[0] > 0

        if already_reset:
            QMessageBox.warning(self, "تحذير", f"تم بالفعل إعادة الضبط لهذه السنة: {year}")
            conn.close()
            return

        backup_filename = f"نسخة_احتياطية_قبل_إعادة_الضبط_{year}.xlsx"
        self.export_current_summary_to_excel(backup_filename)

        cursor.execute("SELECT reg_number FROM employees")
        reg_numbers = [row[0] for row in cursor.fetchall()]

        for reg in reg_numbers:
            for t, props in LEAVE_TYPES.items():
                taken_col = f"{t}_taken"
                cursor.execute(f"SELECT {taken_col} FROM employees WHERE reg_number = ?", (reg,))
                taken = cursor.fetchone()[0]
                cursor.execute("""
                    INSERT OR REPLACE INTO leave_history (reg_number, year, leave_type, taken)
                    VALUES (?, ?, ?, ?)
                """, (reg, year, t, taken))

        for t, props in LEAVE_TYPES.items():
            if props['reset']:
                cursor.execute(f"UPDATE employees SET {t}_left = {props['yearly_add']}, {t}_taken = 0")
            else:
                cursor.execute(f"UPDATE employees SET {t}_left = {t}_left + {props['yearly_add']}, {t}_taken = 0")

        conn.commit()
        conn.close()
        QMessageBox.information(self, "تمت العملية", "تم اعادة الضبط السنوي العطل بنجاح")

    def export_employee_summary_to_excel(self, reg, year):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        # Get employee data and store column info
        cursor.execute("SELECT * FROM employees WHERE reg_number = ?", (reg,))
        employee_row = cursor.fetchone()
        if not employee_row:
            QMessageBox.warning(self, "خطأ", "الموظف غير موجود")
            conn.close()
            return

        # Store column names from employee table BEFORE changing queries
        employee_cols = [col[0] for col in cursor.description]
        name, surname = employee_row[1], employee_row[2]
        
        # Now get history data
        cursor.execute("SELECT leave_type, taken FROM leave_history WHERE reg_number = ? AND year = ?", (reg, year))
        history = dict(cursor.fetchall())

        wb = Workbook()
        ws = wb.active
        ws.title = f"ملخص {reg}"

        ws.append(["رقم التسجيل", "الاسم", "اللقب", "السنة"])
        ws.append([reg, name, surname, year])
        ws.append([])
        ws.append(["نوع العطلة", "الأيام المستعملة", "الأيام المتبقية"])

        for t in LEAVE_TYPES:
            taken = history.get(t, 0)
            # Use employee_cols instead of cursor.description
            try:
                left_index = employee_cols.index(f"{t}_left")
                left = employee_row[left_index]
            except ValueError:
                left = "غير متوفر"  # Fallback if column doesn't exist
            
            ws.append([t, taken, left])

        filename = f"ملخص_{reg}_{year}.xlsx"
        wb.save(filename)
        conn.close()
        QMessageBox.information(self, "تم الحفظ", f"تم حفظ ملخص الموظف في الملف:\n{filename}")

    
    def export_current_summary_to_excel(self, filename="نسخة_احتياطية_قبل_إعادة_الضبط.xlsx"):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        wb = Workbook()
        ws = wb.active
        ws.title = "نسخة احتياطية"

        headers = ["رقم التسجيل", "الاسم", "اللقب"] + [
            f"المستعملة ({t})" for t in LEAVE_TYPES
        ] + [
            f"المتبقية ({t})" for t in LEAVE_TYPES
        ]
        ws.append(headers)

        cursor.execute("SELECT * FROM employees")
        rows = cursor.fetchall()
        for row in rows:
            base = list(row[:3])  # رقم التسجيل، الاسم، اللقب
            taken = [row[get_column_index(cursor, f"{t}_taken")] for t in LEAVE_TYPES]
            left = [row[get_column_index(cursor, f"{t}_left")] for t in LEAVE_TYPES]
            ws.append(base + taken + left)

        wb.save(filename)
        conn.close()

        QMessageBox.information(self, "تم الحفظ", f"تم حفظ الملخص في الملف:\n{filename}")

    def export_all_summaries_to_excel(self, year):
        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        wb = Workbook()
        ws = wb.active
        ws.title = f"ملخص {year}"

        headers = ["رقم التسجيل", "الاسم", "اللقب"] + [
            f"المستعملة ({t})" for t in LEAVE_TYPES
        ] + [
            f"المتبقية ({t})" for t in LEAVE_TYPES
        ]
        ws.append(headers)

        # Get all employee data and store column info
        cursor.execute("SELECT * FROM employees")
        all_employee_rows = cursor.fetchall()
        employee_cols = [col[0] for col in cursor.description]  # Store column names

        for employee_row in all_employee_rows:
            reg = employee_row[0]
            name = employee_row[1]
            surname = employee_row[2]
            
            # Get history for this employee
            cursor.execute("SELECT leave_type, taken FROM leave_history WHERE reg_number = ? AND year = ?", (reg, year))
            history = dict(cursor.fetchall())
            
            # Build taken and left arrays using stored column info
            taken = [history.get(t, 0) for t in LEAVE_TYPES]
            left = []
            
            for t in LEAVE_TYPES:
                try:
                    left_index = employee_cols.index(f"{t}_left")
                    left.append(employee_row[left_index])
                except ValueError:
                    left.append("غير متوفر")  # Fallback if column doesn't exist
            
            ws.append([reg, name, surname] + taken + left)

        filename = f"ملخص_جميع_الموظفين_{year}.xlsx"
        wb.save(filename)
        conn.close()
        QMessageBox.information(self, "تم الحفظ", f"تم حفظ الملخص في الملف:\n{filename}")

    def export_selected_summary(self):
        reg = self.reg_summary.text()
        year = self.year_input.value()
        if not reg:
            QMessageBox.warning(self, "خطأ", "يرجى إدخال رقم التسجيل أولاً")
            return
        self.export_employee_summary_to_excel(reg, year)

    def export_all_summaries(self):
        year = self.year_input.value()
        self.export_all_summaries_to_excel(year)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = LeaveApp()
    window.show()
    sys.exit(app.exec_())
