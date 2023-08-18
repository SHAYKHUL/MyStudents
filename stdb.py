import sqlite3
import sys
import csv
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from docx import Document

from PyQt5.QtGui import QStandardItemModel
from PyQt5.QtGui import QStandardItem, QStandardItemModel
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QMessageBox, QTableWidgetItem, QTableWidget, QDialog, QFormLayout, QHeaderView, QComboBox, QCompleter, QAction, QMenu, QMenuBar,QFileDialog
class StudentDialog(QDialog):
    def __init__(self, parent=None, student_id=None, cursor=None):
        super().__init__(parent)
        self.student_id = student_id
        self.cursor = cursor
        self.setWindowTitle("Add Student" if student_id is None else "Edit Student")
        self.setup_ui()
        
    def setup_ui(self):
        layout = QFormLayout()
        
        self.name_edit = QLineEdit()
        self.class_edit = QLineEdit()
        self.roll_edit = QLineEdit()
        self.father_edit = QLineEdit()
        self.mother_edit = QLineEdit()
        self.phone_edit = QLineEdit()
        self.address_edit = QLineEdit()

        layout.addRow("Name:", self.name_edit)
        layout.addRow("Class:", self.class_edit)
        layout.addRow("Roll:", self.roll_edit)
        layout.addRow("Father's Name:", self.father_edit)
        layout.addRow("Mother's Name:", self.mother_edit)
        layout.addRow("Phone No:", self.phone_edit)
        layout.addRow("Address:", self.address_edit)

        self.save_button = QPushButton("Save")
        self.save_button.clicked.connect(self.save_student)
        layout.addWidget(self.save_button)

        self.setLayout(layout)

        if self.student_id:
            self.load_student_data()

    def load_student_data(self):
        query = "SELECT * FROM students WHERE id = ?"
        self.cursor.execute(query, (self.student_id,))
        student = self.cursor.fetchone()
        self.name_edit.setText(student[1])
        self.class_edit.setText(student[2])
        self.roll_edit.setText(student[3])
        self.father_edit.setText(student[4])
        self.mother_edit.setText(student[5])
        self.phone_edit.setText(student[6])
        self.address_edit.setText(student[7])

    def save_student(self):
        name = self.name_edit.text()
        class_ = self.class_edit.text()
        roll = self.roll_edit.text()
        father = self.father_edit.text()
        mother = self.mother_edit.text()
        phone = self.phone_edit.text()
        address = self.address_edit.text()

        if not name or not class_ or not roll:
            QMessageBox.warning(self, "Warning", "Name, Class, and Roll are required fields.")
            return

        if self.student_id:
            query = "UPDATE students SET name=?, class=?, roll=?, father_name=?, mother_name=?, phone=?, address=? WHERE id=?"
            self.cursor.execute(query, (name, class_, roll, father, mother, phone, address, self.student_id))
        else:
            query = "INSERT INTO students (name, class, roll, father_name, mother_name, phone, address) VALUES (?, ?, ?, ?, ?, ?, ?)"
            self.cursor.execute(query, (name, class_, roll, father, mother, phone, address))

        self.cursor.connection.commit()
        self.accept()

class StudentManagementApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Student Management Software")
        self.setGeometry(100, 100, 1000, 600)
        
        self.setup_ui()
        self.init_db()
        self.setup_menu_bar()

    def setup_ui(self):
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.layout = QVBoxLayout()

        self.search_label = QLabel("Search by Name:")
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("Type to search")
        self.search_edit.textChanged.connect(self.update_suggestions)
        
        self.search_suggestions = QCompleter()
        self.search_suggestions.setCaseSensitivity(False)
        self.search_edit.setCompleter(self.search_suggestions)

        self.search_button = QPushButton("Search")
        self.search_button.clicked.connect(self.search_student)

        self.sort_by_label = QLabel("Sort by:")
        self.sort_by_combo = QComboBox()
        self.sort_by_combo.addItems(["Name", "Class", "Roll"])
        self.sort_by_combo.currentIndexChanged.connect(self.sort_students)
        
        self.filter_label = QLabel("Filter by Class:")
        self.filter_combo = QComboBox()
        self.filter_combo.addItem("All")
        self.filter_combo.currentIndexChanged.connect(self.filter_students)

        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels(["ID", "Name", "Class", "Roll", "Father's Name", "Mother's Name", "Phone No", "Address"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        self.edit_button = QPushButton("Edit")
        self.edit_button.clicked.connect(self.edit_selected_student)
        self.add_button = QPushButton("Add Student")
        self.add_button.clicked.connect(self.add_student)

        button_layout = QHBoxLayout()
        button_layout.addWidget(self.edit_button)
        button_layout.addWidget(self.add_button)

        self.layout.addWidget(self.search_label)
        self.layout.addWidget(self.search_edit)
        self.layout.addWidget(self.search_button)
        self.layout.addWidget(self.sort_by_label)
        self.layout.addWidget(self.sort_by_combo)
        self.layout.addWidget(self.filter_label)
        self.layout.addWidget(self.filter_combo)
        self.layout.addWidget(self.table)
        self.layout.addLayout(button_layout)

        self.central_widget.setLayout(self.layout)

    def setup_menu_bar(self):
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("File")

        add_student_action = QAction("Add Student", self)
        add_student_action.triggered.connect(self.add_student)
        file_menu.addAction(add_student_action)

        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Help menu
        help_menu = menubar.addMenu("Help")

        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)
       
        #Export data
        export_csv_action = QAction("Export Data as CSV", self)
        export_csv_action.triggered.connect(self.export_data_csv)
        file_menu.addAction(export_csv_action)

        export_pdf_action = QAction("Export Data as PDF", self)
        export_pdf_action.triggered.connect(self.export_data_pdf)
        file_menu.addAction(export_pdf_action)

        export_docx_action = QAction("Export Data as DOCX", self)
        export_docx_action.triggered.connect(self.export_data_docx)
        file_menu.addAction(export_docx_action)



    def show_about_dialog(self):
        QMessageBox.about(self, "About", "Student Management System v1.0\n\nA simple application for managing student information.")
    
    def export_data_csv(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Save CSV File", "", "CSV Files (*.csv);;All Files (*)")
        if not file_path:
            return
        
        query = "SELECT * FROM students"
        self.cursor.execute(query)
        students = self.cursor.fetchall()
        
        try:
            with open(file_path, mode='w', newline='') as file:
                writer = csv.writer(file)
                writer.writerow(["ID", "Name", "Class", "Roll", "Father's Name", "Mother's Name", "Phone No", "Address"])
                writer.writerows(students)
                
            QMessageBox.information(self, "Export Successful", "Student data exported to CSV successfully.")
        except Exception as e:
            QMessageBox.warning(self, "Export Error", f"An error occurred during export: {str(e)}")

    def export_data_pdf(self):
        query = "SELECT * FROM students"
        self.cursor.execute(query)
        students = self.cursor.fetchall()
        
        export_to_pdf(students)

    def export_data_docx(self):
        query = "SELECT * FROM students"
        self.cursor.execute(query)
        students = self.cursor.fetchall()
        
        export_to_document(students)



    def init_db(self):
        self.conn = sqlite3.connect("student_db.db")
        self.cursor = self.conn.cursor()
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS students (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                class TEXT,
                roll TEXT,
                father_name TEXT,
                mother_name TEXT,
                phone TEXT,
                address TEXT
            )
        ''')
        self.conn.commit()

    def update_suggestions(self):
        name_query = "SELECT name FROM students WHERE name LIKE ?"
        self.cursor.execute(name_query, (f'%{self.search_edit.text()}%',))
        names = self.cursor.fetchall()
        
        suggestion_items = [QStandardItem(name[0]) for name in names]
        
        self.search_suggestions.setModel(QStandardItemModel())
        self.search_suggestions.model().appendRow(suggestion_items)
        
    def search_student(self):
        name = self.search_edit.text()
        query = "SELECT * FROM students WHERE name LIKE ?"
        self.cursor.execute(query, (f'%{name}%',))
        students = self.cursor.fetchall()
        
        self.table.setRowCount(0)
        for row_idx, student in enumerate(students):
            self.table.insertRow(row_idx)
            for col_idx, data in enumerate(student):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(data)))

    def sort_students(self):
        column_idx = self.sort_by_combo.currentIndex()
        if column_idx == 0:
            sort_column = "name"
        elif column_idx == 1:
            sort_column = "class"
        elif column_idx == 2:
            sort_column = "roll"
        
        query = f"SELECT * FROM students ORDER BY {sort_column}"
        self.cursor.execute(query)
        students = self.cursor.fetchall()
        
        self.table.setRowCount(0)
        for row_idx, student in enumerate(students):
            self.table.insertRow(row_idx)
            for col_idx, data in enumerate(student):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(data)))

    def filter_students(self):
        selected_class = self.filter_combo.currentText()

        if selected_class == "All":
            query = "SELECT * FROM students"
            self.cursor.execute(query)
            students = self.cursor.fetchall()
        else:
            query = "SELECT * FROM students WHERE class = ?"
            self.cursor.execute(query, (selected_class,))
            students = self.cursor.fetchall()

        self.table.setRowCount(0)
        for row_idx, student in enumerate(students):
            self.table.insertRow(row_idx)
            for col_idx, data in enumerate(student):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(str(data)))



    def add_student(self):
        dialog = StudentDialog(self, cursor=self.cursor)
        if dialog.exec_() == QDialog.Accepted:
            self.search_student()

    def edit_selected_student(self):
        selected_row = self.table.currentRow()
        if selected_row >= 0:
            student_id = self.table.item(selected_row, 0).text()
            dialog = StudentDialog(self, student_id=int(student_id), cursor=self.cursor)
            if dialog.exec_() == QDialog.Accepted:
                self.search_student()


def export_to_pdf(data):
    file_path, _ = QFileDialog.getSaveFileName(None, "Save PDF File", "", "PDF Files (*.pdf);;All Files (*)")
    if not file_path:
        return

    doc = SimpleDocTemplate(file_path, pagesize=letter)
    elements = []

    table_data = [["ID", "Name", "Class", "Roll", "Father's Name", "Mother's Name", "Phone No", "Address"]]
    table_data.extend(data)

    table = Table(table_data)
    table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                               ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                               ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                               ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                               ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                               ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                               ('GRID', (0, 0), (-1, -1), 1, colors.black)]))

    elements.append(table)
    doc.build(elements)

    QMessageBox.information(None, "Export Successful", "Student data exported to PDF successfully.")

def export_to_document(data):
    file_path, _ = QFileDialog.getSaveFileName(None, "Save Document File", "", "Document Files (*.docx);;All Files (*)")
    if not file_path:
        return

    doc = Document()
    table = doc.add_table(rows=1, cols=8)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "ID"
    hdr_cells[1].text = "Name"
    hdr_cells[2].text = "Class"
    hdr_cells[3].text = "Roll"
    hdr_cells[4].text = "Father's Name"
    hdr_cells[5].text = "Mother's Name"
    hdr_cells[6].text = "Phone No"
    hdr_cells[7].text = "Address"

    for student in data:
        row_cells = table.add_row().cells
        for i, field in enumerate(student):
            row_cells[i].text = str(field)


    doc.save(file_path)

    QMessageBox.information(None, "Export Successful", "Student data exported to document successfully.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = StudentManagementApp()
    window.show()
    sys.exit(app.exec_())
