from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem
import openpyxl
import os

class SheetEditorWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("📄 Sheet Editor")
        layout = QVBoxLayout()

        self.load_btn = QPushButton("📂 Load Excel Sheet")
        self.load_btn.clicked.connect(self.load_excel)

        self.table = QTableWidget()
        layout.addWidget(QLabel("Edit Excel (.xlsx) Files"))
        layout.addWidget(self.load_btn)
        layout.addWidget(self.table)

        self.setLayout(layout)

    def load_excel(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "user_sheets/", "Excel Files (*.xlsx)")
        if not file:
            return

        try:
            wb = openpyxl.load_workbook(file)
            sheet = wb.active
            self.table.setRowCount(sheet.max_row)
            self.table.setColumnCount(sheet.max_column)

            for i, row in enumerate(sheet.iter_rows(values_only=True)):
                for j, val in enumerate(row):
                    self.table.setItem(i, j, QTableWidgetItem(str(val) if val is not None else ""))

            save_btn = QPushButton("💾 Save Changes")
            save_btn.clicked.connect(lambda: self.save_excel(file))
            self.layout().addWidget(save_btn)

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def save_excel(self, filepath):
        try:
            wb = openpyxl.Workbook()
            sheet = wb.active

            for row in range(self.table.rowCount()):
                values = []
                for col in range(self.table.columnCount()):
                    item = self.table.item(row, col)
                    values.append(item.text() if item else "")
                sheet.append(values)

            wb.save(filepath)
            QMessageBox.information(self, "Success", "Excel file saved successfully.")

        except Exception as e:
            QMessageBox.critical(self, "Save Error", str(e))
