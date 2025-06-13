import os
import openpyxl
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QHBoxLayout, QMessageBox, QHeaderView
from PyQt5.QtCore import Qt

# Define the path to the Excel file for this module
DATA_EXCEL_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'user_sheets', 'estoque.xlsx')
SHEET_NAME = "inventory" # Name of the sheet within estoque.xlsx

class EstoqueTool(QWidget):
    """
    GUI for managing Inventory.
    Allows viewing, adding, and saving inventory data to 'estoque.xlsx'.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerenciador de Estoque")
        self.layout = QVBoxLayout(self)

        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.AnyKeyPressed)
        self.table.setAlternatingRowColors(True)
        self.layout.addWidget(self.table)

        # Control buttons
        button_layout = QHBoxLayout()
        self.add_row_btn = QPushButton("Adicionar Linha")
        self.add_row_btn.clicked.connect(self._add_empty_row)
        self.save_btn = QPushButton("Salvar Dados")
        self.save_btn.clicked.connect(self._save_data)
        self.refresh_btn = QPushButton("Atualizar Dados")
        self.refresh_btn.clicked.connect(self._load_data)

        button_layout.addWidget(self.add_row_btn)
        button_layout.addWidget(self.save_btn)
        button_layout.addWidget(self.refresh_btn)
        self.layout.addLayout(button_layout)

        self._load_data() # Load data on initialization

    def _load_data(self):
        """Loads data from the Excel sheet into the QTableWidget."""
        try:
            if not os.path.exists(DATA_EXCEL_PATH):
                QMessageBox.warning(self, "Arquivo Não Encontrado", f"O arquivo de estoque não foi encontrado: {DATA_EXCEL_PATH}. Criando um novo.")
                self._create_new_excel_file()
                return

            wb = openpyxl.load_workbook(DATA_EXCEL_PATH)
            if SHEET_NAME not in wb.sheetnames:
                QMessageBox.warning(self, "Planilha Não Encontrada", f"A planilha '{SHEET_NAME}' não foi encontrada em '{DATA_EXCEL_PATH}'. Criando uma nova.")
                self._create_new_excel_sheet(wb)
                return

            sheet = wb[SHEET_NAME]
            headers = [cell.value for cell in sheet[1]]
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)

            data = []
            for row in sheet.iter_rows(min_row=2):
                data.append([cell.value for cell in row])

            self.table.setRowCount(len(data))
            for row_idx, row_data in enumerate(data):
                for col_idx, cell_value in enumerate(row_data):
                    item = QTableWidgetItem(str(cell_value) if cell_value is not None else "")
                    self.table.setItem(row_idx, col_idx, item)

            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            QMessageBox.information(self, "Dados Carregados", f"Dados de '{SHEET_NAME}' carregados com sucesso.")

        except Exception as e:
            QMessageBox.critical(self, "Erro de Carregamento", f"Erro ao carregar dados de estoque: {e}")
            self.table.setRowCount(0)
            self.table.setColumnCount(0)

    def _save_data(self):
        """Saves data from the QTableWidget back to the Excel sheet."""
        try:
            if not os.path.exists(DATA_EXCEL_PATH):
                self._create_new_excel_file()

            wb = openpyxl.load_workbook(DATA_EXCEL_PATH)
            if SHEET_NAME not in wb.sheetnames:
                self._create_new_excel_sheet(wb)

            sheet = wb[SHEET_NAME]
            for row_idx in range(sheet.max_row, 1, -1):
                sheet.delete_rows(row_idx)

            current_headers = [self.table.horizontalHeaderItem(col).text() for col in range(self.table.columnCount())]
            if not current_headers:
                current_headers = ["ID Estoque", "Item", "Localização", "Quantidade", "Data Última Atualização"]
                self.table.setColumnCount(len(current_headers))
                self.table.setHorizontalHeaderLabels(current_headers)

            if not sheet[1][0].value:
                sheet.append(current_headers)

            for row_idx in range(self.table.rowCount()):
                row_data = []
                for col_idx in range(self.table.columnCount()):
                    item = self.table.item(row_idx, col_idx)
                    row_data.append(item.text() if item is not None else "")
                sheet.append(row_data)

            wb.save(DATA_EXCEL_PATH)
            QMessageBox.information(self, "Dados Salvos", f"Dados de '{SHEET_NAME}' salvos com sucesso em '{DATA_EXCEL_PATH}'.")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Salvar", f"Erro ao salvar dados de estoque: {e}")

    def _add_empty_row(self):
        """Adds an empty row to the QTableWidget for new data entry."""
        row_count = self.table.rowCount()
        self.table.insertRow(row_count)
        for col_idx in range(self.table.columnCount()):
            self.table.setItem(row_count, col_idx, QTableWidgetItem(""))

    def _create_new_excel_file(self):
        """Creates a new Excel workbook with the specified sheet and headers."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        headers = ["ID Estoque", "Item", "Localização", "Quantidade", "Data Última Atualização"]
        ws.append(headers)
        wb.save(DATA_EXCEL_PATH)
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(0)
        QMessageBox.information(self, "Arquivo Criado", f"Novo arquivo '{DATA_EXCEL_PATH}' com planilha '{SHEET_NAME}' criado.")

    def _create_new_excel_sheet(self, wb):
        """Creates a new sheet within an existing workbook."""
        ws = wb.create_sheet(SHEET_NAME)
        headers = ["ID Estoque", "Item", "Localização", "Quantidade", "Data Última Atualização"]
        ws.append(headers)
        wb.save(DATA_EXCEL_PATH)
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(0)
        QMessageBox.information(self, "Planilha Criada", f"Nova planilha '{SHEET_NAME}' criada em '{DATA_EXCEL_PATH}'.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EstoqueTool()
    window.show()
    sys.exit(app.exec_())
