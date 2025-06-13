import os
import openpyxl
import sys # Added import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QHBoxLayout, QMessageBox, QHeaderView, QLineEdit, QLabel, QComboBox
from PyQt5.QtCore import Qt

# Define the path to the Excel file for this module
# Navigates up three levels (from ui/tools to project root), then down to user_sheets
DATA_EXCEL_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'user_sheets', 'output.xlsx')
DEFAULT_SHEET_NAME = "product_data" # Default sheet name for this tool

class ProductDataTool(QWidget):
    """
    GUI for managing Product Data.
    Allows viewing, adding, and saving product information to 'output.xlsx'.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Dados do Produto")
        self.layout = QVBoxLayout(self)

        # Sheet name selection (ComboBox) and load button
        sheet_control_layout = QHBoxLayout()
        sheet_control_layout.addWidget(QLabel("Planilha:"))
        self.sheet_selector = QComboBox()
        self.sheet_selector.setMinimumWidth(150) # Give it some space
        self.sheet_selector.currentIndexChanged.connect(self._load_data) # Load data when selection changes
        sheet_control_layout.addWidget(self.sheet_selector)

        # A "Refresh Sheets" button to re-populate the combobox
        self.refresh_sheets_btn = QPushButton("Atualizar Abas")
        self.refresh_sheets_btn.clicked.connect(self._populate_sheet_selector)
        sheet_control_layout.addWidget(self.refresh_sheets_btn)

        self.layout.addLayout(sheet_control_layout)

        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.AnyKeyPressed)
        self.table.setAlternatingRowColors(True) # Make rows alternate colors for readability
        self.layout.addWidget(self.table)

        # Control buttons
        button_layout = QHBoxLayout()
        self.add_row_btn = QPushButton("Adicionar Linha")
        self.add_row_btn.clicked.connect(self._add_empty_row)
        self.save_btn = QPushButton("Salvar Dados")
        self.save_btn.clicked.connect(self._save_data)
        # The refresh_btn now just triggers _load_data, which will use the current selection
        self.refresh_btn = QPushButton("Recarregar Dados da Aba Atual")
        self.refresh_btn.clicked.connect(self._load_data)

        button_layout.addWidget(self.add_row_btn)
        button_layout.addWidget(self.save_btn)
        button_layout.addWidget(self.refresh_btn)
        self.layout.addLayout(button_layout)

        self._populate_sheet_selector() # Populate dropdown and load initial data

    def _populate_sheet_selector(self):
        """Populates the QComboBox with sheet names from the Excel file."""
        self.sheet_selector.clear()
        try:
            if not os.path.exists(DATA_EXCEL_PATH):
                # If file doesn't exist, ensure default sheet name is an option
                self.sheet_selector.addItem(DEFAULT_SHEET_NAME)
                QMessageBox.warning(self, "Arquivo Não Encontrado", f"O arquivo de dados não foi encontrado: {DATA_EXCEL_PATH}. Será criado com a aba padrão '{DEFAULT_SHEET_NAME}' ao salvar.")
                self._load_data() # Try to load data which will create the file
                return

            wb = openpyxl.load_workbook(DATA_EXCEL_PATH, read_only=True)
            sheet_names = wb.sheetnames
            
            if not sheet_names:
                self.sheet_selector.addItem(DEFAULT_SHEET_NAME) # Always offer default
                QMessageBox.warning(self, "Nenhuma Planilha Encontrada", f"Nenhuma planilha encontrada em '{DATA_EXCEL_PATH}'. Adicionando a aba padrão '{DEFAULT_SHEET_NAME}'.")
            else:
                for sheet_name in sheet_names:
                    self.sheet_selector.addItem(sheet_name)
                
                # Set default sheet if it exists, otherwise select the first one
                default_index = self.sheet_selector.findText(DEFAULT_SHEET_NAME)
                if default_index != -1:
                    self.sheet_selector.setCurrentIndex(default_index)
                elif sheet_names:
                    self.sheet_selector.setCurrentIndex(0) # Select the first available sheet
                else: # Only default was added
                    self.sheet_selector.setCurrentIndex(0)

            # Manually trigger _load_data if no initial sheet was selected or if refreshing
            self._load_data()

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Listar Planilhas", f"Erro ao listar planilhas em '{DATA_EXCEL_PATH}': {e}")
            self.sheet_selector.addItem(DEFAULT_SHEET_NAME) # Fallback to default name if error

    def _load_data(self):
        """Loads data from the currently selected Excel sheet into the QTableWidget."""
        current_sheet_name = self.sheet_selector.currentText()
        if not current_sheet_name:
            # If no sheet selected (e.g., initial load and no default, or file doesn't exist yet)
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        try:
            wb = None
            if not os.path.exists(DATA_EXCEL_PATH):
                # If the file doesn't exist, it means we need to create it on save.
                # For now, just clear the table and inform the user.
                self.table.setRowCount(0)
                self.table.setColumnCount(0)
                QMessageBox.information(self, "Arquivo Inexistente", f"O arquivo '{DATA_EXCEL_PATH}' não existe. Ele será criado com a aba '{current_sheet_name}' ao salvar os dados.")
                # Set default headers in table if file is new
                headers = ["ID", "Nome do Produto", "Código", "Revisão", "Descrição"]
                self.table.setColumnCount(len(headers))
                self.table.setHorizontalHeaderLabels(headers)
                return

            wb = openpyxl.load_workbook(DATA_EXCEL_PATH)
            if current_sheet_name not in wb.sheetnames:
                # Sheet doesn't exist, create it if we try to load it.
                QMessageBox.warning(self, "Planilha Não Encontrada", f"A planilha '{current_sheet_name}' não foi encontrada em '{DATA_EXCEL_PATH}'. Criando uma nova.")
                self._create_new_excel_sheet(wb, current_sheet_name) # This also saves the workbook
                # After creating, it will be an empty sheet with headers, so load its (empty) content
                sheet = wb[current_sheet_name]
            else:
                sheet = wb[current_sheet_name]

            # Get headers from the first row
            headers = [cell.value for cell in sheet[1]]
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)

            # Load data from the second row onwards
            data = []
            for row in sheet.iter_rows(min_row=2):
                data.append([cell.value for cell in row])

            self.table.setRowCount(len(data))
            for row_idx, row_data in enumerate(data):
                for col_idx, cell_value in enumerate(row_data):
                    item = QTableWidgetItem(str(cell_value) if cell_value is not None else "")
                    self.table.setItem(row_idx, col_idx, item)

            # Resize columns to fit content
            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            QMessageBox.information(self, "Dados Carregados", f"Dados de '{current_sheet_name}' carregados com sucesso.")

        except Exception as e:
            QMessageBox.critical(self, "Erro de Carregamento", f"Erro ao carregar dados do produto da aba '{current_sheet_name}': {e}")
            # Clear table on error
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            # Re-add default headers in case of error for clarity
            headers = ["ID", "Nome do Produto", "Código", "Revisão", "Descrição"]
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)

    def _save_data(self):
        """Saves data from the QTableWidget back to the Excel sheet."""
        current_sheet_name = self.sheet_selector.currentText()
        if not current_sheet_name:
            QMessageBox.warning(self, "Nome da Planilha Inválido", "O nome da planilha não pode estar vazio. Por favor, selecione ou adicione uma aba.")
            return

        try:
            wb = None
            if not os.path.exists(DATA_EXCEL_PATH):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = current_sheet_name
                # Define default headers if file is new
                headers = ["ID", "Nome do Produto", "Código", "Revisão", "Descrição"]
                ws.append(headers)
                wb.save(DATA_EXCEL_PATH)
                QMessageBox.information(self, "Arquivo e Planilha Criados", f"Novo arquivo '{DATA_EXCEL_PATH}' com planilha '{current_sheet_name}' criado.")
                # Refresh selector after creating file/sheet
                self._populate_sheet_selector() 
                return # Exit as the file/sheet was just created, no data to save yet
            
            wb = openpyxl.load_workbook(DATA_EXCEL_PATH)
            if current_sheet_name not in wb.sheetnames:
                ws = wb.create_sheet(current_sheet_name)
                # Define default headers for new sheet
                headers = ["ID", "Nome do Produto", "Código", "Revisão", "Descrição"]
                ws.append(headers)
                wb.save(DATA_EXCEL_PATH)
                QMessageBox.information(self, "Planilha Criada", f"Nova planilha '{current_sheet_name}' criada em '{DATA_EXCEL_PATH}'.")
                # Refresh selector after creating file/sheet
                self._populate_sheet_selector()
                return # Exit as the sheet was just created, no data to save yet
            
            sheet = wb[current_sheet_name]
            
            # Clear existing data but keep header (row 1)
            # Iterate backwards to avoid issues with shifting rows
            for row_idx in range(sheet.max_row, 1, -1):
                sheet.delete_rows(row_idx)

            # Write current headers from table (in case they were changed in GUI)
            current_headers = [self.table.horizontalHeaderItem(col).text() for col in range(self.table.columnCount())]
            if not current_headers: # If no headers are set in the table (e.g., table was empty on load)
                current_headers = ["ID", "Nome do Produto", "Código", "Revisão", "Descrição"]
                self.table.setColumnCount(len(current_headers))
                self.table.setHorizontalHeaderLabels(current_headers)
            
            # Ensure the first row of the sheet (headers) matches current_headers.
            # If not, update it. This handles cases where the default headers were written
            # but the table's headers might have been manually changed (less common) or loaded from an empty sheet.
            existing_sheet_headers = [cell.value for cell in sheet[1]]
            if existing_sheet_headers != current_headers:
                # Clear and rewrite headers if they don't match exactly
                for col_idx, header_value in enumerate(current_headers):
                    sheet.cell(row=1, column=col_idx + 1, value=header_value)
            
            # Append all rows from the QTableWidget
            for row_idx in range(self.table.rowCount()):
                row_data = []
                for col_idx in range(self.table.columnCount()):
                    item = self.table.item(row_idx, col_idx)
                    row_data.append(item.text() if item is not None else "")
                sheet.append(row_data)

            wb.save(DATA_EXCEL_PATH)
            QMessageBox.information(self, "Dados Salvos", f"Dados de '{current_sheet_name}' salvos com sucesso em '{DATA_EXCEL_PATH}'.")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Salvar", f"Erro ao salvar dados do produto: {e}")

    def _add_empty_row(self):
        """Adds an empty row to the QTableWidget for new data entry."""
        row_count = self.table.rowCount()
        self.table.insertRow(row_count)
        for col_idx in range(self.table.columnCount()):
            self.table.setItem(row_count, col_idx, QTableWidgetItem(""))

    # Renamed helper functions to use dynamic sheet_name and existing workbook (wb)
    def _create_new_excel_file(self, sheet_name):
        """Creates a new Excel workbook with the specified sheet and headers."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name
        headers = ["ID", "Nome do Produto", "Código", "Revisão", "Descrição"]
        ws.append(headers)
        wb.save(DATA_EXCEL_PATH)
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(0) # Start with no data rows
        QMessageBox.information(self, "Arquivo Criado", f"Novo arquivo '{DATA_EXCEL_PATH}' com planilha '{sheet_name}' criado.")
        self._populate_sheet_selector() # Refresh selector after creating

    def _create_new_excel_sheet(self, wb, sheet_name):
        """Creates a new sheet within an existing workbook."""
        # Check if sheet already exists, if so, just switch to it.
        if sheet_name in wb.sheetnames:
            QMessageBox.information(self, "Planilha Existente", f"Planilha '{sheet_name}' já existe. Carregando dados dessa aba.")
            self.sheet_selector.setCurrentText(sheet_name) # Set selector to existing sheet
            return

        ws = wb.create_sheet(sheet_name)
        headers = ["ID", "Nome do Produto", "Código", "Revisão", "Descrição"]
        ws.append(headers)
        wb.save(DATA_EXCEL_PATH)
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(0)
        QMessageBox.information(self, "Planilha Criada", f"Nova planilha '{sheet_name}' criada em '{DATA_EXCEL_PATH}'.")
        self._populate_sheet_selector() # Refresh selector after creating

# Example usage (for testing this single module)
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ProductDataTool()
    window.show()
    sys.exit(app.exec_())
