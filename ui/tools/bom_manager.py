import sys
import os
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QHBoxLayout, QMessageBox, QHeaderView, QLabel, QComboBox
from PyQt5.QtCore import Qt

DEFAULT_DATA_EXCEL_FILENAME = "bom_data.xlsx"
DEFAULT_SHEET_NAME = "BOM"

BOM_HEADERS = ["ID do BOM", "ID do Componente", "Nome do Componente", "Quantidade", "Unidade", "Ref Designator"]
# New headers from engenharia.xlsx mapping
ENGENHARIA_BOM_MAP = {
    "part_number": "ID do Componente",
    "parent_part_number": "ID do BOM",
    "quantidade": "Quantidade",
    "materia_prima": "Tipo (Matéria Prima)" # Custom column for BOM view
}

class BomManagerTool(QWidget):
    """
    GUI para gerenciar Listas de Materiais (BOMs).
    Permite visualizar, adicionar e salvar informações de BOM.
    Os cabeçalhos da tabela são dinamicamente carregados do arquivo Excel.
    Pode mapear dados de engenharia.xlsx para a visualização BOM.
    """
    def __init__(self, file_path=None, sheet_name=None):
        super().__init__()
        if file_path:
            self.file_path = file_path
        else:
            project_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
            self.file_path = os.path.join(project_root, 'user_sheets', DEFAULT_DATA_EXCEL_FILENAME)
        
        self.sheet_name = sheet_name if sheet_name else DEFAULT_SHEET_NAME

        self.is_engenharia_file = (os.path.basename(self.file_path) == "engenharia.xlsx")

        self.setWindowTitle(f"Gerenciador de BOM: {os.path.basename(self.file_path)}")
        if self.is_engenharia_file:
            self.setWindowTitle(self.windowTitle() + " (Dados de Engenharia)")

        self.layout = QVBoxLayout(self)

        header_layout = QHBoxLayout()
        self.file_name_label = QLabel(f"<b>Arquivo:</b> {os.path.basename(self.file_path)}")
        if self.is_engenharia_file:
            self.file_name_label.setText(self.file_name_label.text() + " (Mapeado para BOM)")
        header_layout.addWidget(self.file_name_label)
        header_layout.addStretch()

        header_layout.addWidget(QLabel("Planilha:"))
        self.sheet_selector = QComboBox()
        self.sheet_selector.setMinimumWidth(150)
        self.sheet_selector.currentIndexChanged.connect(self._load_data_from_selected_sheet)
        header_layout.addWidget(self.sheet_selector)

        self.refresh_sheets_btn = QPushButton("Atualizar Abas")
        self.refresh_sheets_btn.clicked.connect(self._populate_sheet_selector)
        header_layout.addWidget(self.refresh_sheets_btn)
        self.layout.addLayout(header_layout)

        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.AnyKeyPressed)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.layout.addWidget(self.table)

        button_layout = QHBoxLayout()
        self.add_row_btn = QPushButton("Adicionar Linha")
        self.add_row_btn.clicked.connect(self._add_empty_row)
        self.save_btn = QPushButton("Salvar Dados")
        self.save_btn.clicked.connect(self._save_data)
        self.refresh_btn = QPushButton("Recarregar Dados da Aba Atual")
        self.refresh_btn.clicked.connect(self._load_data_from_selected_sheet)

        button_layout.addWidget(self.add_row_btn)
        button_layout.addWidget(self.save_btn)
        button_layout.addWidget(self.refresh_btn)
        self.layout.addLayout(button_layout)

        # Disable editing if opening engenharia.xlsx, as BOM Manager is for viewing that as BOM, not editing its raw data
        if self.is_engenharia_file:
            self.add_row_btn.setEnabled(False)
            self.save_btn.setEnabled(False)
            self.table.setEditTriggers(QTableWidget.NoEditTriggers)
            QMessageBox.information(self, "Modo Somente Leitura", "Ao visualizar dados de engenharia como BOM, esta ferramenta opera em modo somente leitura. Para editar, use a ferramenta 'Engenharia (Dados)'.")


        self._populate_sheet_selector()

    def _populate_sheet_selector(self):
        """Popula o QComboBox com os nomes das planilhas do arquivo Excel."""
        self.sheet_selector.clear()
        user_sheets_dir = os.path.dirname(self.file_path)
        os.makedirs(user_sheets_dir, exist_ok=True)

        if not os.path.exists(self.file_path):
            QMessageBox.warning(self, "Arquivo Não Encontrado", f"O arquivo de dados não foi encontrado: {os.path.basename(self.file_path)}. Ele será criado com a aba padrão '{DEFAULT_SHEET_NAME}' ao salvar.")
            self.sheet_selector.addItem(DEFAULT_SHEET_NAME)
            self.table.setRowCount(0)
            self.table.setColumnCount(len(BOM_HEADERS))
            self.table.setHorizontalHeaderLabels(BOM_HEADERS)
            return

        try:
            wb = openpyxl.load_workbook(self.file_path, read_only=True)
            sheet_names = wb.sheetnames
            
            if not sheet_names:
                self.sheet_selector.addItem(DEFAULT_SHEET_NAME)
                QMessageBox.warning(self, "Nenhuma Planilha Encontrada", f"Nenhuma planilha encontrada em '{os.path.basename(self.file_path)}'. Adicionando a aba padrão '{DEFAULT_SHEET_NAME}'.")
            else:
                for sheet_name in sheet_names:
                    self.sheet_selector.addItem(sheet_name)
                
                # Try to set the sheet passed in constructor, else default
                default_index = self.sheet_selector.findText(self.sheet_name)
                if default_index != -1:
                    self.sheet_selector.setCurrentIndex(default_index)
                elif sheet_names:
                    self.sheet_selector.setCurrentIndex(0)
                else:
                    self.sheet_selector.setCurrentIndex(0)

            self._load_data_from_selected_sheet()

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Listar Planilhas", f"Erro ao listar planilhas em '{os.path.basename(self.file_path)}': {e}")
            self.sheet_selector.addItem(DEFAULT_SHEET_NAME)
            self.table.setRowCount(0)
            self.table.setColumnCount(0)

    def _load_data_from_selected_sheet(self):
        """Carrega dados da planilha Excel atualmente selecionada para o QTableWidget, usando cabeçalhos reais."""
        current_sheet_name = self.sheet_selector.currentText()
        if not current_sheet_name or not self.file_path:
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        try:
            wb = None
            if not os.path.exists(self.file_path):
                self.table.setRowCount(0)
                self.table.setColumnCount(len(BOM_HEADERS))
                self.table.setHorizontalHeaderLabels(BOM_HEADERS)
                return

            wb = openpyxl.load_workbook(self.file_path)
            if current_sheet_name not in wb.sheetnames:
                QMessageBox.information(self, "Planilha Não Encontrada", f"A planilha '{current_sheet_name}' não foi encontrada em '{os.path.basename(self.file_path)}'. Criando uma nova com cabeçalhos padrão.")
                ws = wb.create_sheet(current_sheet_name)
                ws.append(BOM_HEADERS)
                wb.save(self.file_path)
                self._populate_sheet_selector() 
                return

            sheet = wb[current_sheet_name]

            # Determine headers to use for the table
            source_headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
            display_headers = []
            data_to_display = []

            if self.is_engenharia_file:
                # Mapeia cabeçalhos de engenharia.xlsx para cabeçalhos BOM
                # Cria uma lista de cabeçalhos de exibição com base no mapeamento e na ordem de BOM_HEADERS
                # Se um cabeçalho BOM_HEADERS não estiver no mapeamento, ele será adicionado como está.
                display_headers = list(BOM_HEADERS) # Start with standard BOM headers
                
                # Create a reverse map for easy lookup of original column indices
                source_header_idx_map = {h: idx for idx, h in enumerate(source_headers)}

                # Prepare the data by mapping engenharia.xlsx columns to BOM view
                for row_idx in range(2, sheet.max_row + 1):
                    row_values = [cell.value for cell in sheet[row_idx]]
                    mapped_row = [""] * len(BOM_HEADERS) # Initialize with empty strings

                    # Mapeia part_number para "ID do Componente"
                    pn_idx = source_header_idx_map.get("part_number")
                    if pn_idx is not None and pn_idx < len(row_values):
                        mapped_row[BOM_HEADERS.index("ID do Componente")] = row_values[pn_idx]

                    # Mapeia parent_part_number para "ID do BOM"
                    ppn_idx = source_header_idx_map.get("parent_part_number")
                    if ppn_idx is not None and ppn_idx < len(row_values):
                        mapped_row[BOM_HEADERS.index("ID do BOM")] = row_values[ppn_idx]
                    
                    # Mapeia quantidade para "Quantidade"
                    qty_idx = source_header_idx_map.get("quantidade")
                    if qty_idx is not None and qty_idx < len(row_values):
                        mapped_row[BOM_HEADERS.index("Quantidade")] = row_values[qty_idx]

                    # Mapeia materia_prima para uma coluna "Tipo (Matéria Prima)" ou similar
                    mp_idx = source_header_idx_map.get("materia_prima")
                    if mp_idx is not None and mp_idx < len(row_values):
                        # Add a new column if "Tipo (Matéria Prima)" is not in BOM_HEADERS
                        if "Tipo (Matéria Prima)" not in display_headers:
                            display_headers.append("Tipo (Matéria Prima)")
                            # Expand mapped_row to accommodate the new column for already processed rows
                            for existing_row in data_to_display:
                                existing_row.append("")
                        
                        mapped_row_idx = display_headers.index("Tipo (Matéria Prima)")
                        if mapped_row_idx < len(mapped_row): # Ensure we don't go out of bounds if headers are changing mid-loop
                            mapped_row[mapped_row_idx] = row_values[mp_idx]
                        else: # Extend if necessary
                            mapped_row.append(row_values[mp_idx])

                    # Include any other original headers not explicitly mapped
                    for h_idx, h_name in enumerate(source_headers):
                        if h_name not in ENGENHARIA_BOM_MAP:
                            if h_name not in display_headers:
                                display_headers.append(h_name)
                                for existing_row in data_to_display:
                                    existing_row.append("") # Pad existing rows
                            
                            # Find position in display_headers
                            display_idx = display_headers.index(h_name)
                            # Ensure mapped_row has enough elements
                            while len(mapped_row) <= display_idx:
                                mapped_row.append("")
                            if h_idx < len(row_values):
                                mapped_row[display_idx] = row_values[h_idx]

                    # Adjust mapped_row length to match current display_headers length
                    while len(mapped_row) < len(display_headers):
                        mapped_row.append("")

                    data_to_display.append(mapped_row)

                # Finalize display headers, ensuring standard BOM headers come first
                final_display_headers = [h for h in BOM_HEADERS if h in display_headers] + \
                                        [h for h in display_headers if h not in BOM_HEADERS]
                display_headers = final_display_headers

            else:
                # Default behavior for non-engenharia.xlsx files
                if not source_headers:
                    display_headers = BOM_HEADERS
                else:
                    display_headers = source_headers
                
                for row in sheet.iter_rows(min_row=2):
                    row_values = [cell.value for cell in row]
                    while len(row_values) < len(display_headers):
                        row_values.append("")
                    data_to_display.append(row_values)
            
            self.table.setColumnCount(len(display_headers))
            self.table.setHorizontalHeaderLabels(display_headers)
            self.table.setRowCount(len(data_to_display))
            for row_idx, row_data in enumerate(data_to_display):
                for col_idx, cell_value in enumerate(row_data):
                    item = QTableWidgetItem(str(cell_value) if cell_value is not None else "")
                    self.table.setItem(row_idx, col_idx, item)

            # Re-apply read-only status based on file type
            if self.is_engenharia_file:
                self.table.setEditTriggers(QTableWidget.NoEditTriggers)
            else:
                self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.AnyKeyPressed)

            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
            self.table.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
            QMessageBox.information(self, "Dados Carregados", f"Dados de '{current_sheet_name}' carregados com sucesso.")

        except Exception as e:
            QMessageBox.critical(self, "Erro de Carregamento", f"Erro ao carregar dados do BOM da aba '{current_sheet_name}': {e}")
            self.table.setRowCount(0)
            self.table.setColumnCount(len(BOM_HEADERS)) 
            self.table.setHorizontalHeaderLabels(BOM_HEADERS)

    def _save_data(self):
        """Salva dados do QTableWidget de volta para a planilha Excel, mantendo cabeçalhos existentes ou usando padrão."""
        if self.is_engenharia_file: # Prevent saving if it's the engenharia.xlsx file being viewed as BOM
            QMessageBox.warning(self, "Ação Não Permitida", "Esta ferramenta está visualizando dados de engenharia em modo somente leitura. Não é possível salvar alterações aqui.")
            return

        if not self.file_path:
            QMessageBox.critical(self, "Erro", "Nenhum arquivo especificado para salvar.")
            return

        current_sheet_name = self.sheet_selector.currentText()
        if not current_sheet_name:
            QMessageBox.warning(self, "Nome da Planilha Inválido", "O nome da planilha não pode estar vazio. Por favor, selecione ou adicione uma aba.")
            return

        try:
            wb = None
            if not os.path.exists(self.file_path):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = current_sheet_name
                
                headers_to_save = [self.table.horizontalHeaderItem(col).text() for col in range(self.table.columnCount())]
                if not headers_to_save:
                    headers_to_save = BOM_HEADERS
                ws.append(headers_to_save)
                
                wb.save(self.file_path)
                QMessageBox.information(self, "Arquivo e Planilha Criados", f"Novo arquivo '{os.path.basename(self.file_path)}' com planilha '{current_sheet_name}' criado.")
                self._populate_sheet_selector() 
            else:
                wb = openpyxl.load_workbook(self.file_path)
                if current_sheet_name not in wb.sheetnames:
                    ws = wb.create_sheet(current_sheet_name)
                    headers_to_save = [self.table.horizontalHeaderItem(col).text() for col in range(self.table.columnCount())]
                    if not headers_to_save:
                        headers_to_save = BOM_HEADERS
                    ws.append(headers_to_save)
                    wb.save(self.file_path)
                    QMessageBox.information(self, "Planilha Criada", f"Nova planilha '{current_sheet_name}' criada em '{os.path.basename(self.file_path)}'.")
                    self._populate_sheet_selector()

            sheet = wb[current_sheet_name]
            
            for row_idx in range(sheet.max_row, 1, -1):
                sheet.delete_rows(row_idx)

            current_headers = [self.table.horizontalHeaderItem(col).text() for col in range(self.table.columnCount())]
            if not current_headers:
                current_headers = BOM_HEADERS
            
            existing_sheet_headers = [cell.value for cell in sheet[1]]
            if existing_sheet_headers != current_headers:
                sheet.delete_rows(1)
                sheet.insert_rows(1)
                sheet.append(current_headers)
            elif not existing_sheet_headers and current_headers:
                sheet.append(current_headers)
            
            for row_idx in range(self.table.rowCount()):
                row_data = []
                for col_idx in range(self.table.columnCount()):
                    item = self.table.item(row_idx, col_idx)
                    row_data.append(item.text() if item is not None else "")
                sheet.append(row_data)

            wb.save(self.file_path)
            QMessageBox.information(self, "Dados Salvos", f"Dados de '{current_sheet_name}' salvos com sucesso em '{os.path.basename(self.file_path)}'.")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Salvar", f"Erro ao salvar dados do BOM: {e}")

    def _add_empty_row(self):
        """Adiciona uma linha vazia ao QTableWidget para nova entrada de dados."""
        if self.is_engenharia_file: # Prevent adding rows if it's the engenharia.xlsx file being viewed as BOM
            QMessageBox.warning(self, "Ação Não Permitida", "Esta ferramenta está visualizando dados de engenharia em modo somente leitura. Não é possível adicionar linhas.")
            return

        row_count = self.table.rowCount()
        self.table.insertRow(row_count)
        for col_idx in range(self.table.columnCount()):
            self.table.setItem(row_count, col_idx, QTableWidgetItem(""))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    project_root_test = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
    test_file_dir = os.path.join(project_root_test, 'user_sheets')
    os.makedirs(test_file_dir, exist_ok=True)
    test_file_path = os.path.join(test_file_dir, DEFAULT_DATA_EXCEL_FILENAME)
    
    if not os.path.exists(test_file_path):
        wb = openpyxl.Workbook()
        wb.save(test_file_path)

    window = BomManagerTool(file_path=test_file_path)
    window.show()
    sys.exit(app.exec_())
