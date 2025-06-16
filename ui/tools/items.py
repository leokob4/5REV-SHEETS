import sys
import os
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QHBoxLayout, QMessageBox, QHeaderView, QLabel, QComboBox
from PyQt5.QtCore import Qt

# Define o nome do arquivo Excel padrão para esta ferramenta
DEFAULT_DATA_EXCEL_FILENAME = "estoque.xlsx"
DEFAULT_SHEET_NAME = "inventory" # Nome da planilha padrão alterado para "inventory"

# Novos cabeçalhos fornecidos pelo usuário para estoque.xlsx
# Estes serão usados como padrão se a planilha estiver vazia ou for nova.
ITEMS_HEADERS = [
    "part_number", "id_movimentacao", "data_movimentacao", "id_item",
    "tipo_movimentacao", "quantidade_movimentada", "deposito_origem",
    "deposito_destino", "lote_item", "validade_lote",
    "custo_unitario_movimentacao", "referencia_documento",
    "responsavel_movimentacao", "saldo_final_deposito", "motivo_ajuste",
    "status_inspecao_recebimento", "posicao_estoque_fisica",
    "reserva_para_ordem_producao", "reserva_para_pedido_venda",
    "estoque_em_transito", "estoque_disponivel_para_venda"
]

class ItemsTool(QWidget):
    """
    GUI para gerenciar movimentações de estoque.
    Permite visualizar, adicionar e salvar informações de estoque em 'estoque.xlsx'.
    Os cabeçalhos da tabela são dinamicamente carregados do arquivo Excel.
    Pode operar em modo somente leitura se o arquivo for 'engenharia.xlsx'.
    """
    def __init__(self, file_path=None, read_only=False): # Added read_only parameter
        super().__init__()
        if file_path:
            self.file_path = file_path
        else:
            project_root = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
            self.file_path = os.path.join(project_root, 'user_sheets', DEFAULT_DATA_EXCEL_FILENAME)

        # Força somente leitura se o arquivo passado for especificamente 'engenharia.xlsx'
        self.is_read_only = read_only or (os.path.basename(self.file_path) == "engenharia.xlsx") 

        self.setWindowTitle(f"Movimentações de Estoque: {os.path.basename(self.file_path)}")
        if self.is_read_only:
            self.setWindowTitle(self.windowTitle() + " (Somente Leitura)")

        self.layout = QVBoxLayout(self)

        header_layout = QHBoxLayout()
        self.file_name_label = QLabel(f"<b>Arquivo:</b> {os.path.basename(self.file_path)}")
        if self.is_read_only:
            self.file_name_label.setText(self.file_name_label.text() + " (Somente Leitura)")
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
        # Desabilita a edição se estiver em modo somente leitura
        if self.is_read_only:
            self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        else:
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

        # Desabilita botões se estiver em modo somente leitura
        if self.is_read_only:
            self.add_row_btn.setEnabled(False)
            self.save_btn.setEnabled(False)
            QMessageBox.information(self, "Modo Somente Leitura", f"A ferramenta está operando em modo somente leitura para {os.path.basename(self.file_path)}. Edições não são permitidas.")

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
            self.table.setColumnCount(len(ITEMS_HEADERS))
            self.table.setHorizontalHeaderLabels(ITEMS_HEADERS)
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
                
                default_index = self.sheet_selector.findText(DEFAULT_SHEET_NAME)
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
                self.table.setColumnCount(len(ITEMS_HEADERS))
                self.table.setHorizontalHeaderLabels(ITEMS_HEADERS)
                return

            wb = openpyxl.load_workbook(self.file_path)
            if current_sheet_name not in wb.sheetnames:
                QMessageBox.information(self, "Planilha Não Encontrada", f"A planilha '{current_sheet_name}' não foi encontrada em '{os.path.basename(self.file_path)}'. Criando uma nova com cabeçalhos padrão.")
                ws = wb.create_sheet(current_sheet_name)
                ws.append(ITEMS_HEADERS)
                wb.save(self.file_path)
                self._populate_sheet_selector() 
                return

            sheet = wb[current_sheet_name]

            headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
            if not headers:
                headers = ITEMS_HEADERS
            
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)

            data = []
            for row in sheet.iter_rows(min_row=2):
                row_values = [cell.value for cell in row]
                while len(row_values) < len(headers):
                    row_values.append("")
                data.append(row_values)

            self.table.setRowCount(len(data))
            for row_idx, row_data in enumerate(data):
                for col_idx, cell_value in enumerate(row_data):
                    item = QTableWidgetItem(str(cell_value) if cell_value is not None else "")
                    self.table.setItem(row_idx, col_idx, item)

            # Re-aplica a configuração de somente leitura após carregar os dados
            if self.is_read_only:
                self.table.setEditTriggers(QTableWidget.NoEditTriggers)
            else:
                self.table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.AnyKeyPressed)

            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
            self.table.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
            QMessageBox.information(self, "Dados Carregados", f"Dados de '{current_sheet_name}' carregados com sucesso.")

        except Exception as e:
            QMessageBox.critical(self, "Erro de Carregamento", f"Erro ao carregar dados de itens da aba '{current_sheet_name}': {e}")
            self.table.setRowCount(0)
            self.table.setColumnCount(len(ITEMS_HEADERS)) 
            self.table.setHorizontalHeaderLabels(ITEMS_HEADERS)

    def _save_data(self):
        """Salva dados do QTableWidget de volta para a planilha Excel, mantendo cabeçalhos existentes ou usando padrão."""
        if self.is_read_only: # Impede o salvamento em modo somente leitura
            QMessageBox.warning(self, "Ação Não Permitida", "Esta ferramenta está em modo somente leitura. Não é possível salvar alterações.")
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
                    headers_to_save = ITEMS_HEADERS
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
                        headers_to_save = ITEMS_HEADERS
                    ws.append(headers_to_save)
                    wb.save(self.file_path)
                    QMessageBox.information(self, "Planilha Criada", f"Nova planilha '{current_sheet_name}' criada em '{os.path.basename(self.file_path)}'.")
                    self._populate_sheet_selector()

            sheet = wb[current_sheet_name]
            
            for row_idx in range(sheet.max_row, 1, -1):
                sheet.delete_rows(row_idx)

            current_headers = [self.table.horizontalHeaderItem(col).text() for col in range(self.table.columnCount())]
            if not current_headers:
                current_headers = ITEMS_HEADERS
            
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
            QMessageBox.critical(self, "Erro ao Salvar", f"Erro ao salvar dados de itens: {e}")

    def _add_empty_row(self):
        """Adiciona uma linha vazia ao QTableWidget para nova entrada de dados."""
        if self.is_read_only: # Impede a adição de linhas em modo somente leitura
            QMessageBox.warning(self, "Ação Não Permitida", "Esta ferramenta está em modo somente leitura. Não é possível adicionar linhas.")
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
    
    # Criar um workbook vazio se não existir para teste
    if not os.path.exists(test_file_path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = DEFAULT_SHEET_NAME
        ws.append(ITEMS_HEADERS) # Add headers to the newly created sheet
        # A linha de dados de exemplo foi removida aqui, conforme solicitado
        wb.save(test_file_path)

    window = ItemsTool(file_path=test_file_path)
    window.show()
    sys.exit(app.exec_())
