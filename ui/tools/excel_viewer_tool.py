import sys
import os
import openpyxl
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, QHBoxLayout, QMessageBox, QHeaderView, QLabel, QComboBox
from PyQt5.QtCore import Qt

# Nenhuma lista de cabeçalhos default, pois o visualizador lê diretamente do arquivo.
# Nenhuma necessidade de DEFAULT_SHEET_NAME pois ele apenas mostra o que existe.

class ExcelViewerTool(QWidget):
    """
    GUI para visualizar qualquer arquivo Excel (.xlsx).
    Permite alternar entre as planilhas e visualizar seu conteúdo.
    Não permite edição. Os cabeçalhos são lidos da primeira linha de cada planilha.
    """
    def __init__(self, file_path): # Removido read_only do construtor, será sempre True
        super().__init__()
        self.file_path = file_path
        self.setWindowTitle(f"Visualizador Excel: {os.path.basename(self.file_path)}")
        self.layout = QVBoxLayout(self)

        header_layout = QHBoxLayout()
        self.file_name_label = QLabel(f"<b>Arquivo:</b> {os.path.basename(self.file_path)}")
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
        self.table.setEditTriggers(QTableWidget.NoEditTriggers) # O visualizador NÃO permite edição
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.layout.addWidget(self.table)
        
        self._populate_sheet_selector() # Inicia carregando as planilhas

    def _populate_sheet_selector(self):
        """Popula o QComboBox com os nomes das planilhas do arquivo Excel."""
        self.sheet_selector.clear()
        
        # O visualizador não cria o arquivo se ele não existir, apenas avisa.
        if not os.path.exists(self.file_path):
            QMessageBox.warning(self, "Arquivo Não Encontrado", f"O arquivo '{os.path.basename(self.file_path)}' não foi encontrado.")
            # Limpa a tabela se o arquivo não existe
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        try:
            # Sempre abre em modo de leitura para um visualizador
            wb = openpyxl.load_workbook(self.file_path, read_only=True)
            sheet_names = wb.sheetnames
            
            if not sheet_names:
                QMessageBox.warning(self, "Nenhuma Planilha Encontrada", f"Nenhuma planilha encontrada em '{os.path.basename(self.file_path)}'.")
                self.table.setRowCount(0)
                self.table.setColumnCount(0)
            else:
                for sheet_name in sheet_names:
                    self.sheet_selector.addItem(sheet_name)
                
                # Seleciona a primeira sheet por padrão se houver alguma
                if sheet_names:
                    self.sheet_selector.setCurrentIndex(0)
                
            self._load_data_from_selected_sheet() # Carrega os dados da planilha selecionada

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Listar Planilhas", f"Erro ao listar planilhas em '{os.path.basename(self.file_path)}': {e}")
            self.table.setRowCount(0)
            self.table.setColumnCount(0) # Limpa a tabela em caso de erro

    def _load_data_from_selected_sheet(self):
        """Carrega dados da planilha Excel atualmente selecionada para o QTableWidget."""
        current_sheet_name = self.sheet_selector.currentText()
        # Verifica se há uma planilha selecionada e se o arquivo existe
        if not current_sheet_name or not self.file_path or not os.path.exists(self.file_path):
            self.table.setRowCount(0)
            self.table.setColumnCount(0)
            return

        try:
            # Abre em modo de leitura
            wb = openpyxl.load_workbook(self.file_path, read_only=True)
            
            if current_sheet_name not in wb.sheetnames:
                QMessageBox.warning(self, "Planilha Não Encontrada", 
                                    f"A planilha '{current_sheet_name}' não foi encontrada em '{os.path.basename(self.file_path)}'.")
                self.table.setRowCount(0)
                self.table.setColumnCount(0)
                return

            sheet = wb[current_sheet_name]

            # Carrega cabeçalhos da primeira linha da planilha. Se não houver, assume 0 colunas.
            headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
            self.table.setColumnCount(len(headers))
            self.table.setHorizontalHeaderLabels(headers)

            data = []
            # Itera a partir da segunda linha para os dados
            for row in sheet.iter_rows(min_row=2): 
                row_values = [cell.value for cell in row]
                # Garante que a linha tenha células suficientes para os cabeçalhos
                while len(row_values) < len(headers):
                    row_values.append("")
                data.append(row_values)

            self.table.setRowCount(len(data))
            for row_idx, row_data in enumerate(data):
                for col_idx, cell_value in enumerate(row_data):
                    item = QTableWidgetItem(str(cell_value) if cell_value is not None else "")
                    self.table.setItem(row_idx, col_idx, item)

            self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
            self.table.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
            QMessageBox.information(self, "Dados Carregados", f"Dados de '{current_sheet_name}' carregados com sucesso.")

        except Exception as e:
            QMessageBox.critical(self, "Erro de Carregamento", f"Erro ao carregar dados da aba '{current_sheet_name}': {e}")
            self.table.setRowCount(0)
            self.table.setColumnCount(0) # Limpa a tabela em caso de erro grave

# Exemplo de uso (para testar este módulo individualmente)
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Define o caminho para o diretório user_sheets para o teste
    current_dir_test = os.path.dirname(os.path.abspath(__file__))
    project_root_test = os.path.dirname(os.path.dirname(current_dir_test))
    user_sheets_dir_test = os.path.join(project_root_test, 'user_sheets')
    os.makedirs(user_sheets_dir_test, exist_ok=True)

    # Crie um arquivo de teste para o visualizador Excel
    test_excel_file = os.path.join(user_sheets_dir_test, "test_viewer.xlsx")
    
    # Cria o arquivo de teste com dados e múltiplas sheets, se não existir
    if not os.path.exists(test_excel_file):
        wb_test = openpyxl.Workbook()
        ws1 = wb_test.active
        ws1.title = "DadosPrincipais"
        ws1.append(["ID", "Nome", "Valor"])
        ws1.append([1, "Item A", 100])
        ws1.append([2, "Item B", 200])
        ws1.append([3, "Item C", 150])
        ws2 = wb_test.create_sheet("OutrosDados")
        ws2.append(["Código", "Descrição", "Status"])
        ws2.append(["XYZ", "Teste 123", "Ativo"])
        ws2.append(["ABC", "Exemplo 456", "Inativo"])
        ws3 = wb_test.create_sheet("PlanilhaVazia") # Uma planilha sem dados nem cabeçalhos
        wb_test.save(test_excel_file)
        print(f"Arquivo de teste '{test_excel_file}' criado.")
    else:
        print(f"Arquivo de teste '{test_excel_file}' já existe, usando existente.")


    window = ExcelViewerTool(file_path=test_excel_file)
    window.show()
    sys.exit(app.exec_())
