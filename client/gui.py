import sys
import os
import bcrypt
import openpyxl
import json # Para salvar/carregar estrutura de diagramas complexos
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QAction, QTabWidget, QMenu, QToolButton,
    QWidget, QVBoxLayout, QSplitter, QTreeWidget, QTreeWidgetItem,
    QLabel, QLineEdit, QPushButton, QHBoxLayout, QMessageBox, QGraphicsView,
    QGraphicsScene, QGraphicsRectItem, QGraphicsLineItem, QDialog, QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView, QInputDialog, QComboBox
)
from PyQt5.QtCore import Qt, QPointF, QFileInfo
from PyQt5.QtGui import QBrush, QPen, QColor, QFont # Importar QFont explicitamente para QGraphicsScene.addText

# --- Correção para ModuleNotFoundError: No module named 'ui' ---
# Obtém o caminho absoluto do diretório contendo gui.py
current_dir = os.path.dirname(os.path.abspath(__file__))
# Navega até a raiz do projeto (assumindo gui.py está em client/, e client/ está na raiz do projeto/)
project_root = os.path.dirname(current_dir)
# Adiciona a raiz do projeto ao sys.path para que Python possa encontrar 'ui' e 'user_sheets' etc.
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# --- Importar Módulos das Ferramentas ---
# Garanta que esses arquivos existam em client/ui/tools/
from ui.tools.product_data import ProductDataTool
from ui.tools.bom_manager import BomManagerTool
from ui.tools.configurador import ConfiguradorTool
from ui.tools.colaboradores import ColaboradoresTool
from ui.tools.items import ItemsTool
from ui.tools.manufacturing import ManufacturingTool
from ui.tools.pcp import PcpTool
from ui.tools.estoque import EstoqueTool # Esta ferramenta é para o estoque_data.xlsx, não confundir com items.py usando estoque.xlsx
from ui.tools.financeiro import FinanceiroTool
from ui.tools.pedidos import PedidosTool
from ui.tools.manutencao import ManutencaoTool
from ui.tools.engenharia_data import EngenhariaDataTool 
from ui.tools.excel_viewer_tool import ExcelViewerTool 
from ui.tools.structure_view_tool import StructureViewTool
from ui.tools.rpi_tool import RpiTool 

# --- Configuração dos Caminhos dos Arquivos ---
USER_SHEETS_DIR = os.path.join(project_root, "user_sheets")
APP_SHEETS_DIR = os.path.join(project_root, "app_sheets")

# Caminhos para arquivos Excel gerenciados pelo usuário (na pasta user_sheets)
COLABORADORES_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "colaboradores.xlsx")
CONFIGURADOR_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "configurador.xlsx")
FINANCEIRO_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "financeiro.xlsx")
MANUTENCAO_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "manutencao.xlsx")
OUTPUT_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "output.xlsx") # Usado pela ProductDataTool
PEDIDOS_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "pedidos.xlsx")
PROGRAMACAO_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "programacao.xlsx") # Usado pela PcpTool
RPI_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "RPI.xlsx")
ESTOQUE_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "estoque.xlsx") # Usado pela ItemsTool
DB_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "db.xlsx")
ENGENHARIA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "engenharia.xlsx")
BOM_DATA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "bom_data.xlsx") # Padrão para BomManagerTool (se não for engenharia.xlsx)
ITEMS_DATA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "items_data.xlsx") # Arquivo items_data.xlsx original, se ainda for usado por outra ferramenta
MANUFACTURING_DATA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "manufacturing_data.xlsx")

# Caminhos para arquivos Excel gerenciados pelo aplicativo (na pasta app_sheets)
TOOLS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "tools.xlsx")
MODULES_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "modules.xlsx")
PERMISSIONS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "permissions.xlsx")
ROLES_TOOLS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "roles_tools.xlsx")
USERS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "users.xlsx") # Conteúdo da planilha 'users' no db.xlsx
MAIN_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "main.xlsx") # Assumindo que este arquivo existe ou será criado

# Lista de arquivos protegidos (não podem ser excluídos ou renomeados via GUI)
PROTECTED_FILES = [
    os.path.basename(COLABORADORES_EXCEL_PATH),
    os.path.basename(CONFIGURADOR_EXCEL_PATH),
    os.path.basename(FINANCEIRO_EXCEL_PATH),
    os.path.basename(MANUTENCAO_EXCEL_PATH),
    os.path.basename(OUTPUT_EXCEL_PATH),
    os.path.basename(PEDIDOS_EXCEL_PATH),
    os.path.basename(PROGRAMACAO_EXCEL_PATH),
    os.path.basename(RPI_EXCEL_PATH),
    os.path.basename(ESTOQUE_EXCEL_PATH),
    os.path.basename(DB_EXCEL_PATH), # db.xlsx é protegido
    os.path.basename(ENGENHARIA_EXCEL_PATH),
    os.path.basename(TOOLS_EXCEL_PATH),
    os.path.basename(MODULES_EXCEL_PATH),
    os.path.basename(PERMISSIONS_EXCEL_PATH),
    os.path.basename(ROLES_TOOLS_EXCEL_PATH),
    os.path.basename(USERS_EXCEL_PATH), # Redundante se db.xlsx for protegido, mas mantido para clareza
    os.path.basename(MAIN_EXCEL_PATH),
]

# Garante que os diretórios existam
os.makedirs(USER_SHEETS_DIR, exist_ok=True)
os.makedirs(APP_SHEETS_DIR, exist_ok=True)

# Itens de espaço de trabalho codificados (em um aplicativo real, viriam de um backend/banco de dados)
WORKSPACE_ITEMS = [
    "Demo Project - Rev A",
    "Part-001",
    "Assembly-001",
    "Sample Variant - V1.0",
    "Component-XYZ",
    "Specification-005",
    "Drawing-CAD-001",
    "PROD-001", 
    "ASSY-A", 
    "RAW-MAT-001", 
    "100001" 
]

# === FUNÇÕES AUXILIARES DE PLANILHA ===
def load_users_from_excel():
    """Carrega dados de usuário do arquivo Excel do banco de dados."""
    try:
        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        users_sheet = wb["users"]
        users = {}
        # Iterar a partir da segunda linha para pular os cabeçalhos
        for row in users_sheet.iter_rows(min_row=2):
            # Verifica se a linha tem células suficientes antes de acessar
            if len(row) >= 4:
                users[row[1].value] = {
                    "id": row[0].value,
                    "username": row[1].value,
                    "password_hash": row[2].value,
                    "role": row[3].value
                }
        return users
    except FileNotFoundError:
        QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo do banco de dados não foi encontrado: {DB_EXCEL_PATH}")
        return {}
    except KeyError:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha 'users' não foi encontrada em {DB_EXCEL_PATH}")
        return {}
    except Exception as e:
        QMessageBox.critical(None, "Erro de Carregamento", f"Erro ao carregar usuários: {e}")
        return {}

def register_user(username, password, role="user"):
    """Registra um novo usuário no arquivo Excel do banco de dados."""
    try:
        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        sheet = wb["users"]
        next_id = sheet.max_row # Obtém o próximo número de linha disponível para o ID
        # Garante nome de usuário único
        for row in sheet.iter_rows(min_row=2):
            if row[1].value == username:
                raise ValueError("Nome de usuário já existe.")

        password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        # Adiciona novos dados de usuário à planilha
        sheet.append([next_id, username, password_hash, role])
        wb.save(DB_EXCEL_PATH)
    except FileNotFoundError:
        QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo do banco de dados não foi encontrado em: {DB_EXCEL_PATH}. Não é possível registrar o usuário.")
    except KeyError:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha 'users' não foi encontrada em {DB_EXCEL_PATH}. Não é possível registrar o usuário.")
    except Exception as e:
        QMessageBox.critical(None, "Erro de Registro", f"Ocorreu um erro durante o registro do usuário: {e}")

def load_tools_from_excel():
    """
    Carrega dados da ferramenta do arquivo Excel dedicado às ferramentas.
    """
    tools = {}
    try:
        if not os.path.exists(TOOLS_EXCEL_PATH):
            QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo de ferramentas não foi encontrado em: {TOOLS_EXCEL_PATH}. Por favor, certifique-se de que ele exista.")
            return {}

        wb = openpyxl.load_workbook(TOOLS_EXCEL_PATH)
        sheet = wb["tools"] 
        
        if sheet.max_row < 2:
            QMessageBox.warning(None, "Planilha Vazia", f"A planilha 'tools' em {TOOLS_EXCEL_PATH} parece estar vazia ou conter apenas cabeçalhos.")
            return {}

        for row in sheet.iter_rows(min_row=2):
            if len(row) >= 4 and all(cell.value is not None for cell in row[:4]):
                tools[row[0].value] = {
                    "id": row[0].value,
                    "name": row[1].value,
                    "description": row[2].value,
                    "path": row[3].value
                }
    except KeyError:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha 'tools' não foi encontrada em {TOOLS_EXCEL_PATH}. Por favor, certifique-se de que o nome da planilha seja 'tools'.")
        return {}
    except Exception as e:
        QMessageBox.critical(None, "Erro de Carregamento", f"Erro ao carregar ferramentas: {e}")
        return {}
    return tools


def load_role_permissions():
    """Carrega permissões de papel do arquivo Excel do banco de dados."""
    perms = {}
    try:
        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        sheet = wb["access"] # Assumindo que a planilha 'access' existe em db.xlsx
        perms = {}
        for row in sheet.iter_rows(min_row=2):
            if len(row) >= 2 and row[1].value is not None:
                perms[row[0].value] = row[1].value.split(",") if row[1].value.lower() != "all" else "all"
            else:
                print(f"Aviso: Ignorando linha malformada na planilha 'access': {', '.join(str(c.value) for c in row)}")
        return perms
    except FileNotFoundError:
        QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo do banco de dados não foi encontrado em: {DB_EXCEL_PATH}")
        return {}
    except KeyError:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha 'access' não foi encontrada em {DB_EXCEL_PATH}")
        return {}
    except Exception as e:
        QMessageBox.critical(None, "Erro de Carregamento", f"Erro ao carregar permissões: {e}")
        return {}


# === JANELA DE LOGIN ===
class LoginWindow(QWidget):
    """
    A janela de login para o aplicativo.
    Lida com a autenticação e registro de usuários.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("5revolution Login")
        self.setGeometry(400, 200, 300, 180) 
        self.users = load_users_from_excel() 

        self._init_ui()

    def _init_ui(self):
        """Inicializa os elementos da interface do usuário para a janela de login."""
        layout = QVBoxLayout()

        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Nome de Usuário")
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Senha")
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.returnPressed.connect(self.authenticate) # Conecta a tecla Enter para autenticar

        login_btn = QPushButton("Entrar")
        login_btn.clicked.connect(self.authenticate)

        register_btn = QPushButton("Registrar")
        register_btn.clicked.connect(self.handle_register)

        layout.addWidget(QLabel("Bem-vindo ao 5revolution"))
        layout.addWidget(self.username_input)
        layout.addWidget(self.password_input)

        btns_layout = QHBoxLayout()
        btns_layout.addWidget(login_btn)
        btns_layout.addWidget(register_btn)

        layout.addLayout(btns_layout)
        self.setLayout(layout)

    def authenticate(self):
        """Autentica o usuário com base nas credenciais fornecidas."""
        uname = self.username_input.text().strip()
        pwd = self.password_input.text().strip()

        if not uname or not pwd:
            QMessageBox.warning(self, "Falha no Login", "Nome de usuário e senha não podem estar vazios.")
            return

        user = self.users.get(uname)

        if not user or not bcrypt.checkpw(pwd.encode(), user["password_hash"].encode()):
            QMessageBox.warning(self, "Falha no Login", "Nome de usuário ou senha inválidos.")
            return

        self.main = TeamcenterStyleGUI(user)
        self.main.showMaximized() # Abre em tela cheia
        self.close() 

    def handle_register(self):
        """Lida com o registro de usuário."""
        uname = self.username_input.text().strip()
        pwd = self.password_input.text().strip()

        if not uname or not pwd:
            QMessageBox.warning(self, "Erro de Validação", "Nome de usuário e senha são obrigatórios para o registro.")
            return

        try:
            register_user(uname, pwd)
            QMessageBox.information(self, "Registrado", f"Usuário '{uname}' registrado com sucesso com o papel 'user'.")
            self.users = load_users_from_excel() # Recarrega usuários após o registro
            self.username_input.clear()
            self.password_input.clear()
        except ValueError as ve:
            QMessageBox.warning(self, "Falha no Registro", str(ve))
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro durante o registro: {e}")

# === FERRAMENTA: ENGENHARIA WORKFLOW DIAGRAM ===
class EngenhariaWorkflowTool(QWidget):
    """
    Ferramenta para criar e visualizar diagramas de fluxo de trabalho.
    Pretende ser similar ao software "Dia" para diagramas básicos.
    Permite salvar e carregar dados do diagrama para/de um arquivo Excel.
    """
    DEFAULT_DATA_EXCEL_FILENAME = "engenharia.xlsx"
    DEFAULT_SHEET_NAME = "Workflows" # Planilha padrão para salvar/carregar workflows

    def __init__(self, file_path=None):
        super().__init__()
        self.setWindowTitle("Engenharia (Workflow) Tool")
        
        self.file_path = file_path if file_path else os.path.join(
            os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 
            'user_sheets', self.DEFAULT_DATA_EXCEL_FILENAME
        )
        self.current_sheet_name = self.DEFAULT_SHEET_NAME

        self.layout = QVBoxLayout(self)

        # Controles de arquivo e planilha
        file_sheet_layout = QHBoxLayout()
        file_sheet_layout.addWidget(QLabel(f"<b>Arquivo:</b> {os.path.basename(self.file_path)}"))
        file_sheet_layout.addStretch()
        file_sheet_layout.addWidget(QLabel("Planilha:"))
        self.sheet_selector = QComboBox()
        self.sheet_selector.setMinimumWidth(150)
        self.sheet_selector.currentIndexChanged.connect(self._load_workflow_from_excel)
        file_sheet_layout.addWidget(self.sheet_selector)
        self.refresh_sheets_btn = QPushButton("Atualizar Abas")
        self.refresh_sheets_btn.clicked.connect(self._populate_sheet_selector)
        file_sheet_layout.addWidget(self.refresh_sheets_btn)
        self.layout.addLayout(file_sheet_layout)

        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        self.layout.addWidget(self.view)

        # Botões de controle de diagrama
        control_layout = QHBoxLayout()
        add_node_btn = QPushButton("Adicionar Nó de Tarefa")
        add_node_btn.clicked.connect(self._add_task_node)
        add_link_btn = QPushButton("Adicionar Ligação de Dependência")
        add_link_btn.clicked.connect(self._add_dependency_link)
        clear_btn = QPushButton("Limpar Diagrama")
        clear_btn.clicked.connect(self._clear_diagram)
        save_btn = QPushButton("Salvar Workflow")
        save_btn.clicked.connect(self._save_workflow_to_excel)
        load_btn = QPushButton("Carregar Workflow")
        load_btn.clicked.connect(self._load_workflow_from_excel)

        control_layout.addWidget(add_node_btn)
        control_layout.addWidget(add_link_btn)
        control_layout.addWidget(clear_btn)
        control_layout.addWidget(save_btn)
        control_layout.addWidget(load_btn)
        self.layout.addLayout(control_layout)

        self.nodes = [] # Para rastrear nós adicionados (QGraphicsRectItem)
        self.node_properties = {} # Para armazenar propriedades personalizadas (texto, posição)
        self.links = [] # Para rastrear links adicionados (QGraphicsLineItem)

        self._populate_sheet_selector() # Popula o seletor de abas e carrega o workflow inicial

    def _populate_sheet_selector(self):
        """Popula o QComboBox com os nomes das planilhas do arquivo Excel."""
        self.sheet_selector.clear()
        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)

        if not os.path.exists(self.file_path):
            QMessageBox.warning(self, "Arquivo Não Encontrado", f"O arquivo de dados não foi encontrado: {os.path.basename(self.file_path)}. Ele será criado com a aba padrão '{self.DEFAULT_SHEET_NAME}' ao salvar.")
            self.sheet_selector.addItem(self.DEFAULT_SHEET_NAME)
            self.current_sheet_name = self.DEFAULT_SHEET_NAME
            return

        try:
            wb = openpyxl.load_workbook(self.file_path, read_only=True)
            sheet_names = wb.sheetnames
            
            if not sheet_names:
                self.sheet_selector.addItem(self.DEFAULT_SHEET_NAME)
                QMessageBox.warning(self, "Nenhuma Planilha Encontrada", f"Nenhuma planilha encontrada em '{os.path.basename(self.file_path)}'. Adicionando a aba padrão '{self.DEFAULT_SHEET_NAME}'.")
            else:
                for sheet_name in sheet_names:
                    self.sheet_selector.addItem(sheet_name)
                
                default_index = self.sheet_selector.findText(self.DEFAULT_SHEET_NAME)
                if default_index != -1:
                    self.sheet_selector.setCurrentIndex(default_index)
                elif sheet_names:
                    self.sheet_selector.setCurrentIndex(0) # Seleciona a primeira sheet disponível
                
                self.current_sheet_name = self.sheet_selector.currentText()
            
            self._load_workflow_from_excel() # Carrega o workflow da aba selecionada/padrão

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Listar Planilhas", f"Erro ao listar planilhas em '{os.path.basename(self.file_path)}': {e}")
            self.sheet_selector.addItem(self.DEFAULT_SHEET_NAME) # Fallback
            self.current_sheet_name = self.DEFAULT_SHEET_NAME

    def _save_workflow_to_excel(self):
        """
        Salva o estado atual do diagrama para a planilha Excel selecionada.
        Formato de exemplo no Excel:
        Sheet: "Workflows"
        Colunas: "Tipo", "ID", "X", "Y", "Largura", "Altura", "Texto", "Cor", "Conexões" (JSON de IDs conectados)
        Exemplo de Linha de Nó: ["Node", "node1", 50, 50, 100, 50, "Fase de Design", "#ADD8E6", "[]"]
        Exemplo de Linha de Link: ["Link", "link1", start_node_id, end_node_id, 0, 0, "", "", "[]"]
        """
        current_sheet_name = self.sheet_selector.currentText()
        if not current_sheet_name:
            QMessageBox.warning(self, "Erro", "Selecione uma planilha para salvar.")
            return

        try:
            wb = None
            if not os.path.exists(self.file_path):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = current_sheet_name
            else:
                wb = openpyxl.load_workbook(self.file_path)
                if current_sheet_name not in wb.sheetnames:
                    ws = wb.create_sheet(current_sheet_name)
                else:
                    ws = wb[current_sheet_name]
            
            # Limpa o conteúdo existente, mas mantém o cabeçalho se houver
            if ws.max_row > 1: # Se tiver mais que 1 linha (cabeçalho + dados)
                ws.delete_rows(2, ws.max_row) # Deleta todas as linhas de dados

            # Define os cabeçalhos da planilha para o workflow (se não existirem ou estiverem incorretos)
            workflow_headers = ["Tipo", "ID", "X", "Y", "Largura", "Altura", "Texto", "Cor", "Conexões"]
            current_excel_headers = [cell.value for cell in ws[1]] if ws.max_row > 0 else []

            if current_excel_headers != workflow_headers:
                if ws.max_row > 0: # Se houver cabeçalhos existentes, os apaga
                    ws.delete_rows(1)
                ws.insert_rows(1) # Insere uma nova primeira linha
                ws.append(workflow_headers)
            elif not current_excel_headers: # Se a planilha estiver completamente vazia
                ws.append(workflow_headers)


            # Salvar nós
            for i, node_item in enumerate(self.nodes):
                node_id = f"node_{i}" # Um ID simples para cada nó
                text_item = None
                for item in self.scene.items(node_item.boundingRect()): # Encontra o QGraphicsTextItem associado
                    if isinstance(item, QGraphicsTextItem):
                        text_item = item
                        break
                
                node_text = text_item.toPlainText() if text_item else ""
                node_x = node_item.rect().x()
                node_y = node_item.rect().y()
                node_width = node_item.rect().width()
                node_height = node_item.rect().height()
                # Cores precisam ser convertidas de QBrush para string (hex ou nome)
                node_color = node_item.brush().color().name() # Ex: "#ADD8E6"
                
                # Para as conexões, precisaria de uma maneira de identificar os links conectados a este nó.
                # Por simplicidade, aqui é um stub.
                connections = [] 
                
                ws.append(["Node", node_id, node_x, node_y, node_width, node_height, node_text, node_color, json.dumps(connections)])

            # Salvar links (conceitual, exigiria rastrear nós de origem/destino do link)
            for i, link_item in enumerate(self.links):
                link_id = f"link_{i}"
                # Exemplo: Se você tivesse rastreado os IDs dos nós conectados:
                # source_node_id = link_item.source_node.id (precisaria ser implementado)
                # dest_node_id = link_item.dest_node.id (precisaria ser implementado)
                ws.append(["Link", link_id, "", "", "", "", "", "", json.dumps({"source": "id_origem", "target": "id_destino"})])

            wb.save(self.file_path)
            QMessageBox.information(self, "Sucesso", f"Workflow salvo em '{current_sheet_name}' em '{os.path.basename(self.file_path)}'.")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar o workflow: {e}")

    def _load_workflow_from_excel(self):
        """
        Carrega um diagrama de fluxo de trabalho da planilha Excel selecionada.
        """
        self._clear_diagram() # Limpa o diagrama atual antes de carregar
        current_sheet_name = self.sheet_selector.currentText()
        if not current_sheet_name or not os.path.exists(self.file_path):
            QMessageBox.warning(self, "Erro", "Arquivo ou planilha não selecionados/encontrados para carregar.")
            return

        try:
            wb = openpyxl.load_workbook(self.file_path)
            if current_sheet_name not in wb.sheetnames:
                QMessageBox.warning(self, "Planilha Não Encontrada", f"A planilha '{current_sheet_name}' não foi encontrada em '{os.path.basename(self.file_path)}'.")
                return

            sheet = wb[current_sheet_name]
            
            # Mapeia cabeçalhos para índices para carregamento flexível
            headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
            header_map = {h: idx for idx, h in enumerate(headers)}

            loaded_nodes = {} # Para mapear IDs de nó de volta aos QGraphicsRectItem

            for row_idx in range(2, sheet.max_row + 1): # Começa da segunda linha (dados)
                row_values = [cell.value for cell in sheet[row_idx]]
                row_type = row_values[header_map.get("Tipo")] if "Tipo" in header_map and header_map["Tipo"] < len(row_values) else None

                if row_type == "Node":
                    node_id = row_values[header_map.get("ID")] if "ID" in header_map and header_map["ID"] < len(row_values) else None
                    x = row_values[header_map.get("X")] if "X" in header_map and header_map["X"] < len(row_values) else 0
                    y = row_values[header_map.get("Y")] if "Y" in header_map and header_map["Y"] < len(row_values) else 0
                    width = row_values[header_map.get("Largura")] if "Largura" in header_map and header_map["Largura"] < len(row_values) else 100
                    height = row_values[header_map.get("Altura")] if "Altura" in header_map and header_map["Altura"] < len(row_values) else 50
                    text = row_values[header_map.get("Texto")] if "Texto" in header_map and header_map["Texto"] < len(row_values) else ""
                    color_name = row_values[header_map.get("Cor")] if "Cor" in header_map and header_map["Cor"] < len(row_values) else "lightblue"

                    if node_id:
                        node_rect = self.scene.addRect(x, y, width, height, QPen(Qt.black), QBrush(QColor(color_name)))
                        node_text_item = self.scene.addText(str(text))
                        node_text_item.setPos(x + 5, y + 15) # Ajuste de posição para o texto dentro do nó
                        
                        self.nodes.append(node_rect)
                        loaded_nodes[node_id] = node_rect # Armazena para referências de link

                elif row_type == "Link":
                    # Carregar links: exigiria referenciar nós já criados
                    link_data_str = row_values[header_map.get("Conexões")] if "Conexões" in header_map and header_map["Conexões"] < len(row_values) else "{}"
                    try:
                        link_data = json.loads(link_data_str)
                        source_id = link_data.get("source")
                        target_id = link_data.get("target")

                        source_node = loaded_nodes.get(source_id)
                        target_node = loaded_nodes.get(target_id)

                        if source_node and target_node:
                            # Adicionar a linha (precisaria de lógica para calcular os pontos de conexão exatos)
                            pen = QPen(Qt.darkGray, 2)
                            line = self.scene.addLine(
                                source_node.rect().x() + source_node.rect().width(), source_node.rect().y() + source_node.rect().height() / 2,
                                target_node.rect().x(), target_node.rect().y() + target_node.rect().height() / 2,
                                pen
                            )
                            self.links.append(line)
                    except json.JSONDecodeError:
                        print(f"Aviso: Dados de conexão inválidos para link: {link_data_str}")

            QMessageBox.information(self, "Sucesso", f"Workflow carregado de '{current_sheet_name}' em '{os.path.basename(self.file_path)}'.")

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Carregar", f"Não foi possível carregar o workflow: {e}")

    def _add_sample_diagram_elements(self):
        """Adiciona alguns elementos de exemplo à cena do diagrama ao iniciar."""
        # Task nodes
        node1 = self.scene.addRect(50, 50, 100, 50, QPen(Qt.black), QBrush(QColor("lightblue")))
        node2 = self.scene.addRect(200, 150, 100, 50, QPen(Qt.black), QBrush(QColor("lightgreen")))
        node3 = self.scene.addRect(350, 50, 100, 50, QPen(Qt.black), QBrush(QColor("lightcoral")))

        # Corrigido: QGraphicsScene.addText agora aceita o texto, fonte (opcional) e depois a posição.
        # Estamos usando setPos para maior clareza e flexibilidade.
        text_item1 = self.scene.addText("Fase de Design")
        text_item1.setPos(50 + 5, 50 + 15) # Ajusta posição para dentro do nó

        text_item2 = self.scene.addText("Revisão (Aprovado)")
        text_item2.setPos(200 + 5, 150 + 15)

        text_item3 = self.scene.addText("Preparação da Produção")
        text_item3.setPos(350 + 5, 50 + 15)

        self.nodes.extend([node1, node2, node3]) # Adiciona os nódes à lista

        # Links/Arrows
        pen = QPen(Qt.darkGray)
        pen.setWidth(2)
        link1 = self.scene.addLine(node1.x() + node1.rect().width(), node1.y() + node1.rect().height() / 2,
                           node2.x(), node2.y() + node2.rect().height() / 2, pen)
        link2 = self.scene.addLine(node2.x() + node2.rect().width(), node2.y() + node2.rect().height() / 2,
                           node3.x(), node3.y() + node3.rect().height() / 2, pen)
        self.links.extend([link1, link2])


    def _add_task_node(self):
        """Adiciona um novo nó de tarefa genérico ao diagrama."""
        x = 10 + len(self.nodes) * 120 
        y = 10 + (len(self.nodes) % 3) * 70
        
        node = self.scene.addRect(x, y, 100, 50, QPen(Qt.black), QBrush(QColor("#FFD700"))) # Cor ouro
        
        text_item = self.scene.addText(f"Nova Tarefa {len(self.nodes) + 1}")
        text_item.setPos(x + 5, y + 15) # Posiciona o texto dentro do nó
        
        self.nodes.append(node)
        self.view.centerOn(node)

    def _add_dependency_link(self):
        """
        Prompts user to select two nodes to link. (Conceptual, requires selection logic).
        Para uma implementação completa:
        1. Permitir que o usuário clique no primeiro nó.
        2. Armazenar o primeiro nó.
        3. Permitir que o usuário clique no segundo nó.
        4. Desenhar uma linha entre os centros ou bordas dos dois nós.
        """
        QMessageBox.information(self, "Adicionar Ligação", "Clique em dois nós de tarefa para criar uma ligação. (Lógica de seleção a ser implementada na próxima etapa)")

    def _clear_diagram(self):
        """Limpa todos os elementos do diagrama."""
        self.scene.clear()
        self.nodes = [] 
        self.links = []
        self.node_properties = {}
        QMessageBox.information(self, "Diagrama Limpo", "O diagrama foi limpo.")

# === NOVA FERRAMENTA: ATUALIZADOR DE CABEÇALHOS DO BD ===
class DbHeadersUpdaterTool(QWidget):
    """
    Ferramenta para atualizar a planilha 'db_db' em db.xlsx com cabeçalhos de todos
    os arquivos Excel nas pastas user_sheets e app_sheets.
    """
    # Mapeamento de descrições para a coluna 'descr_variavel'
    DESCRIPTION_MAP = {
        ('user_sheets/RPI.xlsx', 'id_rota'): 'Identificador único para a rota de produção.',
        ('user_sheets/RPI.xlsx', 'part_number'): 'Número de identificação da peça ou item associado à rota.',
        ('user_sheets/RPI.xlsx', 'description'): 'Descrição detalhada da rota ou do item principal da rota.',
        ('user_sheets/RPI.xlsx', 'recurso'): 'Recurso (máquina, posto de trabalho) utilizado na operação da rota.',
        ('user_sheets/RPI.xlsx', 'operacao'): 'Etapa específica do processo de produção dentro da rota.',
        ('user_sheets/RPI.xlsx', 'tempo_ciclo'): 'Tempo necessário para completar um ciclo de produção em uma operação.',
        ('user_sheets/RPI.xlsx', 'quantidade_por_ciclo'): 'Quantidade de unidades produzidas em um ciclo de operação.',
        ('user_sheets/RPI.xlsx', 'observacoes'): 'Campo para notas ou informações adicionais sobre a rota.',
        ('user_sheets/RPI.xlsx', 'deposito_padrao'): 'Depósito principal para matéria-prima ou componentes desta rota.',
        ('user_sheets/RPI.xlsx', 'ferramenta'): 'Nome ou código da ferramenta específica utilizada na operação da rota.',
        ('user_sheets/RPI.xlsx', 'deposito_ferramenta'): 'Depósito onde a ferramenta utilizada nesta rota está armazenada.',
        ('user_sheets/RPI.xlsx', 'endereco_ferramenta'): 'Endereço físico da ferramenta dentro do depósito.',
        ('user_sheets/RPI.xlsx', 'recurso_tipo'): 'Categoria do recurso (ex: Máquina, Posto de Trabalho, Mão de Obra indireta).',
        ('user_sheets/RPI.xlsx', 'operacao_sequencia'): 'Ordem sequencial da operação dentro da rota de produção.',
        ('user_sheets/RPI.xlsx', 'operacao_instrucoes'): 'Instruções detalhadas ou procedimento padrão para a execução da operação.',
        ('user_sheets/RPI.xlsx', 'set_up_time'): 'Tempo de preparação (ajuste, configuração) necessário antes da operação iniciar.',
        ('user_sheets/RPI.xlsx', 'down_time_estimado'): 'Tempo de inatividade estimado para a operação ou recurso.',
        ('user_sheets/RPI.xlsx', 'criterio_qualidade'): 'Padrões ou critérios para inspeção de qualidade nesta operação.',
        ('user_sheets/RPI.xlsx', 'tolerancia_qualidade'): 'Variação aceitável ou limite de tolerância para o critério de qualidade.',
        ('user_sheets/RPI.xlsx', 'necessidade_mao_obra'): 'Número de operadores ou colaboradores necessários para a operação.',
        ('user_sheets/RPI.xlsx', 'habilidade_necessaria'): 'Qualificação ou tipo de habilidade exigida para o operador nesta operação.',
        ('user_sheets/RPI.xlsx', 'custo_hora_recurso'): 'Custo operacional por hora do recurso utilizado na operação.',
        ('user_sheets/RPI.xlsx', 'custo_hora_mao_obra'): 'Custo por hora da mão de obra alocada à operação.',
        ('user_sheets/RPI.xlsx', 'lote_minimo_producao'): 'Tamanho mínimo de lote para esta rota de produção.',
        ('user_sheets/RPI.xlsx', 'versao_rota'): 'Número da versão ou revisão da rota de produção.',
        ('user_sheets/RPI.xlsx', 'data_ultima_revisao_rota'): 'Data da última alteração ou revisão da rota de produção.',
        ('user_sheets/RPI.xlsx', 'responsavel_revisao_rota'): 'Colaborador responsável pela última revisão da rota.',
        ('user_sheets/RPI.xlsx', 'custo_total_rota_estimado'): 'Custo total estimado para completar toda a rota de produção.',
        ('user_sheets/RPI.xlsx', 'tempo_total_rota_estimado'): 'Tempo total estimado para completar toda a rota de produção.',
        ('user_sheets/colaboradores.xlsx', 'id_colab'): 'Identificador único do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'matricula_colab'): 'Número de matrícula do colaborador na empresa.',
        ('user_sheets/colaboradores.xlsx', 'nome_colab'): 'Nome completo do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'data_nasc'): 'Data de nascimento do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'data_contrat'): 'Data de contratação do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'data_disp'): 'Data em que o colaborador estará disponível para novas tarefas/projetos.',
        ('user_sheets/colaboradores.xlsx', 'setor_colab'): 'Setor ou departamento ao qual o colaborador pertence.',
        ('user_sheets/colaboradores.xlsx', 'recurso_colab'): 'Recurso (máquina, ferramenta) principal associado ao colaborador (se aplicável).',
        ('user_sheets/colaboradores.xlsx', 'enabled_colab'): 'Status de habilitação do colaborador no sistema (Ativo/Inativo).',
        ('user_sheets/colaboradores.xlsx', 'cpf'): 'Número de Cadastro de Pessoa Física (CPF) do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'endereco'): 'Endereço residencial completo do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'telefone'): 'Número de telefone para contato do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'email'): 'Endereço de e-mail profissional ou pessoal do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'cargo'): 'Cargo ou função atual do colaborador na empresa.',
        ('user_sheets/colaboradores.xlsx', 'departamento'): 'Departamento principal do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'data_contratacao'): 'Data formal de contratação do colaborador (reafirmação de data_contrat).',
        ('user_sheets/colaboradores.xlsx', 'status_contrato'): 'Status atual do contrato de trabalho (Ativo, Férias, Afastado, Demitido).',
        ('user_sheets/colaboradores.xlsx', 'salario_base'): 'Valor do salário base do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'horas_trabalho_semanais'): 'Número de horas de trabalho semanais previstas para o colaborador.',
        ('user_sheets/colaboradores.xlsx', 'habilidades_principais'): 'Lista de habilidades ou qualificações principais do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'data_ultima_avaliacao'): 'Data da última avaliação de desempenho ou performance do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'supervisor'): 'Nome ou ID do supervisor direto do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'turno_trabalho'): 'Turno de trabalho atribuído ao colaborador.',
        ('user_sheets/colaboradores.xlsx', 'custo_hora_colaborador'): 'Custo por hora da mão de obra do colaborador.',
        ('user_sheets/colaboradores.xlsx', 'motivo_saida'): 'Motivo do desligamento ou saída do colaborador (se aplicável).',
        ('user_sheets/configurador.xlsx', 'id_configurador'): 'Identificador único de um parâmetro ou configuração do sistema.',
        ('user_sheets/configurador.xlsx', 'part_number'): 'Número da peça ou item associado a esta configuração (se aplicável).',
        ('user_sheets/configurador.xlsx', 'description'): 'Descrição geral da configuração ou parâmetro.',
        ('user_sheets/configurador.xlsx', 'config_desc'): 'Descrição específica do valor ou comportamento da configuração.',
        ('user_sheets/configurador.xlsx', 'config_type'): 'Tipo de configuração (ex: numérico, texto, booleano, lista).',
        ('user_sheets/configurador.xlsx', 'ID Config'): 'Identificador da configuração (redundante com id_configurador, se for o mesmo conceito).',
        ('user_sheets/configurador.xlsx', 'Nome Config'): 'Nome ou título descritivo da configuração/parâmetro.',
        ('user_sheets/configurador.xlsx', 'Opção 1'): 'Primeira opção de valor para a configuração (para tipos de lista ou booleanos).',
        ('user_sheets/configurador.xlsx', 'Opção 2'): 'Segunda opção de valor para a configuração.',
        ('user_sheets/configurador.xlsx', 'Item Associado'): 'Item específico ao qual esta configuração se aplica (se não for global).',
        ('user_sheets/db.xlsx', 'id_item'): 'Identificador único para cada item cadastrado (produto, matéria-prima, componente).',
        ('user_sheets/db.xlsx', 'nome_item'): 'Nome comercial ou descritivo do item.',
        ('user_sheets/db.xlsx', 'descricao_detalhada_item'): 'Descrição completa e detalhada das características do item.',
        ('user_sheets/db.xlsx', 'tipo_item'): 'Classificação do item (ex: Matéria-prima, Componente, Produto Acabado, Embalagem).',
        ('user_sheets/db.xlsx', 'unidade_medida_padrao'): 'Unidade de medida padrão para o item (ex: KG, UN, M).',
        ('user_sheets/db.xlsx', 'peso_unitario'): 'Peso de uma única unidade do item.',
        ('user_sheets/db.xlsx', 'volume_unitario'): 'Volume ocupado por uma única unidade do item.',
        ('user_sheets/db.xlsx', 'custo_padrao_unitario'): 'Custo unitário estimado ou padrão para o item.',
        ('user_sheets/db.xlsx', 'custo_medio_unitario'): 'Custo unitário médio ponderado do item em estoque.',
        ('user_sheets/db.xlsx', 'fornecedor_principal'): 'Nome do fornecedor principal para este item.',
        ('user_sheets/db.xlsx', 'marca_item'): 'Marca comercial do item (se aplicável).',
        ('user_sheets/db.xlsx', 'categoria_item'): 'Categoria de classificação do item (ex: Eletrônicos, Metais, Plásticos).',
        ('user_sheets/db.xlsx', 'vida_util_estimada'): 'Vida útil estimada do item em dias ou meses (para perecíveis ou com obsolescência).',
        ('user_sheets/db.xlsx', 'data_cadastro_item'): 'Data em que o item foi cadastrado no sistema.',
        ('user_sheets/db.xlsx', 'status_item'): 'Status atual do item (ex: Ativo, Obsoleto, Descontinuado).',
        ('user_sheets/db.xlsx', 'historico_precos_compra'): 'Referência ou resumo do histórico de preços de compra do item.',
        ('user_sheets/db.xlsx', 'foto_url_item'): 'URL para uma imagem ou foto do item.',
        ('user_sheets/db.xlsx', 'codigo_barras'): 'Código de barras associado ao item para leitura automatizada.',
        ('user_sheets/db.xlsx', 'dados_tecnicos_adicionais'): 'Campo para dados técnicos adicionais em formato livre (JSON, texto).',
        ('user_sheets/db.xlsx', 'prazo_validade_dias'): 'Prazo de validade do item em dias a partir da fabricação/compra.',
        ('user_sheets/estoque.xlsx', 'id_movimentacao'): 'Identificador único de cada transação de movimentação de estoque.',
        ('user_sheets/estoque.xlsx', 'data_movimentacao'): 'Data em que a movimentação de estoque ocorreu.',
        ('user_sheets/estoque.xlsx', 'id_item'): 'Identificador do item que foi movimentado.',
        ('user_sheets/estoque.xlsx', 'tipo_movimentacao'): 'Tipo de transação de estoque (Ex: Entrada por Compra, Saída por Venda, Transferência).',
        ('user_sheets/estoque.xlsx', 'quantidade_movimentada'): 'Quantidade do item que foi movimentada.',
        ('user_sheets/estoque.xlsx', 'deposito_origem'): 'Depósito de onde o item foi movimentado (para saídas e transferências).',
        ('user_sheets/estoque.xlsx', 'deposito_destino'): 'Depósito para onde o item foi movimentado (para entradas e transferências).',
        ('user_sheets/estoque.xlsx', 'lote_item'): 'Número do lote específico do item movimentado (para rastreabilidade).',
        ('user_sheets/estoque.xlsx', 'validade_lote'): 'Data de validade do lote específico do item.',
        ('user_sheets/estoque.xlsx', 'custo_unitario_movimentacao'): 'Custo unitário do item no momento desta movimentação.',
        ('user_sheets/estoque.xlsx', 'referencia_documento'): 'Número ou ID do documento que originou a movimentação (ex: Ordem de Compra, Nota Fiscal).',
        ('user_sheets/estoque.xlsx', 'responsavel_movimentacao'): 'Colaborador responsável pela execução da movimentação de estoque.',
        ('user_sheets/estoque.xlsx', 'saldo_final_deposito'): 'Saldo de estoque do item no depósito após a conclusão da movimentação.',
        ('user_sheets/estoque.xlsx', 'motivo_ajuste'): 'Motivo específico para um ajuste de estoque (ex: inventário, quebra, furto).',
        ('user_sheets/estoque.xlsx', 'status_inspecao_recebimento'): 'Status da inspeção de qualidade no momento do recebimento (Aprovado, Reprovado).',
        ('user_sheets/estoque.xlsx', 'posicao_estoque_fisica'): 'Localização física detalhada do item dentro do depósito (corredor, prateleira).',
        ('user_sheets/estoque.xlsx', 'reserva_para_ordem_producao'): 'ID da Ordem de Produção para a qual este item está reservado.',
        ('user_sheets/estoque.xlsx', 'reserva_para_pedido_venda'): 'ID do Pedido de Venda para o qual este item está reservado.',
        ('user_sheets/estoque.xlsx', 'estoque_em_transito'): 'Quantidade de item que foi expedida mas ainda não foi recebida no destino.',
        ('user_sheets/estoque.xlsx', 'estoque_disponivel_para_venda'): 'Quantidade de item que pode ser vendida ou utilizada imediatamente.',
        ('user_sheets/financeiro.xlsx', 'id_lancamento'): 'Identificador único para cada lançamento financeiro.',
        ('user_sheets/financeiro.xlsx', 'data_lancamento'): 'Data em que o lançamento financeiro ocorreu ou foi registrado.',
        ('user_sheets/financeiro.xlsx', 'tipo_lancamento'): 'Classificação do lançamento como Receita ou Despesa.',
        ('user_sheets/financeiro.xlsx', 'descricao_lancamento'): 'Descrição detalhada do lançamento financeiro.',
        ('user_sheets/financeiro.xlsx', 'valor_lancamento'): 'O valor monetário do lançamento.',
        ('user_sheets/financeiro.xlsx', 'moeda'): 'A moeda em que o valor do lançamento está expresso.',
        ('user_sheets/financeiro.xlsx', 'conta_contabil'): 'Conta contábil associada ao lançamento.',
        ('user_sheets/financeiro.xlsx', 'centro_custo'): 'Centro de custo ou centro de lucro relacionado ao lançamento.',
        ('user_sheets/financeiro.xlsx', 'id_referencia_origem'): 'ID do documento ou transação que originou o lançamento (ex: id_pedido, id_nota_fiscal).',
        ('user_sheets/financeiro.xlsx', 'status_pagamento'): 'Status do pagamento (Pago, Aberto, Atrasado, Parcial).',
        ('user_sheets/financeiro.xlsx', 'data_vencimento'): 'Data de vencimento para pagamentos ou data esperada para recebimentos.',
        ('user_sheets/financeiro.xlsx', 'data_pagamento'): 'Data em que o pagamento foi efetuado ou o recebimento foi concretizado.',
        ('user_sheets/financeiro.xlsx', 'meio_pagamento'): 'Forma como o pagamento foi realizado (ex: Boleto, Pix, Cartão de Crédito).',
        ('user_sheets/financeiro.xlsx', 'banco_origem_destino'): 'Informações do banco de origem ou destino da transação.',
        ('user_sheets/financeiro.xlsx', 'observacoes_financeiras'): 'Campo para observações adicionais relacionadas ao lançamento financeiro.',
        ('user_sheets/financeiro.xlsx', 'imposto_valor'): 'Valor do imposto incidente sobre o lançamento.',
        ('user_sheets/financeiro.xlsx', 'imposto_tipo'): 'Tipo de imposto (ex: ICMS, IPI, PIS/COFINS).',
        ('user_sheets/financeiro.xlsx', 'fornecedor_cliente_associado'): 'ID do fornecedor ou cliente associado a este lançamento.',
        ('user_sheets/financeiro.xlsx', 'conciliado_banco'): 'Indica se o lançamento foi conciliado com o extrato bancário (Sim/Não).',
        ('user_sheets/financeiro.xlsx', 'saldo_conta'): 'Saldo da conta financeira após o lançamento (calculado).',
        ('user_sheets/manutencao.xlsx', 'id_ordem_manutencao'): 'Identificador único de cada ordem de manutenção.',
        ('user_sheets/manutencao.xlsx', 'id_ativo_equipamento'): 'Identificador do ativo ou equipamento que necessita de manutenção.',
        ('user_sheets/manutencao.xlsx', 'tipo_manutencao'): 'Classificação da manutenção (Preventiva, Corretiva, Preditiva).',
        ('user_sheets/manutencao.xlsx', 'data_solicitacao'): 'Data em que a manutenção foi solicitada ou gerada.',
        ('user_sheets/manutencao.xlsx', 'data_inicio_manutencao'): 'Data e hora de início real da manutenção.',
        ('user_sheets/manutencao.xlsx', 'data_fim_manutencao'): 'Data e hora de fim real da manutenção.',
        ('user_sheets/manutencao.xlsx', 'descricao_problema'): 'Descrição do problema que motivou a manutenção.',
        ('user_sheets/manutencao.xlsx', 'acoes_executadas'): 'Descrição das ações tomadas para realizar a manutenção.',
        ('user_sheets/manutencao.xlsx', 'pecas_substituidas'): 'Lista ou descrição das peças que foram substituídas na manutenção.',
        ('user_sheets/manutencao.xlsx', 'custo_pecas_manutencao'): 'Custo total das peças utilizadas na manutenção.',
        ('user_sheets/manutencao.xlsx', 'horas_mao_obra_manutencao'): 'Total de horas de mão de obra dedicadas à manutenção.',
        ('user_sheets/manutencao.xlsx', 'responsavel_manutencao'): 'Colaborador ou equipe responsável pela execução da manutenção.',
        ('user_sheets/manutencao.xlsx', 'status_manutencao'): 'Status atual da ordem de manutenção (Programada, Em Andamento, Concluída).',
        ('user_sheets/manutencao.xlsx', 'proxima_manutencao_programada'): 'Data da próxima manutenção agendada para este ativo/equipamento.',
        ('user_sheets/manutencao.xlsx', 'horimetro_leitura'): 'Leitura do horímetro do equipamento no momento da manutenção (se aplicável).',
        ('user_sheets/manutencao.xlsx', 'km_leitura'): 'Leitura do hodômetro do equipamento em KM (se aplicável a veículos).',
        ('user_sheets/manutencao.xlsx', 'observacoes_manutencao'): 'Campo para observações adicionais sobre a manutenção.',
        ('user_sheets/manutencao.xlsx', 'falha_causa_raiz'): 'Causa raiz identificada para a falha do equipamento.',
        ('user_sheets/manutencao.xlsx', 'criticidade_ativo'): 'Nível de criticidade do ativo (ex: Alta, Média, Baixa) para operações.',
        ('user_sheets/manutencao.xlsx', 'anexo_documentacao_manutencao'): 'Link ou referência para documentação, manuais, ou anexos da manutenção.',
        ('user_sheets/output.xlsx', 'id'): 'Identificador único do item em estoque ou cadastro.',
        ('user_sheets/output.xlsx', 'part_number'): 'Número da peça ou código do item.',
        ('user_sheets/output.xlsx', 'description'): 'Descrição do item.',
        ('user_sheets/output.xlsx', 'unidade_medida_compra'): 'Unidade de medida utilizada ao comprar o item.',
        ('user_sheets/output.xlsx', 'unidade_medida_consumo'): 'Unidade de medida utilizada ao consumir/expedir o item.',
        ('user_sheets/output.xlsx', 'deposito_padrao'): 'Depósito principal onde o item é geralmente armazenado.',
        ('user_sheets/output.xlsx', 'endereco_padrao'): 'Endereço padrão do item dentro do depósito.',
        ('user_sheets/output.xlsx', 'quantidade'): 'Quantidade total atual do item em estoque (soma de todos os locais).',
        ('user_sheets/output.xlsx', 'quantidade_comprada'): 'Quantidade total acumulada do item que foi comprada.',
        ('user_sheets/output.xlsx', 'quantidade_reservada'): 'Quantidade do item que está reservada para ordens de produção ou venda.',
        ('user_sheets/output.xlsx', 'grupo'): 'Classificação do item por grupo (ex: eletrônicos, mecânicos).',
        ('user_sheets/output.xlsx', 'quantidade_disponivel'): 'Quantidade do item disponível para uso ou venda (quantidade - quantidade_reservada).',
        ('user_sheets/output.xlsx', 'estoque_minimo'): 'Nível mínimo de estoque desejado para o item.',
        ('user_sheets/output.xlsx', 'estoque_maximo'): 'Nível máximo de estoque desejado para o item.',
        ('user_sheets/output.xlsx', 'ponto_de_pedido'): 'Nível de estoque que aciona um novo pedido de compra/produção.',
        ('user_sheets/output.xlsx', 'lead_time_reposicao_dias'): 'Tempo em dias para o item ser reposto no estoque.',
        ('user_sheets/output.xlsx', 'custo_unitario_medio'): 'Custo unitário médio ponderado do item em estoque.',
        ('user_sheets/output.xlsx', 'valor_total_estoque'): 'Valor monetário total do item em estoque.',
        ('user_sheets/output.xlsx', 'tipo_item'): 'Classificação do item (ex: matéria-prima, componente, produto acabado).',
        ('user_sheets/output.xlsx', 'fornecedor_principal_id'): 'Identificador do fornecedor principal do item.',
        ('user_sheets/output.xlsx', 'data_ultima_compra'): 'Data da última vez que o item foi comprado.',
        ('user_sheets/output.xlsx', 'data_ultima_saida'): 'Data da última vez que o item saiu do estoque.',
        ('user_sheets/output.xlsx', 'status_item'): 'Status do item no sistema (ativo, obsoleto, descontinuado).',
        ('user_sheets/output.xlsx', 'codigo_barras'): 'Código de barras do item para leitura.;',
        ('user_sheets/output.xlsx', 'data_validade'): 'Data de validade do item (se perecível).',
        ('user_sheets/producao.xlsx', 'id_producao_finalizada'): 'Identificador único para cada registro de produção finalizada.',
        ('user_sheets/producao.xlsx', 'id_ordem_producao'): 'Identificador da ordem de produção à qual este registro se refere.',
        ('user_sheets/producao.xlsx', 'part_number_produzido'): 'Número da peça ou código do produto que foi produzido.',
        ('user_sheets/producao.xlsx', 'quantidade_produzida'): 'Quantidade de itens que foram produzidos e finalizados.',
        ('user_sheets/producao.xlsx', 'unidade_medida_produzida'): 'Unidade de medida para a quantidade produzida.',
        ('user_sheets/producao.xlsx', 'data_producao'): 'Data em que a produção foi finalizada.',
        ('user_sheets/producao.xlsx', 'hora_inicio_producao'): 'Hora de início da produção.',
        ('user_sheets/producao.xlsx', 'hora_fim_producao'): 'Hora de fim da produção.',
        ('user_sheets/producao.xlsx', 'responsavel_producao'): 'Colaborador responsável pela finalização da produção.',
        ('user_sheets/producao.xlsx', 'lote_producao_gerado'): 'Número do lote gerado para os itens produzidos.',
        ('user_sheets/producao.xlsx', 'quantidade_refugada'): 'Quantidade de itens que foram descartados ou não passaram no controle de qualidade.',
        ('user_sheets/producao.xlsx', 'motivo_refugo'): 'Motivo específico pelo qual os itens foram refugados.',
        ('user_sheets/producao.xlsx', 'custo_real_producao_unitario'): 'Custo real unitário para produzir o item.',
        ('user_sheets/producao.xlsx', 'destino_estoque'): 'Depósito para onde os itens produzidos foram encaminhados.',
        ('user_sheets/producao.xlsx', 'status_qualidade_lote'): 'Status de qualidade do lote produzido (Aprovado, Reprovado, Retrabalho).',
        ('user_sheets/producao.xlsx', 'tempo_total_operacional'): 'Tempo total gasto nas operações de produção.',
        ('user_sheets/producao.xlsx', 'identificacao_maquina_utilizada'): 'Identificador da máquina principal utilizada na produção.',
        ('user_sheets/producao.xlsx', 'observacoes_producao'): 'Campo para observações adicionais sobre o processo ou resultado da produção.',
        ('user_sheets/producao.xlsx', 'sequencia_operacao_concluida'): 'A última operação da rota que foi concluída para este registro de produção.',
        ('user_sheets/producao.xlsx', 'consumo_materiais_real'): 'Informações sobre o consumo real de matérias-primas na produção.',
        ('user_sheets/pedidos.xlsx', 'id_pedido'): 'Identificador único para cada pedido (compra ou venda).',
        ('user_sheets/pedidos.xlsx', 'tipo_pedido'): 'Classificação do pedido (Compra ou Venda).',
        ('user_sheets/pedidos.xlsx', 'data_emissao_pedido'): 'Data em que o pedido foi emitido ou registrado.',
        ('user_sheets/pedidos.xlsx', 'id_cliente_fornecedor'): 'Identificador do cliente (para pedido de venda) ou fornecedor (para pedido de compra).',
        ('user_sheets/pedidos.xlsx', 'nome_cliente_fornecedor'): 'Nome completo do cliente ou razão social do fornecedor.',
        ('user_sheets/pedidos.xlsx', 'status_pedido'): 'Status atual do pedido (Aberto, Em Processamento, Atendido, Cancelado).',
        ('user_sheets/pedidos.xlsx', 'data_entrega_prevista'): 'Data de entrega prevista para o pedido.',
        ('user_sheets/pedidos.xlsx', 'data_entrega_real'): 'Data em que o pedido foi efetivamente entregue ou recebido.',
        ('user_sheets/pedidos.xlsx', 'valor_total_pedido'): 'Valor monetário total do pedido.',
        ('user_sheets/pedidos.xlsx', 'condicao_pagamento'): 'Condições de pagamento acordadas para o pedido.',
        ('user_sheets/pedidos.xlsx', 'observacoes_pedido'): 'Campo para observações adicionais sobre o pedido.',
        ('user_sheets/pedidos.xlsx', 'id_item_pedido'): 'Identificador do item específico dentro do pedido.',
        ('user_sheets/pedidos.xlsx', 'quantidade_item_pedido'): 'Quantidade do item solicitado no pedido.',
        ('user_sheets/pedidos.xlsx', 'preco_unitario_item_pedido'): 'Preço unitário do item no momento do pedido.',
        ('user_sheets/pedidos.xlsx', 'subtotal_item_pedido'): 'Subtotal do item dentro do pedido antes de impostos/descontos.',
        ('user_sheets/pedidos.xlsx', 'impostos_item_pedido'): 'Valor dos impostos incidentes sobre o item no pedido.',
        ('user_sheets/pedidos.xlsx', 'descontos_item_pedido'): 'Valor dos descontos aplicados ao item no pedido.',
        ('user_sheets/pedidos.xlsx', 'data_ultima_atualizacao_pedido'): 'Data da última vez que o pedido foi atualizado no sistema.',
        ('user_sheets/pedidos.xlsx', 'usuario_ultima_atualizacao_pedido'): 'Usuário que realizou a última atualização no pedido.',
        ('user_sheets/pedidos.xlsx', 'rastreamento_envio'): 'Código ou link de rastreamento para o envio do pedido.',
        ('user_sheets/programacao.xlsx', 'id_programacao'): 'Identificador único de cada item de programação da produção.',
        ('user_sheets/programacao.xlsx', 'id_ordem_producao'): 'Identificador da ordem de produção que está sendo programada.',
        ('user_sheets/programacao.xlsx', 'data_inicio_programada'): 'Data de início programada para a operação ou ordem de produção.',
        ('user_sheets/programacao.xlsx', 'hora_inicio_programada'): 'Hora de início programada para a operação ou ordem de produção.',
        ('user_sheets/programacao.xlsx', 'data_fim_programada'): 'Data de fim programada para a operação ou ordem de produção.',
        ('user_sheets/programacao.xlsx', 'hora_fim_programada'): 'Hora de fim programada para a operação ou ordem de produção.',
        ('user_sheets/programacao.xlsx', 'recurso_alocado'): 'Recurso (máquina, posto) que foi alocado para esta programação.',
        ('user_sheets/programacao.xlsx', 'operacao_programada'): 'Operação específica da rota de produção que está sendo programada.',
        ('user_sheets/programacao.xlsx', 'sequencia_na_fila'): 'Posição na fila de trabalho do recurso ou prioridade de execução.',
        ('user_sheets/programacao.xlsx', 'status_programacao'): 'Status da programação (Agendado, Em Andamento, Concluído, Atrasado).',
        ('user_sheets/programacao.xlsx', 'tempo_restante_programado'): 'Tempo restante estimado para a conclusão da operação/ordem programada.',
        ('user_sheets/programacao.xlsx', 'desvio_tempo_real'): 'Diferença entre o tempo programado e o tempo real de execução.',
        ('user_sheets/programacao.xlsx', 'motivo_desvio'): 'Causa específica para o desvio de tempo na programação.',
        ('user_sheets/programacao.xlsx', 'prioridade_programacao'): 'Nível de prioridade da programação (Alta, Média, Baixa).',
        ('user_sheets/programacao.xlsx', 'dependencias_operacionais'): 'Outras operações ou tarefas que precisam ser concluídas antes desta programação.',
        ('user_sheets/programacao.xlsx', 'capacidade_utilizada_percentual'): 'Percentual da capacidade do recurso que será utilizada por esta programação.',
        ('user_sheets/programacao.xlsx', 'tempo_setup_alocado'): 'Tempo de setup alocado para a operação nesta programação.'
    }

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Atualizador de Cabeçalhos do Banco de Dados")
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.status_label = QLabel("Clique em 'Atualizar' para coletar e salvar os cabeçalhos das planilhas.")
        self.layout.addWidget(self.status_label)

        self.update_button = QPushButton("Atualizar Cabeçalhos")
        self.update_button.clicked.connect(self._update_db_headers)
        self.layout.addWidget(self.update_button)

        self.table = QTableWidget()
        self.table.setEditTriggers(QTableWidget.NoEditTriggers) # Esta ferramenta é para visualizar/atualizar, não para edição direta de cabeçalhos
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.layout.addWidget(self.table)
        
        # Novas colunas para pagina_arquivo e descr_variavel
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Arquivo (Caminho)", "Nome da Coluna (Cabeçalho)", "pagina_arquivo", "descr_variavel"])

    def _update_db_headers(self):
        """Coleta cabeçalhos de todos os arquivos Excel especificados e os salva em db.xlsx (planilha db_db)."""
        self.status_label.setText("Coletando cabeçalhos...")
        QApplication.processEvents() # Atualiza a GUI imediatamente

        all_headers_data = []

        # Lista de todos os diretórios relevantes
        directories = {
            "user_sheets": USER_SHEETS_DIR,
            "app_sheets": APP_SHEETS_DIR
        }

        for dir_name, directory_path in directories.items():
            if not os.path.exists(directory_path):
                self.status_label.setText(f"Aviso: Diretório não encontrado: {directory_path}")
                QMessageBox.warning(self, "Diretório Não Encontrado", f"O diretório '{directory_path}' não existe.")
                continue

            for filename in os.listdir(directory_path):
                if filename.endswith('.xlsx'):
                    file_path = os.path.join(directory_path, filename)
                    try:
                        wb = openpyxl.load_workbook(file_path)
                        for sheet_name in wb.sheetnames:
                            sheet = wb[sheet_name]
                            if sheet.max_row > 0:
                                headers = [cell.value for cell in sheet[1]]
                                for header in headers:
                                    # Normalize path for dictionary lookup (use forward slashes)
                                    file_path_rel_normalized = os.path.relpath(file_path, project_root).replace('\\', '/')
                                    
                                    # Obtém a descrição da variável do mapeamento, se existir
                                    description = self.DESCRIPTION_MAP.get(
                                        (file_path_rel_normalized, str(header) if header is not None else ""), 
                                        "Descrição não disponível."
                                    )
                                    
                                    # Adiciona o caminho relativo, o cabeçalho, a página da planilha e a descrição
                                    all_headers_data.append([
                                        os.path.relpath(file_path, project_root), # Mantém o separador original para exibição
                                        str(header) if header is not None else "",
                                        "obs IA olhar em runtime e salvar", # Valor fixo para pagina_arquivo
                                        description
                                    ])
                    except Exception as e:
                        print(f"Erro ao ler {file_path}: {e}")
                        self.status_label.setText(f"Erro ao ler {filename}: {e}")
                        QMessageBox.critical(self, "Erro de Leitura", f"Não foi possível ler o arquivo {filename}: {e}")
                        continue
        
        try:
            wb_db = openpyxl.load_workbook(DB_EXCEL_PATH)
            if "db_db" not in wb_db.sheetnames:
                ws_db = wb_db.create_sheet("db_db")
            else:
                ws_db = wb_db["db_db"]
            
            # Limpa os dados existentes em db_db, mas mantém a própria planilha
            for row_idx in range(ws_db.max_row, 0, -1): # Itera da última linha para a primeira (inclusive)
                ws_db.delete_rows(row_idx)

            # Adiciona cabeçalhos para a planilha db_db
            # Inclui as novas colunas
            ws_db.append(["Arquivo (Caminho)", "Nome da Coluna (Cabeçalho)", "pagina_arquivo", "descr_variavel"])
            
            # Adiciona dados coletados
            for row_data in all_headers_data:
                ws_db.append(row_data)
            
            wb_db.save(DB_EXCEL_PATH)
            self.status_label.setText("Cabeçalhos atualizados e salvos em db.xlsx (db_db).")
            QMessageBox.information(self, "Sucesso", "Cabeçalhos atualizados e salvos em db.xlsx (db_db).")
            
            # Exibe dados coletados na tabela
            self.table.setRowCount(len(all_headers_data))
            for row_idx, row_data in enumerate(all_headers_data):
                for col_idx, cell_value in enumerate(row_data):
                    item = QTableWidgetItem(str(cell_value))
                    self.table.setItem(row_idx, col_idx, item)

        except Exception as e:
            self.status_label.setText(f"Erro ao salvar em db.xlsx: {e}")
            QMessageBox.critical(self, "Erro ao Salvar DB", f"Não foi possível salvar os cabeçalhos em db.xlsx: {e}")

# === NOVA JANELA: CONFIGURAÇÕES DE PERFIL ===
class ProfileSettingsWindow(QDialog):
    """
    Janela para exibir e potencialmente editar as configurações de perfil do usuário.
    """
    def __init__(self, username, role):
        super().__init__()
        self.setWindowTitle("Configurações de Perfil")
        self.setGeometry(200, 200, 400, 300) # x, y, largura, altura

        self.username = username
        self.role = role

        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("<h2>Informações do Perfil</h2>"))
        
        # Campo de Usuário
        user_layout = QHBoxLayout()
        user_layout.addWidget(QLabel("<b>Usuário:</b>"))
        user_label = QLabel(self.username)
        user_layout.addWidget(user_label)
        user_layout.addStretch()
        layout.addLayout(user_layout)

        # Campo de Papel
        role_layout = QHBoxLayout()
        role_layout.addWidget(QLabel("<b>Papel:</b>"))
        role_label = QLabel(self.role)
        role_layout.addWidget(role_label)
        role_layout.addStretch()
        layout.addLayout(role_layout)

        layout.addStretch() # Empurra o conteúdo para cima

        # Campos de configurações de exemplo
        layout.addWidget(QLabel("<h3>Configurações Adicionais (Exemplo):</h3>"))
        
        email_layout = QHBoxLayout()
        email_layout.addWidget(QLabel("Email:"))
        email_input = QLineEdit("usuario@example.com") # Valor padrão
        email_layout.addWidget(email_input)
        layout.addLayout(email_layout)

        lang_layout = QHBoxLayout()
        lang_layout.addWidget(QLabel("Idioma:"))
        lang_input = QLineEdit("Português (Brasil)") # Valor padrão
        lang_layout.addWidget(lang_input)
        layout.addLayout(lang_layout)

        # Botão de Fechar
        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(self.accept) # Fecha o diálogo

        layout.addWidget(close_btn)
        self.setLayout(layout)


# === GUI PRINCIPAL ===
class TeamcenterStyleGUI(QMainWindow):
    """
    A GUI do aplicativo principal, estilizada para se assemelhar ao Teamcenter.
    Fornece uma visualização em árvore do espaço de trabalho, área de conteúdo com abas e uma barra de ferramentas.
    """
    def __init__(self, user):
        super().__init__()
        self.setWindowTitle("Plataforma 5revolution")
        # self.setGeometry(100, 100, 1280, 800) # O tamanho inicial é definido, mas será maximizado
        
        self.username = user["username"]
        self.role = user["role"]
        self.tools = load_tools_from_excel() 
        self.permissions = load_role_permissions()

        self._create_toolbar()
        self._create_main_layout()

        self.statusBar().showMessage(f"Logado como: {self.username} | Papel: {self.role}")

    def _create_toolbar(self):
        """Cria a barra de ferramentas principal do aplicativo."""
        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setMovable(False) 
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)

        self.tools_btn = QToolButton()
        self.tools_btn.setText("🛠 Ferramentas")
        self.tools_btn.setPopupMode(QToolButton.InstantPopup) 
        tools_menu = QMenu()

        allowed_tools = self.permissions.get(self.role, []) 
        for tid, tool in self.tools.items():
            if allowed_tools == "all" or tid in allowed_tools:
                action = tools_menu.addAction(tool["name"])
                
                # --- Abertura Dinâmica de Ferramentas com base no ID da ferramenta ---
                if tool["id"] == "mod4": # Engenharia (Workflow)
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, EngenhariaWorkflowTool(file_path=ENGENHARIA_EXCEL_PATH)))
                elif tool["id"] == "mes_pcp": 
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, self._create_mes_pcp_tool_widget()))
                elif tool["id"] == "prod_data":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ProductDataTool(file_path=OUTPUT_EXCEL_PATH)))
                elif tool["id"] == "bom_manager":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, BomManagerTool(file_path=ENGENHARIA_EXCEL_PATH, sheet_name="Estrutura")))
                elif tool["id"] == "configurador":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ConfiguradorTool(file_path=CONFIGURADOR_EXCEL_PATH)))
                elif tool["id"] == "colab":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ColaboradoresTool(file_path=COLABORADORES_EXCEL_PATH)))
                elif tool["id"] == "items_tool":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ItemsTool(file_path=ESTOQUE_EXCEL_PATH))) 
                elif tool["id"] == "manuf":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ManufacturingTool(file_path=MANUFACTURING_DATA_EXCEL_PATH)))
                elif tool["id"] == "pcp_tool":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, PcpTool(file_path=PROGRAMACAO_EXCEL_PATH)))
                elif tool["id"] == "estoque_tool":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, EstoqueTool(file_path=ESTOQUE_EXCEL_PATH))) 
                elif tool["id"] == "financeiro":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, FinanceiroTool(file_path=FINANCEIRO_EXCEL_PATH)))
                elif tool["id"] == "pedidos":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, PedidosTool(file_path=PEDIDOS_EXCEL_PATH)))
                elif tool["id"] == "manutencao":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ManutencaoTool(file_path=MANUTENCAO_EXCEL_PATH)))
                elif tool["id"] == "rpi_tool":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, RpiTool(file_path=RPI_EXCEL_PATH)))
                elif tool["id"] == "engenharia_data":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, EngenhariaDataTool(file_path=ENGENHARIA_EXCEL_PATH)))
                elif tool["id"] == "db_headers_updater": # Nova ferramenta para db_db
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, DbHeadersUpdaterTool()))
                else: 
                    action.triggered.connect(lambda chk=False, title=tool["name"], desc=tool["description"]: self._open_tab(title, QLabel(desc)))
        
        # Ações somente para administrador
        if self.role == "admin":
            tools_menu.addSeparator()
            admin_menu = tools_menu.addMenu("👑 Ferramentas Admin")
            create_engenharia_action = admin_menu.addAction("Criar/Atualizar engenharia.xlsx")
            create_engenharia_action.triggered.connect(self._run_create_engenharia_script)
        
        self.tools_btn.setMenu(tools_menu)
        self.toolbar.addWidget(self.tools_btn)

        self.profile_btn = QToolButton()
        self.profile_btn.setText(f"👤 {self.username}") 
        self.profile_btn.setPopupMode(QToolButton.InstantPopup)
        profile_menu = QMenu()
        profile_menu.addAction("⚙️ Configurações", self._open_options)
        profile_menu.addSeparator() 
        profile_menu.addAction("🔒 Sair", self._logout)
        self.profile_btn.setMenu(profile_menu)
        self.toolbar.addWidget(self.profile_btn)

    def _create_main_layout(self):
        """Cria o layout principal dividido com a visualização em árvore e as abas."""
        self.splitter = QSplitter() 

        left_pane_widget = QWidget()
        left_pane_layout = QVBoxLayout(left_pane_widget)
        left_pane_layout.setContentsMargins(0, 0, 0, 0) 

        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Espaço de Trabalho")
        self._populate_sample_tree() 
        self.tree.expandAll() 
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu) 
        self.tree.customContextMenuRequested.connect(self._show_tree_context_menu)
        self.tree.itemDoubleClicked.connect(self._open_file_from_tree) 

        search_layout = QHBoxLayout()
        self.item_search_bar = QLineEdit()
        self.item_search_bar.setPlaceholderText("Pesquisar itens...")
        self.item_search_bar.returnPressed.connect(self.handle_item_search) 
        self.search_items_btn = QPushButton("🔍")
        self.search_items_btn.clicked.connect(self.handle_item_search)

        search_layout.addWidget(self.item_search_bar)
        search_layout.addWidget(self.search_items_btn)

        left_pane_layout.addWidget(QLabel("Espaço de Trabalho"))
        left_pane_layout.addLayout(search_layout)
        left_pane_layout.addWidget(self.tree)

        self.tabs = QTabWidget()
        self.tabs.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tabs.customContextMenuRequested.connect(self._show_tab_context_menu)
        self.tabs.setTabsClosable(True) 
        self.tabs.tabCloseRequested.connect(self.tabs.removeTab) 

        welcome_widget = QWidget()
        welcome_layout = QVBoxLayout()
        welcome_layout.addWidget(QLabel(f"Bem-vindo {self.username} – Papel: {self.role}"))
        welcome_widget.setLayout(welcome_layout)
        self.tabs.addTab(welcome_widget, "Início")

        self.splitter.addWidget(left_pane_widget) 
        self.splitter.addWidget(self.tabs)
        self.splitter.setStretchFactor(1, 4) 

        container = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(self.splitter)
        container.setLayout(layout)
        self.setCentralWidget(container)

    def _populate_sample_tree(self):
        """
        Popula a árvore com dados de projeto/variante de exemplo e arquivos .xlsx
        dos diretórios user_sheets e app_sheets.
        Garante que não haja duplicatas limpando primeiro e adicionando explicitamente com base nos caminhos dos arquivos.
        """
        self.tree.clear() # Limpa itens existentes para evitar duplicatas
        
        # Seção de Planilhas do Usuário
        user_sheets_root = QTreeWidgetItem(["Arquivos do Usuário (user_sheets)"])
        self.tree.addTopLevelItem(user_sheets_root)

        for filename in sorted(os.listdir(USER_SHEETS_DIR)): # Ordena para exibição consistente
            if filename.endswith('.xlsx'):
                file_path = os.path.join(USER_SHEETS_DIR, filename)
                file_item = QTreeWidgetItem([filename])
                file_item.setData(0, Qt.UserRole, file_path) # Armazena o caminho completo
                user_sheets_root.addChild(file_item)
        
        # Seção de Planilhas do Aplicativo (somente leitura para usuários, mas visível)
        app_sheets_root = QTreeWidgetItem(["Arquivos do Sistema (app_sheets)"])
        self.tree.addTopLevelItem(app_sheets_root)

        for filename in sorted(os.listdir(APP_SHEETS_DIR)): # Ordena para exibição consistente
            if filename.endswith('.xlsx'):
                file_path = os.path.join(APP_SHEETS_DIR, filename)
                file_item = QTreeWidgetItem([filename])
                file_item.setData(0, Qt.UserRole, file_path) # Armazena o caminho completo
                app_sheets_root.addChild(file_item)

        # Seção de Itens de Projeto/Espaço de Trabalho
        projects_root = QTreeWidgetItem(["Projetos/Espaço de Trabalho"])
        self.tree.addTopLevelItem(projects_root)

        project1 = QTreeWidgetItem(["Projeto Demo - Rev A"])
        project1.addChild(QTreeWidgetItem(["Peça-001"]))
        project1.addChild(QTreeWidgetItem(["Montagem-001"]))
        projects_root.addChild(project1)

        project2 = QTreeWidgetItem(["Variante Amostra - V1.0"])
        project2.addChild(QTreeWidgetItem(["Componente-XYZ"]))
        projects_root.addChild(project2)

        self.tree.expandAll() 

    def _open_file_from_tree(self, item, column):
        """
        Abre um arquivo Excel da visualização em árvore em um ExcelViewerTool genérico
        ou uma ferramenta específica com base no nome do arquivo.
        """
        file_path = item.data(0, Qt.UserRole)
        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "Erro ao Abrir", f"Não foi possível abrir o arquivo: {file_path}. Arquivo não encontrado ou caminho inválido.")
            return

        file_name = os.path.basename(file_path)
        tab_title = f"Visualizador: {file_name}"
        
        # Determina a ferramenta
        tool_widget = None

        if file_name == os.path.basename(ENGENHARIA_EXCEL_PATH):
            tool_widget = EngenhariaDataTool(file_path=file_path)
        elif file_name == os.path.basename(COLABORADORES_EXCEL_PATH):
             tool_widget = ColaboradoresTool(file_path=file_path)
        elif file_name == os.path.basename(OUTPUT_EXCEL_PATH):
             tool_widget = ProductDataTool(file_path=file_path)
        elif file_name == os.path.basename(BOM_DATA_EXCEL_PATH):
             tool_widget = BomManagerTool(file_path=file_path) # Arquivo BOM padrão
        elif file_name == os.path.basename(CONFIGURADOR_EXCEL_PATH):
             tool_widget = ConfiguradorTool(file_path=file_path)
        elif file_name == os.path.basename(ESTOQUE_EXCEL_PATH): # ItemsTool agora gerencia estoque.xlsx
             tool_widget = ItemsTool(file_path=file_path)
        elif file_name == os.path.basename(ITEMS_DATA_EXCEL_PATH): # Para o items_data.xlsx original
             tool_widget = ItemsTool(file_path=file_path, read_only=True) # Exemplo: talvez items_data.xlsx seja somente leitura se usado junto com estoque.xlsx
        elif file_name == os.path.basename(MANUFACTURING_DATA_EXCEL_PATH):
             tool_widget = ManufacturingTool(file_path=file_path)
        elif file_name == os.path.basename(PROGRAMACAO_EXCEL_PATH):
             tool_widget = PcpTool(file_path=file_path)
        elif file_name == os.path.basename(FINANCEIRO_EXCEL_PATH):
             tool_widget = FinanceiroTool(file_path=file_path)
        elif file_name == os.path.basename(PEDIDOS_EXCEL_PATH):
             tool_widget = PedidosTool(file_path=file_path)
        elif file_name == os.path.basename(MANUTENCAO_EXCEL_PATH):
             tool_widget = ManutencaoTool(file_path=file_path)
        elif file_name == os.path.basename(RPI_EXCEL_PATH):
             tool_widget = RpiTool(file_path=file_path)
        elif file_name == os.path.basename(DB_EXCEL_PATH): # db.xlsx deve ser aberto pelo DbHeadersUpdaterTool
            tool_widget = DbHeadersUpdaterTool() # Esta ferramenta carregará/exibirá a planilha db_db
            tab_title = "Atualizador de Cabeçalhos do DB"
        elif file_name == os.path.basename(TOOLS_EXCEL_PATH): # Tools.xlsx
            tool_widget = ExcelViewerTool(file_path=file_path, read_only=True) # Exemplo: tools.xlsx pode ser somente leitura
        elif file_name == os.path.basename(MODULES_EXCEL_PATH):
            tool_widget = ExcelViewerTool(file_path=file_path, read_only=True)
        elif file_name == os.path.basename(PERMISSIONS_EXCEL_PATH):
            tool_widget = ExcelViewerTool(file_path=file_path, read_only=True)
        elif file_name == os.path.basename(ROLES_TOOLS_EXCEL_PATH):
            tool_widget = ExcelViewerTool(file_path=file_path, read_only=True)
        elif file_name == os.path.basename(USERS_EXCEL_PATH): # Se este for um arquivo separado e não uma aba em db.xlsx
            tool_widget = ExcelViewerTool(file_path=file_path, read_only=True)
        elif file_name == os.path.basename(MAIN_EXCEL_PATH):
            tool_widget = ExcelViewerTool(file_path=file_path, read_only=True)
        else: # Visualizador Excel genérico para qualquer outro arquivo .xlsx
            tool_widget = ExcelViewerTool(file_path=file_path)

        # Para casos onde Engenharia.xlsx é aberto por outra ferramenta (como RPI ou Items)
        if file_name == os.path.basename(ENGENHARIA_EXCEL_PATH):
            # A lógica de somente leitura baseada no nome do arquivo já está dentro do __init__ das ferramentas,
            # mas podemos forçar aqui se a intenção for sempre ter somente leitura para engenharia.xlsx
            # quando acessado por certas ferramentas. Por enquanto, a flag `read_only` passada no construtor
            # das ferramentas já cuida disso.
            pass


        if tool_widget:
            self._open_tab(tab_title, tool_widget)
        else:
            QMessageBox.critical(self, "Erro ao Abrir Ferramenta", f"Não foi possível encontrar uma ferramenta adequada para abrir {file_name}.")


    def _create_mes_pcp_tool_widget(self):
        """Cria o widget para a ferramenta MES (Apontamento Fábrica)."""
        mes_widget = QWidget()
        mes_layout = QVBoxLayout()
        mes_layout.addWidget(QLabel("<h2>MES (Apontamento Fábrica)</h2>"))
        mes_layout.addWidget(QLabel("Inserir dados de produção, acompanhar progresso e gerenciar operações de chão de fábrica."))

        form_layout = QVBoxLayout()
        self.mes_order_id_input = QLineEdit()
        self.mes_order_id_input.setPlaceholderText("ID da Ordem de Produção")
        self.mes_item_code_input = QLineEdit()
        self.mes_item_code_input.setPlaceholderText("Código do Item")
        self.mes_quantity_input = QLineEdit()
        self.mes_quantity_input.setPlaceholderText("Quantidade Produzida")
        # For simplicity, using QLineEdit. For actual datetime, consider QDateTimeEdit.
        self.mes_start_time_input = QLineEdit()
        self.mes_start_time_input.setPlaceholderText("Hora de Início (AAAA-MM-DD HH:MM)")
        self.mes_end_time_input = QLineEdit()
        self.mes_end_time_input.setPlaceholderText("Hora de Término (AAAA-MM-DD HH:MM)")

        submit_btn = QPushButton("Enviar Dados de Produção")
        submit_btn.clicked.connect(self._submit_mes_data) # Conecta a um manipulador de envio

        form_layout.addWidget(QLabel("ID da Ordem de Produção:"))
        form_layout.addWidget(self.mes_order_id_input)
        form_layout.addWidget(QLabel("Código do Item:"))
        form_layout.addWidget(self.mes_item_code_input)
        form_layout.addWidget(QLabel("Quantidade Produzida:"))
        form_layout.addWidget(self.mes_quantity_input)
        form_layout.addWidget(QLabel("Hora de Início:"))
        form_layout.addWidget(self.mes_start_time_input)
        form_layout.addWidget(QLabel("Hora de Término:"))
        form_layout.addWidget(self.mes_end_time_input)
        form_layout.addWidget(submit_btn)

        mes_layout.addLayout(form_layout)
        mes_layout.addStretch() # Empurra o conteúdo para cima
        mes_widget.setLayout(mes_layout)
        return mes_widget

    def _submit_mes_data(self):
        """Lida com o envio de dados MES (espaço reservado)."""
        order_id = self.mes_order_id_input.text()
        item_code = self.mes_item_code_input.text()
        quantity = self.mes_quantity_input.text()
        start_time = self.mes_start_time_input.text()
        end_time = self.mes_end_time_input.text()

        if not all([order_id, item_code, quantity, start_time, end_time]):
            QMessageBox.warning(self, "Erro de Entrada", "Todos os campos MES devem ser preenchidos.")
            return

        # Em uma aplicação real, você salvaria esses dados em um banco de dados ou arquivo
        QMessageBox.information(self, "Dados MES Enviados",
                                f"Dados de Produção Enviados:\n"
                                f"ID da Ordem: {order_id}\n"
                                f"Código do Item: {item_code}\n"
                                f"Quantidade: {quantity}\n"
                                f"Início: {start_time}\n"
                                f"Término: {end_time}")
        # Limpa os campos após o envio
        self.mes_order_id_input.clear()
        self.mes_item_code_input.clear()
        self.mes_quantity_input.clear()
        self.mes_start_time_input.clear()
        self.mes_end_time_input.clear()

    def handle_item_search(self):
        """
        Executa uma pesquisa em itens do espaço de trabalho e exibe os resultados em uma caixa de diálogo.
        """
        search_term = self.item_search_bar.text().strip().lower()
        if not search_term:
            QMessageBox.information(self, "Pesquisar", "Por favor, digite um termo de pesquisa.")
            return

        results = [item for item in WORKSPACE_ITEMS if search_term in item.lower()]
        self.display_search_results_dialog(results)

    def display_search_results_dialog(self, results):
        """
        Exibe os resultados da pesquisa em uma nova janela QDialog.
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("Resultados da Pesquisa")
        dialog.setGeometry(self.x() + 200, self.y() + 100, 400, 300) # Posição relativa à janela principal

        layout = QVBoxLayout(dialog)
        
        if not results:
            layout.addWidget(QLabel("Nenhum item encontrado correspondente à sua pesquisa."))
        else:
            list_widget = QListWidget()
            for item in results:
                list_widget.addItem(item)
            list_widget.itemDoubleClicked.connect(
                lambda item: self.open_selected_item_tab(item.text()) or dialog.accept()
            ) # Fecha o diálogo ao dar clique duplo
            layout.addWidget(list_widget)

        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(dialog.accept) # Fecha o diálogo ao clicar no botão
        layout.addWidget(close_btn)

        dialog.exec_() # Exibe o diálogo modalmente

    def open_selected_item_tab(self, item_name):
        """
        Abre uma nova aba na GUI principal para exibir detalhes do item selecionado.
        """
        tab_title = f"Detalhes: {item_name}"

        # Verifica se a aba já está aberta
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == tab_title:
                self.tabs.setCurrentIndex(i)
                QMessageBox.information(self, "Informação", f"A guia para '{item_name}' já está aberta.")
                return

        # Cria um widget para os detalhes do item
        item_details_widget = QWidget()
        item_details_layout = QVBoxLayout()
        item_details_layout.addWidget(QLabel(f"<h2>Detalhes do Item: {item_name}</h2>"))
        item_details_layout.addWidget(QLabel(f"Exibindo detalhes abrangentes para <b>{item_name}</b>."))
        item_details_layout.addWidget(QLabel("Esta seção carregaria dados reais: propriedades, revisões, arquivos associados, etc."))
        item_details_layout.addStretch() # Empurra o conteúdo para cima
        item_details_widget.setLayout(item_details_layout)

        self._open_tab(tab_title, item_details_widget)
        QMessageBox.information(self, "Item Aberto", f"Detalhes abertos para: {item_name}")


    def _open_tab(self, title, widget_instance):
        """
        Abre uma nova aba ou alterna para uma existente.
        Aceita uma instância de widget diretamente.
        """
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == title:
                self.tabs.setCurrentIndex(i)
                return
        # Se a aba não existir, adiciona-a
        self.tabs.addTab(widget_instance, title)
        self.tabs.setCurrentIndex(self.tabs.count() - 1) # Alterna para a aba recém-aberta

    def _open_options(self):
        """Abre a caixa de diálogo de opções/configurações do usuário."""
        # Cria e exibe a janela de configurações de perfil
        settings_dialog = ProfileSettingsWindow(self.username, self.role)
        settings_dialog.exec_() # Executa como diálogo modal

    def _logout(self):
        """Faz logout do usuário atual e retorna para a tela de login."""
        confirm_logout = QMessageBox.question(self, "Confirmação de Saída", "Tem certeza de que deseja sair?",
                                              QMessageBox.Yes | QMessageBox.No)
        if confirm_logout == QMessageBox.Yes:
            self.close() # Fecha a janela principal do aplicativo
            self.login = LoginWindow() # Cria uma nova instância da janela de login
            self.login.show() # Exibe a janela de login

    def _show_tree_context_menu(self, pos):
        """Exibe um menu de contexto para itens na visualização em árvore."""
        item = self.tree.itemAt(pos)
        if not item: return

        menu = QMenu()
        file_path = item.data(0, Qt.UserRole)
        file_name = os.path.basename(file_path) if file_path else ""

        # Verifica se o arquivo está protegido
        is_protected = file_name in PROTECTED_FILES

        # Ações para arquivos .xlsx
        if file_path and file_path.endswith('.xlsx'):
            menu.addAction("Abrir no Visualizador Excel Genérico", lambda: self._open_tab(f"Visualizador: {file_name}", ExcelViewerTool(file_path=file_path)))
            
            if file_name == os.path.basename(ENGENHARIA_EXCEL_PATH):
                menu.addAction("Abrir como Estrutura (Árvore de Componentes)", lambda: self._open_tab(f"Estrutura: {file_name}", StructureViewTool(file_path=file_path, sheet_name="Estrutura")))
                menu.addAction("Abrir como BOM", lambda: self._open_tab(f"BOM: {file_name}", BomManagerTool(file_path=file_path, sheet_name="Estrutura")))
            
            if not is_protected:
                rename_action = menu.addAction("Renomear Arquivo")
                rename_action.triggered.connect(lambda: self._rename_file(item))
                
                delete_action = menu.addAction("Excluir Arquivo")
                delete_action.triggered.connect(lambda: self._delete_file(item))
            else:
                # Adiciona uma ação desabilitada para indicar proteção
                protected_rename_action = menu.addAction("Renomear (Protegido)")
                protected_rename_action.setEnabled(False)
                protected_delete_action = menu.addAction("Excluir (Protegido)")
                protected_delete_action.setEnabled(False)

        # Ações gerais para qualquer item da árvore
        else: 
            menu.addAction("🔍 Ver Detalhes", lambda: QMessageBox.information(self, "Ação Similada", f"Visualizando detalhes para: {item.text(0)} (ação simulada)"))
            menu.addAction("✏️ Editar Propriedades", lambda: QMessageBox.information(self, "Ação Similada", f"Editando propriedades para: {item.text(0)} (ação simulada)"))
            menu.addAction("❌ Excluir Item (Simulado)", lambda: QMessageBox.warning(self, "Ação Similada", f"Excluído: {item.text(0)} (ação simulada)"))

        menu.exec_(self.tree.viewport().mapToGlobal(pos)) 

    def _rename_file(self, item):
        """Renomeia um arquivo no sistema de arquivos e atualiza a visualização em árvore."""
        old_file_path = item.data(0, Qt.UserRole)
        if not old_file_path: return

        old_file_name = os.path.basename(old_file_path)

        # Impede a renomeação de arquivos protegidos
        if old_file_name in PROTECTED_FILES:
            QMessageBox.warning(self, "Operação Não Permitida", f"O arquivo '{old_file_name}' é protegido e não pode ser renomeado.")
            return

        new_file_name, ok = QInputDialog.getText(self, "Renomear Arquivo", f"Novo nome para '{old_file_name}':", QLineEdit.Normal, old_file_name)
        
        if ok and new_file_name and new_file_name != old_file_name:
            # Garante que o novo nome tenha a extensão .xlsx se for um arquivo Excel
            if old_file_name.endswith('.xlsx') and not new_file_name.endswith('.xlsx'):
                new_file_name += '.xlsx'

            new_file_path = os.path.join(os.path.dirname(old_file_path), new_file_name)

            if os.path.exists(new_file_path):
                QMessageBox.warning(self, "Erro ao Renomear", f"Já existe um arquivo com o nome '{new_file_name}'.")
                return

            try:
                os.rename(old_file_path, new_file_path)
                item.setText(0, new_file_name) # Atualiza o texto do item na árvore
                item.setData(0, Qt.UserRole, new_file_path) # Atualiza o caminho armazenado
                QMessageBox.information(self, "Sucesso", f"Arquivo renomeado para '{new_file_name}'.")
            except Exception as e:
                QMessageBox.critical(self, "Erro ao Renomear", f"Não foi possível renomear o arquivo: {e}")

    def _delete_file(self, item):
        """Exclui um arquivo do sistema de arquivos e o remove da visualização em árvore."""
        file_path = item.data(0, Qt.UserRole)
        if not file_path: return

        file_name = os.path.basename(file_path)

        # Impede a exclusão de arquivos protegidos
        if file_name in PROTECTED_FILES:
            QMessageBox.warning(self, "Operação Não Permitida", f"O arquivo '{file_name}' é protegido e não pode ser excluído.")
            return

        confirm = QMessageBox.question(self, "Confirmar Exclusão", 
                                       f"Tem certeza de que deseja excluir o arquivo '{file_name}' permanentemente?",
                                       QMessageBox.Yes | QMessageBox.No)
        
        if confirm == QMessageBox.Yes:
            try:
                os.remove(file_path)
                parent_item = item.parent()
                if parent_item:
                    parent_item.removeChild(item)
                else:
                    self.tree.invisibleRootItem().removeChild(item) # Para itens de nível superior
                QMessageBox.information(self, "Sucesso", f"Arquivo '{file_name}' excluído com sucesso.")
            except Exception as e:
                QMessageBox.critical(self, "Erro ao Excluir", f"Não foi possível excluir o arquivo: {e}")


    def _run_create_engenharia_script(self):
        """
        Executa o script create_engenharia_xlsx.py como um subprocesso.
        Esta ação deve estar disponível apenas para usuários administradores.
        """
        if self.role != "admin":
            QMessageBox.warning(self, "Permissão Negada", "Você não tem permissão para executar esta ação.")
            return

        script_path = os.path.join(project_root, "create_engenharia_xlsx.py")
        if not os.path.exists(script_path):
            QMessageBox.critical(self, "Erro de Script", f"O script '{os.path.basename(script_path)}' não foi encontrado em: {script_path}")
            return

        try:
            # Executa o script e captura a saída
            result = subprocess.run([sys.executable, script_path], capture_output=True, text=True, check=True)
            QMessageBox.information(self, "Script Executado", 
                                    f"Script '{os.path.basename(script_path)}' executado com sucesso.\n\n"
                                    f"Saída:\n{result.stdout}\n"
                                    f"Erros (se houver):\n{result.stderr}")
            self._populate_sample_tree() # Atualiza a árvore para mostrar o novo arquivo se criado
        except subprocess.CalledProcessError as e:
            QMessageBox.critical(self, "Erro de Execução do Script", 
                                 f"O script '{os.path.basename(script_path)}' falhou com erro:\n\n"
                                 f"Saída:\n{e.stdout}\n"
                                 f"Erros:\n{e.stderr}")
        except Exception as e:
            QMessageBox.critical(self, "Erro Inesperado", f"Ocorreu um erro inesperado ao executar o script: {e}")


    def _show_tab_context_menu(self, pos):
        """Exibe um menu de contexto para abas no widget de abas."""
        index = self.tabs.tabBar().tabAt(pos)
        if index < 0: return 

        menu = QMenu()
        menu.addAction("❌ Fechar Guia", lambda: self.tabs.removeTab(index))
        if self.tabs.count() > 1:
            menu.addAction("🔁 Fechar Outras Guias", lambda: self._close_other_tabs(index))
        if self.tabs.count() > 0: 
            menu.addAction("🧹 Fechar Todas as Guias", self.tabs.clear)
        menu.exec_(self.tabs.tabBar().mapToGlobal(pos))

    def _close_other_tabs(self, keep_index):
        """Fecha todas as abas, exceto a que está em 'keep_index'."""
        for i in reversed(range(self.tabs.count())):
            if i != keep_index:
                self.tabs.removeTab(i)

# === PONTO DE ENTRADA ===
if __name__ == "__main__":
    app = QApplication(sys.argv)
    login = LoginWindow()
    login.show()
    sys.exit(app.exec_())
