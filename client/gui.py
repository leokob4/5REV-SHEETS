import sys
import os
import bcrypt
import openpyxl
import json # Necessário para EngenhariaWorkflowTool (salvar/carregar JSON)
import subprocess # Necessário para _run_create_engenharia_script

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QAction, QTabWidget, QMenu, QToolButton,
    QWidget, QVBoxLayout, QSplitter, QTreeWidget, QTreeWidgetItem,
    QLabel, QLineEdit, QPushButton, QHBoxLayout, QMessageBox, QGraphicsView,
    QGraphicsScene, QGraphicsRectItem, QGraphicsLineItem, QDialog, QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView, QInputDialog, QComboBox, QGraphicsTextItem # Adicionado QGraphicsTextItem
)
from PyQt5.QtCore import Qt, QPointF, QFileInfo
from PyQt5.QtGui import QBrush, QPen, QColor, QFont # QFont é bom para QGraphicsTextItem

# --- Correção para ModuleNotFoundError: No module named 'ui' ---
# Obtém o caminho absoluto do diretório contendo gui.py
current_dir = os.path.dirname(os.path.abspath(__file__))
# Navega até a raiz do projeto (assumindo gui.py está em client/, e client/ está na raiz do projeto/)
project_root = os.path.dirname(current_dir)
# Adiciona a raiz do projeto ao sys.path para que Python possa encontrar 'ui' e 'user_sheets' etc.
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# --- Importar Módulos das Ferramentas ---
# Garanta que esses arquivos existam em ui/tools/
from ui.tools.product_data import ProductDataTool
from ui.tools.bom_manager import BomManagerTool
from ui.tools.configurador import ConfiguradorTool
from ui.tools.colaboradores import ColaboradoresTool
from ui.tools.items import ItemsTool
from ui.tools.manufacturing import ManufacturingTool
from ui.tools.pcp import PcpTool
from ui.tools.estoque import EstoqueTool
from ui.tools.financeiro import FinanceiroTool # Corrigido: FinanceiroolTool -> FinanceiroTool
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

# Itens de espaço de trabalho codificados (para a árvore de exemplo, antes da carga dinâmica)
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
        sheet = wb["access"] 
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
        self.password_input.returnPressed.connect(self.authenticate) 

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
        self.main.showMaximized() 
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
            self.users = load_users_from_excel() 
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

        self.nodes = [] 
        self.node_properties = {} 
        self.links = [] 

        self._populate_sheet_selector() 

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
                    self.sheet_selector.setCurrentIndex(0) 
                
                self.current_sheet_name = self.sheet_selector.currentText()
            
            self._load_workflow_from_excel() 

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Listar Planilhas", f"Erro ao listar planilhas em '{os.path.basename(self.file_path)}': {e}")
            self.sheet_selector.addItem(self.DEFAULT_SHEET_NAME) 
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
            
            if ws.max_row > 1: 
                ws.delete_rows(2, ws.max_row) 

            workflow_headers = ["Tipo", "ID", "X", "Y", "Largura", "Altura", "Texto", "Cor", "Conexões"]
            current_excel_headers = [cell.value for cell in ws[1]] if ws.max_row > 0 else []

            if current_excel_headers != workflow_headers:
                if ws.max_row > 0: 
                    ws.delete_rows(1)
                ws.insert_rows(1) 
                ws.append(workflow_headers)
            elif not current_excel_headers: 
                ws.append(workflow_headers)

            for i, node_item in enumerate(self.nodes):
                node_id = f"node_{i}" 
                text_item = None
                for item in self.scene.items(node_item.boundingRect()): 
                    if isinstance(item, QGraphicsTextItem):
                        text_item = item
                        break
                
                node_text = text_item.toPlainText() if text_item else ""
                node_x = node_item.rect().x()
                node_y = node_item.rect().y()
                node_width = node_item.rect().width()
                node_height = node_item.rect().height()
                node_color = node_item.brush().color().name() 
                
                connections = [] 
                
                ws.append(["Node", node_id, node_x, node_y, node_width, node_height, node_text, node_color, json.dumps(connections)])

            for i, link_item in enumerate(self.links):
                link_id = f"link_{i}"
                ws.append(["Link", link_id, "", "", "", "", "", "", json.dumps({"source": "id_origem", "target": "id_destino"})])

            wb.save(self.file_path)
            QMessageBox.information(self, "Sucesso", f"Workflow salvo em '{current_sheet_name}' em '{os.path.basename(self.file_path)}'.")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar o workflow: {e}")

    def _load_workflow_from_excel(self):
        """
        Carrega um diagrama de fluxo de trabalho da planilha Excel selecionada.
        """
        self._clear_diagram() 
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
            
            headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
            header_map = {h: idx for idx, h in enumerate(headers)}

            loaded_nodes = {} 

            for row_idx in range(2, sheet.max_row + 1): 
                row_values = [cell.value for cell in sheet[row_idx]]
                row_type = row_values[header_map.get("Tipo")] if "Tipo" in header_map and header_map["Tipo"] < len(row_values) else None

                if row_type == "Node":
                    node_id = row_values[header_map.get("ID")] if "ID" in header_map and header_map["ID"] < len(row_values) else None
                    x = row_values[header_map.get("X")] if "X" in header_map and header_map["X"] < len(row_values) else 0
                    y = row_values[header_map.get("Y")] if "Y" in header_map and header_map["Y"] < len(row_values) else 0
                    width = row_values[header_map.get("Largura")] if "Largura" in header_map and header_map["Largura"] < len(row_values) else 100
                    height = row_values[header_map.get("Altura")] if "Altura" in header_map and header_map["Altura"] < len(row_values) else 50
                    text = row_values[header_map.get("Texto")] if "Texto" in header_map and header_map["Texto"] < len(row_values) else ""

                    text = str(text) if text is not None else "" 

                    color_name = row_values[header_map.get("Cor")] if "Cor" in header_map and header_map["Cor"] < len(row_values) else "lightblue"

                    if node_id:
                        node_rect = self.scene.addRect(x, y, width, height, QPen(Qt.black), QBrush(QColor(color_name)))
                        node_text_item = self.scene.addText(text) 
                        node_text_item.setPos(x + 5, y + 15) 
                        
                        self.nodes.append(node_rect)
                        loaded_nodes[node_id] = node_rect 

                elif row_type == "Link":
                    link_data_str = row_values[header_map.get("Conexões")] if "Conexões" in header_map and header_map["Conexões"] < len(row_values) else "{}"
                    try:
                        link_data = json.loads(link_data_str)
                        source_id = link_data.get("source")
                        target_id = link_data.get("target")

                        source_node = loaded_nodes.get(source_id)
                        target_node = loaded_nodes.get(target_id)

                        if source_node and target_node:
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
        node1 = self.scene.addRect(50, 50, 100, 50, QPen(Qt.black), QBrush(QColor("lightblue")))
        node2 = self.scene.addRect(200, 150, 100, 50, QPen(Qt.black), QBrush(QColor("lightgreen")))
        node3 = self.scene.addRect(350, 50, 100, 50, QPen(Qt.black), QBrush(QColor("lightcoral")))

        text_item1 = self.scene.addText("Fase de Design")
        text_item1.setPos(50 + 5, 50 + 15) 

        text_item2 = self.scene.addText("Revisão (Aprovado)")
        text_item2.setPos(200 + 5, 150 + 15)

        text_item3 = self.scene.addText("Preparação da Produção")
        text_item3.setPos(350 + 5, 50 + 15)

        self.nodes.extend([node1, node2, node3]) 

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
        
        node = self.scene.addRect(x, y, 100, 50, QPen(Qt.black), QBrush(QColor("#FFD700"))) 
        
        text_item = self.scene.addText(f"Nova Tarefa {len(self.nodes) + 1}")
        text_item.setPos(x + 5, y + 15) 
        
        self.nodes.append(node)
        self.view.centerOn(node)

    def _add_dependency_link(self):
        """
        Prompts user to select two nodes to link. (Conceptual, requires selection logic).
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
    os arquivos Excel nas pastas user_sheets e app_sheets, preservando descrições existentes.
    """
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
        self.table.setEditTriggers(QTableWidget.NoEditTriggers) 
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.layout.addWidget(self.table)
        
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Arquivo (Caminho)", "Nome da Coluna (Cabeçalho)", "pagina_arquivo", "descr_variavel"])

    def _load_existing_db_db_data(self):
        """
        Carrega os dados existentes da planilha 'db_db' para um dicionário de lookup.
        Retorna: um dicionário onde a chave é (caminho_relativo_arquivo, nome_coluna)
                 e o valor é {'pagina_arquivo': ..., 'descr_variavel': ...}.
        """
        existing_data = {}
        try:
            if not os.path.exists(DB_EXCEL_PATH):
                return existing_data

            wb = openpyxl.load_workbook(DB_EXCEL_PATH)
            if "db_db" not in wb.sheetnames:
                return existing_data

            sheet = wb["db_db"]
            headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
            header_map = {h: idx for idx, h in enumerate(headers)}

            # Garante que as colunas essenciais existem
            if not all(col in header_map for col in ["Arquivo (Caminho)", "Nome da Coluna (Cabeçalho)", "pagina_arquivo", "descr_variavel"]):
                print("Aviso: A planilha 'db_db' não possui todos os cabeçalhos esperados.")
                return existing_data # Não podemos carregar corretamente sem os cabeçalhos

            for row_idx in range(2, sheet.max_row + 1):
                row_values = [cell.value for cell in sheet[row_idx]]
                
                file_path_raw = row_values[header_map["Arquivo (Caminho)"]]
                column_name = row_values[header_map["Nome da Coluna (Cabeçalho)"]]
                pagina_arquivo = row_values[header_map["pagina_arquivo"]]
                descr_variavel = row_values[header_map["descr_variavel"]]

                # Use o caminho relativo normalizado como chave
                normalized_path = file_path_raw.replace('\\', '/') # Normaliza para consistência
                
                if normalized_path and column_name:
                    existing_data[(normalized_path, str(column_name))] = {
                        'pagina_arquivo': pagina_arquivo if pagina_arquivo is not None else "",
                        'descr_variavel': descr_variavel if descr_variavel is not None else ""
                    }
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Carregar DB", f"Não foi possível carregar dados existentes de db_db: {e}")
            print(f"Erro ao carregar dados existentes de db_db: {e}")
        return existing_data

    def _update_db_headers(self):
        """Coleta cabeçalhos de todos os arquivos Excel especificados e os salva em db.xlsx (planilha db_db),
        preservando informações existentes."""
        self.status_label.setText("Coletando e mesclando cabeçalhos...")
        QApplication.processEvents() 

        existing_db_data = self._load_existing_db_db_data()
        all_headers_data = [] # Para a nova lista de dados
        processed_keys = set() # Para evitar duplicatas na saída final

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
                            if sheet.max_row > 0: # Pelo menos a linha de cabeçalho existe
                                headers = [cell.value for cell in sheet[1]]
                                for header in headers:
                                    header_str = str(header) if header is not None else ""
                                    # Caminho relativo para a chave do dicionário
                                    file_path_rel_normalized = os.path.relpath(file_path, project_root).replace('\\', '/')
                                    lookup_key = (file_path_rel_normalized, header_str)

                                    if lookup_key not in processed_keys:
                                        pagina_arquivo = "obs IA olhar em runtime e salvar"
                                        descr_variavel = ""

                                        # Tenta carregar descrições existentes
                                        if lookup_key in existing_db_data:
                                            existing_entry = existing_db_data[lookup_key]
                                            pagina_arquivo = existing_entry['pagina_arquivo']
                                            descr_variavel = existing_entry['descr_variavel']
                                        
                                        all_headers_data.append([
                                            os.path.relpath(file_path, project_root), # Mantém separadores nativos para exibição
                                            header_str,
                                            pagina_arquivo,
                                            descr_variavel
                                        ])
                                        processed_keys.add(lookup_key) # Adiciona à lista de processados
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
            
            # Limpa os dados existentes em db_db
            for row_idx in range(ws_db.max_row, 0, -1): 
                ws_db.delete_rows(row_idx)

            # Adiciona cabeçalhos para a planilha db_db
            ws_db.append(["Arquivo (Caminho)", "Nome da Coluna (Cabeçalho)", "pagina_arquivo", "descr_variavel"])
            
            # Adiciona dados coletados
            for row_data in all_headers_data:
                ws_db.append(row_data)
            
            wb_db.save(DB_EXCEL_PATH)
            self.status_label.setText("Cabeçalhos atualizados e mesclados em db.xlsx (db_db).")
            QMessageBox.information(self, "Sucesso", "Cabeçalhos atualizados e mesclados em db.xlsx (db_db).")
            
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
        self.setGeometry(200, 200, 400, 300) 

        self.username = username
        self.role = role

        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("<h2>Informações do Perfil</h2>"))
        
        user_layout = QHBoxLayout()
        user_layout.addWidget(QLabel("<b>Usuário:</b>"))
        user_label = QLabel(self.username)
        user_layout.addWidget(user_label)
        user_layout.addStretch()
        layout.addLayout(user_layout)

        role_layout = QHBoxLayout()
        role_layout.addWidget(QLabel("<b>Papel:</b>"))
        role_label = QLabel(self.role)
        role_layout.addWidget(role_label)
        role_layout.addStretch()
        layout.addLayout(role_layout)

        layout.addStretch() 

        layout.addWidget(QLabel("<h3>Configurações Adicionais (Exemplo):</h3>"))
        
        email_layout = QHBoxLayout()
        email_layout.addWidget(QLabel("Email:"))
        email_input = QLineEdit("usuario@example.com") 
        email_layout.addWidget(email_input)
        layout.addLayout(email_layout)

        lang_layout = QHBoxLayout()
        lang_layout.addWidget(QLabel("Idioma:"))
        lang_input = QLineEdit("Português (Brasil)") 
        lang_layout.addWidget(lang_input)
        layout.addLayout(lang_layout)

        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(self.accept) 

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
        
        self.username = user["username"]
        self.role = user["role"]
        self.tools = load_tools_from_excel() 
        self.permissions = load_role_permissions()
        # self.workspace_items = load_workspace_items_from_excel() # Removido para aguardar "ordem 6"

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
                
                if tool["id"] == "mod4": 
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
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, FinanceiroTool(file_path=FINANCEIRO_EXCEL_PATH))) # Corrigido: FinanceiroolTool -> FinanceiroTool
                elif tool["id"] == "pedidos":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, PedidosTool(file_path=PEDIDOS_EXCEL_PATH)))
                elif tool["id"] == "manutencao":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ManutencaoTool(file_path=MANUTENCAO_EXCEL_PATH)))
                elif tool["id"] == "rpi_tool":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, RpiTool(file_path=RPI_EXCEL_PATH)))
                elif tool["id"] == "engenharia_data":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, EngenhariaDataTool(file_path=ENGENHARIA_EXCEL_PATH)))
                elif tool["id"] == "db_headers_updater": 
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, DbHeadersUpdaterTool()))
                else: 
                    action.triggered.connect(lambda chk=False, title=tool["name"], desc=tool["description"]: self._open_tab(title, QLabel(desc)))
        
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
        self.tree.clear() 
        
        # Adiciona itens de espaço de trabalho fixos
        projects_root = QTreeWidgetItem(["Projetos/Espaço de Trabalho"])
        self.tree.addTopLevelItem(projects_root)

        # Adiciona itens da lista WORKSPACE_ITEMS
        for item_name in WORKSPACE_ITEMS:
            # Lógica simples para estruturar se for "Projeto" ou "Variante"
            if "Project" in item_name or "Variant" in item_name:
                top_level_item = QTreeWidgetItem([item_name])
                projects_root.addChild(top_level_item)
                # Você pode adicionar sub-itens aqui se o WORKSPACE_ITEMS fosse mais estruturado
            else:
                # Tenta anexar a um projeto/variante existente, ou ao root dos projetos
                found_parent = False
                for i in range(projects_root.childCount()):
                    parent_item = projects_root.child(i)
                    if "Project" in parent_item.text(0) or "Variant" in parent_item.text(0):
                        if item_name.startswith(parent_item.text(0).split(' ')[0]): # Ex: Part-001 para Demo Project
                            parent_item.addChild(QTreeWidgetItem([item_name]))
                            found_parent = True
                            break
                if not found_parent: # Se não encontrou um pai adequado
                    projects_root.addChild(QTreeWidgetItem([item_name]))


        user_sheets_root = QTreeWidgetItem(["Arquivos do Usuário (user_sheets)"])
        self.tree.addTopLevelItem(user_sheets_root)

        for filename in sorted(os.listdir(USER_SHEETS_DIR)): 
            if filename.endswith('.xlsx'):
                file_path = os.path.join(USER_SHEETS_DIR, filename)
                file_item = QTreeWidgetItem([filename])
                file_item.setData(0, Qt.UserRole, file_path) 
                user_sheets_root.addChild(file_item)
        
        app_sheets_root = QTreeWidgetItem(["Arquivos do Sistema (app_sheets)"])
        self.tree.addTopLevelItem(app_sheets_root)

        for filename in sorted(os.listdir(APP_SHEETS_DIR)): 
            if filename.endswith('.xlsx'):
                file_path = os.path.join(APP_SHEETS_DIR, filename)
                file_item = QTreeWidgetItem([filename])
                file_item.setData(0, Qt.UserRole, file_path) 
                app_sheets_root.addChild(file_item)

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
        
        tool_widget = None

        if file_name == os.path.basename(ENGENHARIA_EXCEL_PATH):
            tool_widget = EngenhariaDataTool(file_path=file_path)
        elif file_name == os.path.basename(COLABORADORES_EXCEL_PATH):
             tool_widget = ColaboradoresTool(file_path=file_path)
        elif file_name == os.path.basename(OUTPUT_EXCEL_PATH):
             tool_widget = ProductDataTool(file_path=file_path)
        elif file_name == os.path.basename(BOM_DATA_EXCEL_PATH):
             tool_widget = BomManagerTool(file_path=file_path) 
        elif file_name == os.path.basename(CONFIGURADOR_EXCEL_PATH):
             tool_widget = ConfiguradorTool(file_path=file_path)
        elif file_name == os.path.basename(ESTOQUE_EXCEL_PATH): 
             tool_widget = ItemsTool(file_path=file_path)
        elif file_name == os.path.basename(ITEMS_DATA_EXCEL_PATH): 
             tool_widget = ItemsTool(file_path=file_path, read_only=True) 
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
        elif file_name == os.path.basename(ENGENHARIA_EXCEL_PATH):
             tool_widget = EngenhariaDataTool(file_path=file_path)
        elif file_name == os.path.basename(DB_EXCEL_PATH): 
            tool_widget = DbHeadersUpdaterTool() 
            tab_title = "Atualizador de Cabeçalhos do DB"
        elif file_name == os.path.basename(TOOLS_EXCEL_PATH): 
            tool_widget = ExcelViewerTool(file_path=file_path, read_only=True) 
        elif file_name == os.path.basename(MODULES_EXCEL_PATH):
            tool_widget = ExcelViewerTool(file_path=file_path, read_only=True)
        elif file_name == os.path.basename(PERMISSIONS_EXCEL_PATH):
            tool_widget = ExcelViewerTool(file_path=file_path, read_only=True)
        elif file_name == os.path.basename(ROLES_TOOLS_EXCEL_PATH):
            tool_widget = ExcelViewerTool(file_path=file_path, read_only=True)
        elif file_name == os.path.basename(USERS_EXCEL_PATH): 
            tool_widget = ExcelViewerTool(file_path=file_path, read_only=True)
        elif file_name == os.path.basename(MAIN_EXCEL_PATH):
            tool_widget = ExcelViewerTool(file_path=file_path, read_only=True)
        else: 
            tool_widget = ExcelViewerTool(file_path=file_path)

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
        self.mes_start_time_input = QLineEdit()
        self.mes_start_time_input.setPlaceholderText("Hora de Início (AAAA-MM-DD HH:MM)")
        self.mes_end_time_input = QLineEdit()
        self.mes_end_time_input.setPlaceholderText("Hora de Término (AAAA-MM-DD HH:MM)")

        submit_btn = QPushButton("Enviar Dados de Produção")
        submit_btn.clicked.connect(self._submit_mes_data) 

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
        mes_layout.addStretch() 
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

        QMessageBox.information(self, "Dados MES Enviados",
                                f"Dados de Produção Enviados:\n"
                                f"ID da Ordem: {order_id}\n"
                                f"Código do Item: {item_code}\n"
                                f"Quantidade: {quantity}\n"
                                f"Início: {start_time}\n"
                                f"Término: {end_time}")
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
        dialog.setGeometry(self.x() + 200, self.y() + 100, 400, 300) 

        layout = QVBoxLayout(dialog)
        
        if not results:
            layout.addWidget(QLabel("Nenhum item encontrado correspondente à sua pesquisa."))
        else:
            list_widget = QListWidget()
            for item in results:
                list_widget.addItem(item)
            list_widget.itemDoubleClicked.connect(
                lambda item: self.open_selected_item_tab(item.text()) or dialog.accept()
            ) 
            layout.addWidget(list_widget)

        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(dialog.accept) 
        layout.addWidget(close_btn)

        dialog.exec_() 

    def open_selected_item_tab(self, item_name):
        """
        Abre uma nova aba na GUI principal para exibir detalhes do item selecionado.
        """
        tab_title = f"Detalhes: {item_name}"

        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == tab_title:
                self.tabs.setCurrentIndex(i)
                QMessageBox.information(self, "Informação", f"A guia para '{item_name}' já está aberta.")
                return

        item_details_widget = QWidget()
        item_details_layout = QVBoxLayout()
        item_details_layout.addWidget(QLabel(f"<h2>Detalhes do Item: {item_name}</h2>"))
        item_details_layout.addWidget(QLabel(f"Exibindo detalhes abrangentes para <b>{item_name}</b>."))
        item_details_layout.addWidget(QLabel("Esta seção carregaria dados reais: propriedades, revisões, arquivos associados, etc."))
        item_details_layout.addStretch() 
        item_details_widget.setLayout(item_details_layout)

        self._open_tab(tab_title, item_details_widget)
        QMessageBox.information(self, "Item Aberto", f"Detalhes abertos para: {item_name}")


    def _open_tab(self, title, widget_instance):
        """
        Opens a new tab or switches to an existing one.
        Accepts a widget instance directly.
        """
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == title:
                self.tabs.setCurrentIndex(i)
                return
        self.tabs.addTab(widget_instance, title)
        self.tabs.setCurrentIndex(self.tabs.count() - 1) 

    def _open_options(self):
        """Abre a caixa de diálogo de opções/configurações do usuário."""
        settings_dialog = ProfileSettingsWindow(self.username, self.role)
        settings_dialog.exec_() 

    def _logout(self):
        """Faz logout do usuário atual e retorna para a tela de login."""
        confirm_logout = QMessageBox.question(self, "Confirmação de Saída", "Tem certeza de que deseja sair?",
                                              QMessageBox.Yes | QMessageBox.No)
        if confirm_logout == QMessageBox.Yes:
            self.close() 
            self.login = LoginWindow() 
            self.login.show() 

    def _show_tree_context_menu(self, pos):
        """Exibe um menu de contexto para itens na visualização em árvore."""
        item = self.tree.itemAt(pos)
        if not item: return

        menu = QMenu()
        file_path = item.data(0, Qt.UserRole)
        file_name = os.path.basename(file_path) if file_path else ""

        is_protected = file_name in PROTECTED_FILES

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
                protected_rename_action = menu.addAction("Renomear (Protegido)")
                protected_rename_action.setEnabled(False)
                protected_delete_action = menu.addAction("Excluir (Protegido)")
                protected_delete_action.setEnabled(False)

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

        if old_file_name in PROTECTED_FILES:
            QMessageBox.warning(self, "Operação Não Permitida", f"O arquivo '{old_file_name}' é protegido e não pode ser renomeado.")
            return

        new_file_name, ok = QInputDialog.getText(self, "Renomear Arquivo", f"Novo nome para '{old_file_name}':", QLineEdit.Normal, old_file_name)
        
        if ok and new_file_name and new_file_name != old_file_name:
            if old_file_name.endswith('.xlsx') and not new_file_name.endswith('.xlsx'):
                new_file_name += '.xlsx'

            new_file_path = os.path.join(os.path.dirname(old_file_path), new_file_name)

            if os.path.exists(new_file_path):
                QMessageBox.warning(self, "Erro ao Renomear", f"Já existe um arquivo com o nome '{new_file_name}'.")
                return

            try:
                os.rename(old_file_path, new_file_path)
                item.setText(0, new_file_name) 
                item.setData(0, Qt.UserRole, new_file_path) 
                QMessageBox.information(self, "Sucesso", f"Arquivo renomeado para '{new_file_name}'.")
            except Exception as e:
                QMessageBox.critical(self, "Erro ao Renomear", f"Não foi possível renomear o arquivo: {e}")

    def _delete_file(self, item):
        """Exclui um arquivo do sistema de arquivos e o remove da visualização em árvore."""
        file_path = item.data(0, Qt.UserRole)
        if not file_path: return

        file_name = os.path.basename(file_path)

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
                    self.tree.invisibleRootItem().removeChild(item) 
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
            result = subprocess.run([sys.executable, script_path], capture_output=True, text=True, check=True)
            QMessageBox.information(self, "Script Executado", 
                                    f"Script '{os.path.basename(script_path)}' executado com sucesso.\n\n"
                                    f"Saída:\n{result.stdout}\n"
                                    f"Erros (se houver):\n{result.stderr}")
            self._populate_sample_tree() 
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
