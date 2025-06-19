import sys
import os
import bcrypt
import openpyxl
import json
import subprocess
import threading 

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QAction, QTabWidget, QMenu, QToolButton,
    QWidget, QVBoxLayout, QSplitter, QTreeWidget, QTreeWidgetItem,
    QLabel, QLineEdit, QPushButton, QHBoxLayout, QMessageBox, QGraphicsView,
    QGraphicsScene, QGraphicsRectItem, QGraphicsLineItem, QDialog, QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView, QInputDialog, QComboBox, QGraphicsTextItem,
    QTextEdit 
)
from PyQt5.QtCore import Qt, QPointF, QFileInfo
from PyQt5.QtGui import QBrush, QPen, QColor, QFont 

# --- Correção para ModuleNotFoundError: No module named 'ui' ---
# Obtém o caminho absoluto do diretório contendo gui.py
current_dir = os.path.dirname(os.path.abspath(__file__))
# Navega até a raiz do projeto (assumindo gui.py está em client/, e client/ está na raiz do projeto/)
project_root = os.path.dirname(current_dir)
# Adiciona a raiz do projeto ao sys.path para que Python possa encontrar 'ui' e 'app_sheets/tools' etc.
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
from ui.tools.financeiro import FinanceiroTool 
from ui.tools.pedidos import PedidosTool
from ui.tools.manutencao import ManutencaoTool
from ui.tools.engenharia_data import EngenhariaDataTool 
from ui.tools.excel_viewer_tool import ExcelViewerTool 
from ui.tools.structure_view_tool import StructureViewTool
from ui.tools.rpi_tool import RpiTool 
from ui.tools.engenharia_workflow_tool import EngenhariaWorkflowTool 
from ui.tools.user_settings_tool import UserSettingsTool
from app_sheets.tools.tools_line_generator import ToolsLineGeneratorTool 

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
ESTOQUE_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "estoque.xlsx") 
DB_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "db.xlsx") # db.xlsx permanece em user_sheets
WORKSPACE_DATA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "workspace_data.xlsx") 

# Caminhos para arquivos Excel gerenciados pelo aplicativo (na pasta app_sheets)
USERS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "users.xlsx")
ACCESS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "access.xlsx") # Usado pelo frontend
TOOLS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "tools.xlsx") # Usado pelo frontend
# Manter as referências aos arquivos usados pelo backend (main.py)
MODULES_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "modules.xlsx") # Usado pelo backend
PERMISSIONS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "permissions.xlsx") # Usado pelo backend

ENGENHARIA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "engenharia.xlsx")
BOM_DATA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "bom_data.xlsx") 
ITEMS_DATA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "items_data.xlsx") 
MANUFACTURING_DATA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "manufacturing_data.xlsx")
MAIN_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "main.xlsx") 

# --- Caminhos para scripts externos ---
UPDATE_METADATA_SCRIPT_PATH = os.path.join(APP_SHEETS_DIR, "tools", "update_user_sheets_metadata.py")
SHEET_VALIDATOR_SCRIPT_PATH = os.path.join(APP_SHEETS_DIR, "tools", "sheet_validator_simple.py") 
CREATE_ENGENHARIA_SCRIPT_PATH = os.path.join(APP_SHEETS_DIR, "tools", "create_engenharia_xlsx.py")
TOOLS_LINE_GENERATOR_SCRIPT_PATH = os.path.join(APP_SHEETS_DIR, "tools", "tools_line_generator.py") 

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
    os.path.basename(DB_EXCEL_PATH), 
    os.path.basename(ENGENHARIA_EXCEL_PATH),
    os.path.basename(BOM_DATA_EXCEL_PATH),
    os.path.basename(ITEMS_DATA_EXCEL_PATH),
    os.path.basename(MANUFACTURING_DATA_EXCEL_PATH),
    os.path.basename(WORKSPACE_DATA_EXCEL_PATH), 
    
    # NOVOS ARQUIVOS PROTEGIDOS em app_sheets
    os.path.basename(USERS_EXCEL_PATH), 
    os.path.basename(ACCESS_EXCEL_PATH), 
    os.path.basename(TOOLS_EXCEL_PATH), 
    os.path.basename(MAIN_EXCEL_PATH), 
    os.path.basename(MODULES_EXCEL_PATH), 
    os.path.basename(PERMISSIONS_EXCEL_PATH), 
    
    # Scripts protegidos
    os.path.basename(UPDATE_METADATA_SCRIPT_PATH),
    os.path.basename(SHEET_VALIDATOR_SCRIPT_PATH), 
    os.path.basename(CREATE_ENGENHARIA_SCRIPT_PATH),
    os.path.basename(TOOLS_LINE_GENERATOR_SCRIPT_PATH) 
]

# Garante que os diretórios existam
os.makedirs(USER_SHEETS_DIR, exist_ok=True)
os.makedirs(APP_SHEETS_DIR, exist_ok=True)
os.makedirs(os.path.dirname(UPDATE_METADATA_SCRIPT_PATH), exist_ok=True) 
os.makedirs(os.path.dirname(CREATE_ENGENHARIA_SCRIPT_PATH), exist_ok=True) 
os.makedirs(os.path.dirname(SHEET_VALIDATOR_SCRIPT_PATH), exist_ok=True) 
os.makedirs(os.path.dirname(TOOLS_LINE_GENERATOR_SCRIPT_PATH), exist_ok=True) 


# === FUNÇÕES AUXILIARES DE PLANILHA ===
def load_users_from_excel_util():
    """Carrega dados de usuário do arquivo users.xlsx (agora em app_sheets)."""
    users = {}
    try:
        if not os.path.exists(USERS_EXCEL_PATH):
            QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo de usuários não foi encontrado: {USERS_EXCEL_PATH}")
            return {}

        wb = openpyxl.load_workbook(USERS_EXCEL_PATH)
        users_sheet = wb["users"] 
        
        headers = [cell.value for cell in users_sheet[1]] if users_sheet.max_row >= 1 else []
        header_map = {h: idx for idx, h in enumerate(headers)}

        required_headers = ["id", "username", "password_hash", "role"]
        if not all(h in header_map for h in required_headers):
            QMessageBox.warning(None, "Cabeçalhos Ausentes", 
                                f"A planilha 'users' em {USERS_EXCEL_PATH} não possui todos os cabeçalhos esperados. "
                                f"Esperado: {', '.join(required_headers)}")
            return {}

        for row_idx in range(2, users_sheet.max_row + 1):
            row_values = [cell.value for cell in users_sheet[row_idx]]
            
            if all(v is None for v in row_values):
                continue

            user_id = row_values[header_map["id"]] if "id" in header_map and header_map["id"] < len(row_values) else None
            username = row_values[header_map["username"]] if "username" in header_map and header_map["username"] < len(row_values) else None
            password_hash = row_values[header_map["password_hash"]] if "password_hash" in header_map and header_map["password_hash"] < len(row_values) else None
            role = row_values[header_map["role"]] if "role" in header_map and header_map["role"] < len(row_values) else None

            if username is not None and password_hash is not None:
                users[str(username)] = { 
                    "id": user_id,
                    "username": str(username),
                    "password_hash": str(password_hash),
                    "role": str(role) if role is not None else "user"
                }
            else:
                print(f"Aviso: Ignorando linha malformada na planilha 'users' (linha {row_idx}): {row_values}")

    except KeyError:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha 'users' não foi encontrada em {USERS_EXCEL_PATH}. Por favor, certifique-se de que o nome da planilha seja 'users'.")
        return {}
    except Exception as e:
        QMessageBox.critical(None, "Erro de Carregamento", f"Erro ao carregar usuários: {e}")
        return {}
    return users


def register_user(username, password, role="user"):
    """Registra um novo usuário no arquivo users.xlsx (agora em app_sheets)."""
    try:
        wb = openpyxl.load_workbook(USERS_EXCEL_PATH)
        sheet = wb["users"]
        
        headers = [cell.value for cell in sheet[1]] if sheet.max_row >= 1 else []
        header_map = {h: idx for idx, h in enumerate(headers)}

        required_headers = ["id", "username", "password_hash", "role"]
        if not all(h in header_map for h in required_headers):
            for h in required_headers:
                if h not in header_map:
                    headers.append(h)
            sheet.insert_rows(1) 
            for col_idx, h in enumerate(headers):
                sheet.cell(row=1, column=col_idx + 1).value = h
            header_map = {h: idx for idx, h in enumerate(headers)} 

        next_id = 1
        existing_ids = set()
        for row_idx in range(2, sheet.max_row + 1):
            row_values = [cell.value for cell in sheet[row_idx]]
            current_id = row_values[header_map["id"]] if "id" in header_map and header_map["id"] < len(row_values) else None
            current_username = row_values[header_map["username"]] if "username" in header_map and header_map["username"] < len(row_values) else None

            if all(v is None for v in row_values):
                continue

            if current_id is not None:
                existing_ids.add(current_id)
            if current_username == username:
                raise ValueError("Nome de usuário já existe.")
        while next_id in existing_ids:
            next_id += 1

        password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        
        new_row_data = [""] * len(headers)
        new_row_data[header_map["id"]] = next_id
        new_row_data[header_map["username"]] = username
        new_row_data[header_map["password_hash"]] = password_hash
        new_row_data[header_map["role"]] = role
        
        sheet.append(new_row_data)
        wb.save(USERS_EXCEL_PATH)
    except FileNotFoundError:
        QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo de usuários não foi encontrado em: {USERS_EXCEL_PATH}. Não é possível registrar o usuário.")
    except KeyError:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha 'users' não foi encontrada em {USERS_EXCEL_PATH}. Não é possível registrar o usuário.")
    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Ocorreu um erro durante o registro: {e}")


def load_tools_from_excel_util():
    """
    Carrega dados da ferramenta do arquivo Excel dedicado às ferramentas (tools.xlsx).
    Adaptado para os novos cabeçalhos: mod_id, mod_name, module_path, etc.
    """
    tools = {}
    try:
        if not os.path.exists(TOOLS_EXCEL_PATH):
            QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo de ferramentas não foi encontrado em: {TOOLS_EXCEL_PATH}. Por favor, certifique-se de que ele exista.")
            return {}

        wb = openpyxl.load_workbook(TOOLS_EXCEL_PATH)
        sheet = wb["tools"] 
        
        headers = [cell.value for cell in sheet[1]] if sheet.max_row >= 1 else []
        header_map = {h: idx for idx, h in enumerate(headers)}

        required_headers = ["mod_id", "mod_name", "module_path"] 
        if not all(h in header_map for h in required_headers):
            QMessageBox.warning(None, "Cabeçalhos Ausentes", 
                                f"A planilha 'tools' em {TOOLS_EXCEL_PATH} não possui todos os cabeçalhos esperados. "
                                f"Esperado: {', '.join(required_headers)}")
            return {}

        for row_idx in range(2, sheet.max_row + 1):
            row_values = [cell.value for cell in sheet[row_idx]]
            
            if all(v is None for v in row_values):
                continue

            mod_id = row_values[header_map["mod_id"]] if "mod_id" in header_map and header_map["mod_id"] < len(row_values) and row_values[header_map["mod_id"]] is not None else None
            mod_name = row_values[header_map["mod_name"]] if "mod_name" in header_map and header_map["mod_name"] < len(row_values) and row_values[header_map["mod_name"]] is not None else None
            mod_description = row_values[header_map.get("mod_description")] if "mod_description" in header_map and header_map["mod_description"] < len(row_values) else None
            module_path = row_values[header_map["module_path"]] if "module_path" in header_map and header_map["module_path"] < len(row_values) and row_values[header_map["module_path"]] is not None else None
            mod_work_table = row_values[header_map.get("MOD_WORK_TABLE")] if "MOD_WORK_TABLE" in header_map and header_map["MOD_WORK_TABLE"] < len(row_values) else None
            mod_work_table_path = row_values[header_map.get("MOD_WORK_TABLE_PATH")] if "MOD_WORK_TABLE_PATH" in header_map and header_map["MOD_WORK_TABLE_PATH"] < len(row_values) else None

            if mod_id is not None and mod_name is not None and module_path is not None:
                tools[str(mod_id)] = {
                    "id": str(mod_id),
                    "name": str(mod_name),
                    "description": str(mod_description) if mod_description is not None else "",
                    "path": str(module_path), 
                    "mod_work_table": str(mod_work_table) if mod_work_table is not None else "",
                    "mod_work_table_path": str(mod_work_table_path) if mod_work_table_path is not None else ""
                }
            else:
                print(f"Aviso: Ignorando linha malformada ou incompleta na planilha 'tools' (linha {row_idx}): {row_values}")

    except KeyError as ke:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha 'tools' não foi encontrada em {TOOLS_EXCEL_PATH} ou cabeçalho ausente ({ke}). Por favor, certifique-se de que o nome da planilha seja 'tools' e os cabeçalhos estejam corretos.")
        return {}
    except Exception as e:
        QMessageBox.critical(None, "Erro de Carregamento", f"Erro ao carregar ferramentas: {e}")
        return {}
    return tools


def load_role_permissions_util():
    """Carrega permissões de papel do arquivo access.xlsx (agora em app_sheets)."""
    perms = {}
    try:
        if not os.path.exists(ACCESS_EXCEL_PATH):
            QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo de permissões não foi encontrado: {ACCESS_EXCEL_PATH}")
            return {}

        wb = openpyxl.load_workbook(ACCESS_EXCEL_PATH)
        sheet = wb["access"] 
        
        headers = [cell.value for cell in sheet[1]] if sheet.max_row >= 1 else []
        header_map = {h: idx for idx, h in enumerate(headers)}

        required_headers = ["role", "allowed_tools"]
        if not all(h in header_map for h in required_headers):
            QMessageBox.warning(None, "Cabeçalhos Ausentes", 
                                f"A planilha 'access' em {ACCESS_EXCEL_PATH} não possui todos os cabeçalhos esperados. "
                                f"Esperado: {', '.join(required_headers)}")
            return {}

        for row_idx in range(2, sheet.max_row + 1):
            row_values = [cell.value for cell in sheet[row_idx]]
            
            if all(v is None for v in row_values):
                continue

            role_name = row_values[header_map["role"]] if "role" in header_map and header_map["role"] < len(row_values) and row_values[header_map["role"]] is not None else None
            allowed_tools_str = row_values[header_map["allowed_tools"]] if "allowed_tools" in header_map and header_map["allowed_tools"] < len(row_values) and row_values[header_map["allowed_tools"]] is not None else ""

            if role_name:
                perms[str(role_name)] = [s.strip() for s in allowed_tools_str.split(',')] if allowed_tools_str.strip().lower() != "all" else "all"
            else:
                print(f"Aviso: Ignorando linha malformada na planilha 'access' (linha {row_idx}): {row_values}")
        return perms
    except KeyError:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha 'access' não foi encontrada em {ACCESS_EXCEL_PATH}. Por favor, certifique-se de que o nome da planilha seja 'access'.")
        return {}
    except Exception as e:
        QMessageBox.critical(None, "Erro de Carregamento", f"Erro ao carregar permissões: {e}")
        return {}
    return perms


def load_workspace_items_from_excel_util():
    """
    Carrega os itens do espaço de trabalho do arquivo workspace_data.xlsx.
    """
    workspace_items = []
    try:
        if not os.path.exists(WORKSPACE_DATA_EXCEL_PATH): 
            QMessageBox.warning(None, "Arquivo Não Encontrado", 
                                f"O arquivo de itens do espaço de trabalho não foi encontrado: {WORKSPACE_DATA_EXCEL_PATH}\n"
                                "Por favor, crie-o com uma planilha 'WorkspaceItems' e as colunas 'Name' e 'Type'.")
            return []

        wb = openpyxl.load_workbook(WORKSPACE_DATA_EXCEL_PATH) 
        sheet_name = "WorkspaceItems"
        if sheet_name not in wb.sheetnames:
            QMessageBox.warning(None, "Planilha Ausente", 
                                f"A planilha '{sheet_name}' não foi encontrada em {WORKSPACE_DATA_EXCEL_PATH}. "
                                f"Por favor, certifique-se de que a planilha exista e tenha as colunas 'Name' e 'Type'.")
            return []

        sheet = wb[sheet_name]
        
        headers = [cell.value for cell in sheet[1]] if sheet.max_row >= 1 else []
        header_map = {h: idx for idx, h in enumerate(headers)}

        required_headers = ["Name", "Type"]
        if not all(h in header_map for h in required_headers):
            QMessageBox.warning(None, "Cabeçalhos Ausentes", 
                                f"A planilha '{sheet_name}' em {WORKSPACE_DATA_EXCEL_PATH} não possui todos os cabeçalhos esperados. "
                                f"Esperado: {', '.join(required_headers)}")
            return []

        for row_idx in range(2, sheet.max_row + 1):
            row_values = [cell.value for cell in sheet[row_idx]]
            
            if all(v is None for v in row_values):
                continue 

            name = row_values[header_map["Name"]] if "Name" in header_map and header_map["Name"] < len(row_values) else None
            item_type = row_values[header_map["Type"]] if "Type" in header_map and header_map["Type"] < len(row_values) else None

            if name is not None and item_type is not None:
                workspace_items.append({"name": str(name), "type": str(item_type)})
            else:
                print(f"Aviso: Ignorando linha malformada na planilha '{sheet_name}' (linha {row_idx}): {row_values}")

    except Exception as e:
        QMessageBox.critical(None, "Erro de Carregamento", f"Erro ao carregar itens do espaço de trabalho: {e}")
        return []
    return workspace_items


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
        self.users = load_users_from_excel_util() 

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
            self.users = load_users_from_excel_util() 
            self.username_input.clear()
            self.password_input.clear()
        except ValueError as ve:
            QMessageBox.warning(self, "Falha no Registro", str(ve))
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro durante o registro: {e}")


# === NOVA FERRAMENTA: ATUALIZADOR DE CABEÇALHOS DO BD (AGORA FUNCIONAL) ===
class DbHeadersUpdaterTool(QWidget):
    """
    Ferramenta GUI para executar e exibir o resultado do validador de planilhas.
    Permite que o usuário visualize a saída do script 'sheet_validator_simple.py'
    diretamente na interface.
    """
    def __init__(self, refresh_callback=None): # Adicionado refresh_callback
        super().__init__()
        self.setWindowTitle("Validador de Consistência de Planilhas")
        self.script_path = SHEET_VALIDATOR_SCRIPT_PATH 
        self.refresh_callback = refresh_callback # Armazena o callback
        self._init_ui()

    def _init_ui(self):
        """Inicializa os elementos da interface do usuário."""
        layout = QVBoxLayout()

        description_label = QLabel(
            "Esta ferramenta executa o script de validação de consistência das planilhas.\n"
            "Ele compara os cabeçalhos das planilhas com o esquema registrado em 'db.xlsx'."
        )
        layout.addWidget(description_label)

        self.run_button = QPushButton("Executar Validação")
        self.run_button.clicked.connect(self._run_validation_script)
        layout.addWidget(self.run_button)

        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        self.output_text.setPlaceholderText("Aguardando execução da validação... Clique em 'Executar Validação' para iniciar.")
        layout.addWidget(self.output_text)

        self.setLayout(layout)

    def _run_validation_script(self):
        """
        Executa o script de validação externa e exibe a saída no QTextEdit.
        """
        self.output_text.clear()
        self.output_text.append("Executando validação... Por favor, aguarde.")
        self.run_button.setEnabled(False) 

        cmd = [sys.executable, self.script_path, "run"] 

        try:
            process = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, bufsize=1)
            
            def read_stream(stream, is_stderr=False):
                for line in stream:
                    self.output_text.append(f"ERRO: {line.strip()}" if is_stderr else line.strip())
                    QApplication.processEvents() 

            stdout_thread = threading.Thread(target=read_stream, args=(process.stdout,))
            stderr_thread = threading.Thread(target=read_stream, args=(process.stderr, True))

            stdout_thread.start()
            stderr_thread.start()

            stdout_thread.join() 
            stderr_thread.join()

            process.wait() 

            if process.returncode == 0:
                self.output_text.append("\n--- Validação Concluída com Sucesso ---")
                if self.refresh_callback: # Chama o callback para atualizar a GUI se fornecido
                    self.refresh_callback()
            else:
                self.output_text.append(f"\n--- Validação Concluída com Erros (Código: {process.returncode}) ---")
                
        except FileNotFoundError:
            self.output_text.append(f"ERRO: O executável Python ou o script '{os.path.basename(self.script_path)}' não foi encontrado. Verifique o PATH ou o caminho do script.")
        except Exception as e:
            self.output_text.append(f"ERRO Inesperado ao executar script: {e}")
        finally:
            self.run_button.setEnabled(True) 


# === CLASSE PRINCIPAL DA GUI ===
class TeamcenterStyleGUI(QMainWindow):
    def __init__(self, user_data):
        super().__init__()
        self.user_data = user_data
        self.current_user_role = user_data["role"]

        # Atributos para armazenar dados de configuração carregados
        self.users = {}
        self.access_permissions = {}
        self.available_tools_metadata = {}
        self.workspace_items = []

        self._load_all_configuration_data() # Carrega todos os dados na inicialização

        self.setWindowTitle("5revolution ERP")
        self.setGeometry(100, 100, 1200, 800) 
        self._init_ui()

    def _load_all_configuration_data(self):
        """Carrega todos os dados de configuração dos arquivos Excel."""
        self.users = load_users_from_excel_util()
        self.access_permissions = load_role_permissions_util()
        self.available_tools_metadata = load_tools_from_excel_util()
        self.workspace_items = load_workspace_items_from_excel_util()
        print("Dados de configuração carregados/recarregados.")

    def _refresh_gui_data(self):
        """Recarrega os dados de configuração e atualiza os componentes da GUI."""
        self._load_all_configuration_data()
        self._populate_tools_menu(self.findChild(QToolButton, "tools_menu_btn").menu()) # Repopula o menu de ferramentas
        self._populate_workspace_tree() # Repopula a árvore do espaço de trabalho
        self._populate_file_system_tree() # Repopula a árvore do sistema de arquivos
        print("GUI atualizada com novos dados de configuração.")


    def _init_ui(self):
        """Inicializa os componentes principais da interface do usuário."""
        self._create_toolbar_menu()
        self._create_main_layout()
        self._setup_initial_content()

    def _create_toolbar_menu(self):
        """Cria a barra de ferramentas superior e seus menus."""
        toolbar = self.addToolBar("Main Toolbar")
        toolbar.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)

        # Menu Arquivo
        file_menu_btn = QToolButton(self)
        file_menu_btn.setText("Arquivo")
        file_menu_btn.setPopupMode(QToolButton.InstantPopup)
        file_menu = QMenu(self)
        exit_action = QAction("Sair", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        file_menu_btn.setMenu(file_menu)
        toolbar.addWidget(file_menu_btn)

        # Menu Ferramentas (dinâmico com base nas permissões)
        tools_menu_btn = QToolButton(self)
        tools_menu_btn.setObjectName("tools_menu_btn") # Adiciona um nome de objeto para encontrar mais tarde
        tools_menu_btn.setText("Ferramentas")
        tools_menu_btn.setPopupMode(QToolButton.InstantPopup)
        tools_menu = QMenu(self)
        self._populate_tools_menu(tools_menu)
        tools_menu_btn.setMenu(tools_menu)
        toolbar.addWidget(tools_menu_btn)

        # Menu Ajuda
        help_menu_btn = QToolButton(self)
        help_menu_btn.setText("Ajuda")
        help_menu_btn.setPopupMode(QToolButton.InstantPopup)
        help_menu = QMenu(self)
        about_action = QAction("Sobre", self)
        about_action.triggered.connect(lambda: QMessageBox.information(self, "Sobre 5revolution", "Sistema ERP 5revolution v1.0"))
        help_menu.addAction(about_action)
        help_menu_btn.setMenu(help_menu)
        toolbar.addWidget(help_menu_btn)

        # Menu Admin (apenas para administradores)
        if self.current_user_role == "admin":
            admin_menu_btn = QToolButton(self)
            admin_menu_btn.setText("👑 Ferramentas Admin")
            admin_menu_btn.setPopupMode(QToolButton.InstantPopup)
            admin_menu = QMenu(self)
            
            # Ação para criar/reinicializar/atualizar planilhas (texto personalizado)
            create_update_sheets_action = QAction("Criar/Reinicializar/atualizar planilhas", self)
            create_update_sheets_action.setToolTip("Cria ou reinicializa os arquivos '.xlsx' na pasta 'user_sheets' com as estruturas em db_db. sem perda de dados das demais linhas>1(linhas maior que 1 do cabeçalho utilizadas no app)")
            create_update_sheets_action.triggered.connect(self._run_create_or_update_all_sheets) 
            admin_menu.addAction(create_update_sheets_action)

            # Ação para sincronizar o esquema do banco de dados (db_db) (texto personalizado)
            sync_db_schema_action = QAction("Sincronizar pagina db_db com planilhas das pastas", self)
            sync_db_schema_action.setToolTip("Coleta os cabeçalhos de todas as planilhas do projeto (user_sheets e app_sheets, exceto o próprio db.xlsx) e os registra na planilha 'db_db' em 'db.xlsx'. Esta ação reconstrói o dicionário de dados do sistema, que é a base para a validação de consistência.")
            sync_db_schema_action.triggered.connect(self._run_sync_db_db_schema) 
            admin_menu.addAction(sync_db_schema_action)

            # Ação para validar consistência do DB (texto personalizado)
            validate_db_consistency_action = QAction("Verifica planilhas retornando diferenças com db_db", self)
            validate_db_consistency_action.setToolTip("Compara a estrutura real dos cabeçalhos das planilhas do projeto com o esquema registrado na planilha 'db_db' em 'db.xlsx', identificando inconsistências e erros.")
            validate_db_consistency_action.triggered.connect(self._run_validate_db_consistency)
            admin_menu.addAction(validate_db_consistency_action)

            # NOVO: Ação para o Gerador de Linha de Ferramentas
            generate_tool_line_action = QAction("Adicionar Nova Ferramenta ao Sistema", self)
            generate_tool_line_action.setToolTip("Abre uma ferramenta para adicionar novas entradas de módulo/ferramenta na planilha tools.xlsx.")
            # Passa o callback de refresh para a ferramenta
            generate_tool_line_action.triggered.connect(lambda: self._open_tool("MOD000019", refresh_callback=self._refresh_gui_data)) 
            admin_menu.addAction(generate_tool_line_action)


            admin_menu_btn.setMenu(admin_menu)
            toolbar.addWidget(admin_menu_btn)


    def _populate_tools_menu(self, menu):
        """
        Popula o menu de ferramentas com base nas permissões do usuário e nos metadados das ferramentas.
        Utiliza 'allowed_tools' do access.xlsx diretamente.
        """
        menu.clear() # Limpa o menu antes de repopular
        user_allowed_tools_list = self.access_permissions.get(self.current_user_role, [])
        
        allowed_tool_ids_for_user = set()
        if user_allowed_tools_list == "all":
            allowed_tool_ids_for_user = set(self.available_tools_metadata.keys())
        else:
            allowed_tool_ids_for_user = set(user_allowed_tools_list)
            
        for tool_id, tool_info in self.available_tools_metadata.items():
            if tool_id in ["MOD000019", "MOD000018"]:
                continue 
            
            if tool_id in allowed_tool_ids_for_user:
                action = QAction(tool_info["name"], self)
                # Passa o callback de refresh para as ferramentas se elas precisarem
                action.triggered.connect(lambda checked, t_id=tool_id: self._open_tool(t_id, refresh_callback=self._refresh_gui_data))
                menu.addAction(action)

    def _open_tool(self, tool_id, refresh_callback=None):
        """Abre a ferramenta selecionada em uma nova aba."""
        tool_info = self.available_tools_metadata.get(tool_id)
        
        tool_name = tool_info["name"] if tool_info and "name" in tool_info else f"Ferramenta desconhecida (ID: {tool_id})"
        
        tool_classes = {
            "MOD000001": ProductDataTool,
            "MOD000002": BomManagerTool,
            "MOD000003": ConfiguradorTool,
            "MOD000004": ColaboradoresTool,
            "MOD000005": ItemsTool,
            "MOD000006": ManufacturingTool,
            "MOD000007": PcpTool,
            "MOD000008": EstoqueTool,
            "MOD000009": FinanceiroTool,
            "MOD000010": PedidosTool,
            "MOD000011": ManutencaoTool,
            "MOD000012": EngenhariaDataTool, 
            "MOD000013": EngenhariaWorkflowTool, 
            "MOD000014": UserSettingsTool, 
            "MOD000015": ExcelViewerTool, 
            "MOD000016": StructureViewTool, 
            "MOD000017": RpiTool, 
            "MOD000018": DbHeadersUpdaterTool, 
            "MOD000019": ToolsLineGeneratorTool, 
            
            # IDs antigos para compatibilidade
            "mod4": EngenhariaDataTool, 
            "mes_pcp": PcpTool, 
            "prod_data": ProductDataTool, 
            "bom_manager": BomManagerTool, 
            "configurador": ConfiguradorTool, 
            "colab": ColaboradoresTool, 
            "items_tool": ItemsTool, 
            "manuf": ManufacturingTool, 
            "pcp_tool": PcpTool, 
            "estoque_tool": EstoqueTool, 
            "financeiro": FinanceiroTool, 
            "pedidos": PedidosTool, 
            "manutencao": ManutencaoTool, 
            "mod_user_settings": UserSettingsTool, 
            "mod_workflow": EngenhariaWorkflowTool, 
            "mod_excel_viewer": ExcelViewerTool, 
            "mod_structure_view": StructureViewTool, 
            "mod_rpi_control": RpiTool, 
        }

        ToolClass = tool_classes.get(tool_id)
        if ToolClass:
            try:
                tool_instance = None
                
                if tool_id == "MOD000019": 
                    tool_instance = ToolClass(refresh_callback=refresh_callback) # Passa o callback
                elif tool_id == "MOD000018": 
                    tool_instance = ToolClass(refresh_callback=refresh_callback) # Passa o callback
                elif tool_id == "MOD000014" or tool_id == "mod_user_settings": 
                    tool_instance = ToolClass(self.user_data) 
                elif tool_id == "MOD000012" or tool_id == "mod4": 
                    tool_instance = ToolClass(file_path=ENGENHARIA_EXCEL_PATH, sheet_name="Estrutura")
                elif tool_id == "MOD000013" or tool_id == "mod_workflow": 
                    tool_instance = ToolClass(file_path=ENGENHARIA_EXCEL_PATH, sheet_name="Workflows")
                elif tool_id == "MOD000015" or tool_id == "mod_excel_viewer": 
                     tool_instance = ToolClass(file_path=None) 
                else:
                    work_table_path = tool_info.get("mod_work_table_path") if tool_info else None
                    if work_table_path:
                        full_work_table_path = os.path.normpath(os.path.join(project_root, work_table_path.strip('/\\')))
                        if os.path.exists(full_work_table_path):
                            tool_instance = ToolClass(file_path=full_work_table_path)
                        else:
                            QMessageBox.warning(self, "Caminho Inválido", f"O arquivo de trabalho para '{tool_name}' não foi encontrado: {full_work_table_path}")
                            return
                    else:
                        tool_instance = ToolClass() 

                if tool_instance:
                    self.central_widget.addTab(tool_instance, tool_name)
                    self.central_widget.setCurrentWidget(tool_instance)
                else:
                    QMessageBox.warning(self, "Erro de Instanciação", f"Não foi possível criar uma instância para a ferramenta '{tool_name}'.")

            except Exception as e:
                QMessageBox.critical(self, "Erro ao Abrir Ferramenta", f"Não foi possível abrir a ferramenta '{tool_name}': {e}")
        else:
            QMessageBox.warning(self, "Ferramenta Não Implementada", f"A ferramenta '{tool_name}' (ID: {tool_id}) ainda não tem uma classe associada ou não está implementada.")


    def _create_main_layout(self):
        """Cria o layout principal com a árvore de navegação e as abas de trabalho."""
        main_splitter = QSplitter(Qt.Horizontal)
        self.setCentralWidget(main_splitter)

        # Painel Esquerdo (Árvore de Navegação)
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        self.tree_widget = QTreeWidget()
        self.tree_widget.setHeaderLabels(["Nome", "Tipo"]) 
        self.tree_widget.itemDoubleClicked.connect(self._on_tree_item_double_clicked)
        left_layout.addWidget(QLabel("<h2>Espaço de Trabalho</h2>"))
        left_layout.addWidget(self.tree_widget)
        main_splitter.addWidget(left_panel)

        # Painel Direito (Abas de Trabalho)
        self.central_widget = QTabWidget()
        self.central_widget.setTabsClosable(True)
        self.central_widget.tabCloseRequested.connect(self.central_widget.removeTab)
        main_splitter.addWidget(self.central_widget)

        # Define o tamanho inicial dos painéis
        main_splitter.setSizes([300, 900]) 

    def _setup_initial_content(self):
        """Popula a árvore de navegação e pode abrir abas iniciais."""
        self._populate_workspace_tree()
        self._populate_file_system_tree()

    def _populate_workspace_tree(self):
        """Popula a seção 'Espaço de Trabalho' da árvore lendo de workspace_data.xlsx."""
        self.tree_widget.clear() 
        
        workspace_root_item = QTreeWidgetItem(self.tree_widget, ["Projetos/Espaço de Trabalho", "Pasta"])
        workspace_root_item.setExpanded(True) 
        
        # Usa os dados já carregados
        for item_data in self.workspace_items:
            QTreeWidgetItem(workspace_root_item, [item_data["name"], item_data["type"]])

    def _populate_file_system_tree(self):
        """Popula as seções 'Arquivos do Usuário' e 'Arquivos do Sistema' da árvore."""
        # Remove os itens de arquivos existentes para repopular
        # Note: Isso pode ser refinado para atualizar apenas os itens modificados,
        # mas para simplicidade, limpamos e repopulamos.
        # Encontra os itens raiz existentes para user_sheets e app_sheets e os remove
        root_items_to_remove = []
        for i in range(self.tree_widget.topLevelItemCount()):
            item = self.tree_widget.topLevelItem(i)
            if item.text(0) in ["Arquivos do Usuário (user_sheets)", "Arquivos do Sistema (app_sheets)"]:
                root_items_to_remove.append(item)
        
        for item in root_items_to_remove:
            self.tree_widget.takeTopLevelItem(self.tree_widget.indexOfTopLevelItem(item))

        # Adiciona os itens raiz novamente
        user_files_root = QTreeWidgetItem(self.tree_widget, ["Arquivos do Usuário (user_sheets)", "Pasta"])
        user_files_root.setExpanded(True)
        app_files_root = QTreeWidgetItem(self.tree_widget, ["Arquivos do Sistema (app_sheets)", "Pasta"])
        app_files_root.setExpanded(True)

        self._add_files_to_tree(USER_SHEETS_DIR, user_files_root)
        self._add_files_to_tree(APP_SHEETS_DIR, app_files_root)

    def _add_files_to_tree(self, directory, parent_item):
        """Adiciona arquivos .xlsx de um diretório à árvore."""
        try:
            for filename in os.listdir(directory):
                if filename.endswith(".xlsx") and not filename.startswith('~$'): 
                    file_path = os.path.join(directory, filename)
                    file_info = QFileInfo(file_path)
                    item = QTreeWidgetItem(parent_item, [file_info.fileName(), "Arquivo Excel"])
                    item.setData(0, Qt.UserRole, file_path) 
        except Exception as e:
            QMessageBox.warning(self, "Erro ao Listar Arquivos", f"Não foi possível listar arquivos em {directory}: {e}")

    def _on_tree_item_double_clicked(self, item, column):
        """Lida com o clique duplo em um item da árvore."""
        file_path = item.data(0, Qt.UserRole) 
        if file_path and os.path.exists(file_path):
            self._open_excel_file_in_viewer(file_path)
        elif not file_path:
            pass
        else:
            QMessageBox.warning(self, "Arquivo Não Encontrado", f"O arquivo '{os.path.basename(file_path)}' não existe ou o caminho está incorreto.")

    def _open_excel_file_in_viewer(self, file_path):
        """Abre um arquivo Excel usando o ExcelViewerTool."""
        tool_name = f"Viewer: {os.path.basename(file_path)}"
        
        for i in range(self.central_widget.count()):
            if self.central_widget.tabText(i) == tool_name:
                self.central_widget.setCurrentIndex(i)
                return

        try:
            excel_viewer_tool = ExcelViewerTool(file_path=file_path)
            self.central_widget.addTab(excel_viewer_tool, tool_name)
            self.central_widget.setCurrentWidget(excel_viewer_tool)
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Abrir Arquivo", f"Não foi possível abrir '{os.path.basename(file_path)}' no visualizador: {e}")
            
    # --- NOVAS FUNÇÕES PARA EXECUTAR SCRIPTS EXTERNOS (USADAS PELO MENU ADMIN) ---
    def _run_external_python_script(self, script_path, action, *args):
        """
        Executa um script Python externo em um processo separado com uma ação específica.
        Exibe uma caixa de mensagem com o resultado.
        """
        cmd = [sys.executable, script_path, action] + list(args)
        
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            output = result.stdout.strip()
            error = result.stderr.strip()

            if result.returncode == 0:
                QMessageBox.information(self, "Sucesso na Execução do Script", f"Script executado com sucesso:\n\n{output}")
                print(f"Sucesso: {output}")
                self._refresh_gui_data() # Atualiza a GUI após sucesso
            else:
                QMessageBox.critical(self, "Erro na Execução do Script", f"O script retornou um erro (Código: {result.returncode}):\n\n{error}\n{output}")
                print(f"Erro: {error}\nOutput: {output}")
        except FileNotFoundError:
            QMessageBox.critical(self, "Erro de Arquivo", f"O executável Python ou o script '{os.path.basename(script_path)}' não foi encontrado. Verifique o PATH ou o caminho do script.")
            print(f"Erro: Python executable or script '{script_path}' not found.")
        except subprocess.CalledProcessError as e:
            QMessageBox.critical(self, "Erro no Processo", f"O script '{os.path.basename(script_path)}' falhou:\n\n{e.stdout}\n{e.stderr}")
            print(f"Erro no processo: {e.stdout}\n{e.stderr}")
        except Exception as e:
            QMessageBox.critical(self, "Erro Inesperado", f"Ocorreu um erro inesperado ao tentar executar o script '{os.path.basename(script_path)}': {e}")
            print(f"Erro inesperado ao executar script: {e}")

    def _run_create_or_update_all_sheets(self):
        """Executa o script para criar/atualizar todas as planilhas definidas no db_db."""
        if not os.path.exists(UPDATE_METADATA_SCRIPT_PATH): 
            QMessageBox.critical(self, "Erro", f"O script de atualização de metadados não foi encontrado em: {UPDATE_METADATA_SCRIPT_PATH}")
            return
        self._run_external_python_script(UPDATE_METADATA_SCRIPT_PATH, "create_or_update_sheets")
        # _refresh_gui_data() é chamado dentro de _run_external_python_script()

    def _run_sync_db_db_schema(self):
        """Executa o script para sincronizar o schema db_db com os cabeçalhos reais dos arquivos."""
        if not os.path.exists(UPDATE_METADATA_SCRIPT_PATH):
            QMessageBox.critical(self, "Erro", f"O script de atualização de metadados não foi encontrado em: {UPDATE_METADATA_SCRIPT_PATH}")
            return
        self._run_external_python_script(UPDATE_METADATA_SCRIPT_PATH, "update_db_schema")
        # _refresh_gui_data() é chamado dentro de _run_external_python_script()

    def _run_validate_db_consistency(self):
        """
        Abre a ferramenta DbHeadersUpdaterTool em uma aba para validar a consistência do DB.
        Não chama _run_external_python_script diretamente para este, pois a ferramenta já gerencia a execução.
        """
        self._open_tool("MOD000018", refresh_callback=self._refresh_gui_data) # Abre a ferramenta do validador


# === INÍCIO DO APLICATIVO ===
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    login_window = LoginWindow()
    login_window.show()
    
    sys.exit(app.exec_())

