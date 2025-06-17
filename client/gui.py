import sys
import os
import bcrypt
import openpyxl
import json # Necessário para EngenhariaWorkflowTool (salvar/carregar JSON)
import subprocess # Necessário para _run_create_engenharia_script, e agora para o atualizador de metadados

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QAction, QTabWidget, QMenu, QToolButton,
    QWidget, QVBoxLayout, QSplitter, QTreeWidget, QTreeWidgetItem,
    QLabel, QLineEdit, QPushButton, QHBoxLayout, QMessageBox, QGraphicsView,
    QGraphicsScene, QGraphicsRectItem, QGraphicsLineItem, QDialog, QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView, QInputDialog, QComboBox, QGraphicsTextItem 
)
from PyQt5.QtCore import Qt, QPointF, QFileInfo
from PyQt5.QtGui import QBrush, QPen, QColor, QFont 

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
from ui.tools.financeiro import FinanceiroTool 
from ui.tools.pedidos import PedidosTool
from ui.tools.manutencao import ManutencaoTool
from ui.tools.engenharia_data import EngenhariaDataTool 
from ui.tools.excel_viewer_tool import ExcelViewerTool 
from ui.tools.structure_view_tool import StructureViewTool
from ui.tools.rpi_tool import RpiTool 
# Adicionando de volta as ferramentas que o backup mencionou/utilizava
from ui.tools.engenharia_workflow_tool import EngenhariaWorkflowTool 
from ui.tools.user_settings_tool import UserSettingsTool 


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
ESTOQUE_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "estoque_data.xlsx") # Para EstoqueTool
ITEMS_DATA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "estoque.xlsx") # Para ItemsTool (agora estoque.xlsx)
DB_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "db.xlsx") 
ENGENHARIA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "engenharia.xlsx")
BOM_DATA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "bom_data.xlsx") # Padrão para BomManagerTool (se não for engenharia.xlsx)
MANUFACTURING_DATA_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "manufacturing_data.xlsx")

# Caminhos para arquivos Excel gerenciados pelo aplicativo (na pasta app_sheets)
TOOLS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "tools.xlsx")
MODULES_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "modules.xlsx")
PERMISSIONS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "permissions.xlsx")
ROLES_TOOLS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "roles_tools.xlsx")
USERS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "users.xlsx") # Conteúdo da planilha 'users' no db.xlsx
MAIN_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "main.xlsx") # Assumindo que este arquivo existe ou será criado

# --- Caminho para o script de atualização de metadados ---
UPDATE_METADATA_SCRIPT_PATH = os.path.join(APP_SHEETS_DIR, "tools", "update_user_sheets_metadata.py")
# --- Caminho para o script de validação de sheets ---
# CORRIGIDO: O nome do diretório é "sheet_validator" (com underscore)
SHEET_VALIDATOR_SCRIPT_PATH = os.path.join(project_root, "sheet_validator", "sheet_validator.py") 
# Caminho para o script de criação de engenharia.xlsx
CREATE_ENGENHARIA_SCRIPT_PATH = os.path.join(APP_SHEETS_DIR, "tools", "create_engenharia_xlsx.py")


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
    os.path.basename(ITEMS_DATA_EXCEL_PATH), 
    os.path.basename(DB_EXCEL_PATH), 
    os.path.basename(ENGENHARIA_EXCEL_PATH),
    os.path.basename(TOOLS_EXCEL_PATH),
    os.path.basename(MODULES_EXCEL_PATH),
    os.path.basename(PERMISSIONS_EXCEL_PATH),
    os.path.basename(ROLES_TOOLS_EXCEL_PATH),
    os.path.basename(USERS_EXCEL_PATH), 
    os.path.basename(MAIN_EXCEL_PATH),
    os.path.basename(UPDATE_METADATA_SCRIPT_PATH),
    os.path.basename(SHEET_VALIDATOR_SCRIPT_PATH),
    os.path.basename(CREATE_ENGENHARIA_SCRIPT_PATH)
]

# Garante que os diretórios existam
os.makedirs(USER_SHEETS_DIR, exist_ok=True)
os.makedirs(APP_SHEETS_DIR, exist_ok=True)
# Garante que a pasta 'tools' dentro de 'app_sheets' exista
os.makedirs(os.path.dirname(UPDATE_METADATA_SCRIPT_PATH), exist_ok=True)
os.makedirs(os.path.dirname(CREATE_ENGENHARIA_SCRIPT_PATH), exist_ok=True) 
# Garante que a pasta 'sheet_validator' exista
os.makedirs(os.path.dirname(SHEET_VALIDATOR_SCRIPT_PATH), exist_ok=True)


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
        # Assumindo a ordem das colunas para compatibilidade
        # Cabeçalhos: id, username, password_hash, role, full_name, email, phone, department
        headers = [cell.value for cell in users_sheet[1]] if users_sheet.max_row > 0 else []
        header_map = {h: idx for idx, h in enumerate(headers)}

        # Valida que as colunas essenciais existem
        required_cols = ["id", "username", "password_hash", "role"]
        if not all(col in header_map for col in required_cols):
            QMessageBox.critical(None, "Erro de Cabeçalhos", 
                                 f"Planilha 'users' em '{os.path.basename(DB_EXCEL_PATH)}' não contém todos os cabeçalhos obrigatórios: {', '.join(required_cols)}")
            return {}

        for row_idx in range(2, users_sheet.max_row + 1):
            row_values = [cell.value for cell in users_sheet[row_idx]]
            
            # Garante que a linha tem dados suficientes para as colunas essenciais
            if len(row_values) > max(header_map[col] for col in required_cols):
                username = row_values[header_map["username"]]
                users[username] = {
                    "id": row_values[header_map["id"]],
                    "username": username,
                    "password_hash": row_values[header_map["password_hash"]],
                    "role": row_values[header_map["role"]]
                    # Inclui outros campos se existirem e forem mapeados
                }
                # Adiciona campos opcionais se existirem
                for col_name in ["full_name", "email", "phone", "department"]:
                    if col_name in header_map and header_map[col_name] < len(row_values):
                        users[username][col_name] = row_values[header_map[col_name]]
                    else:
                        users[username][col_name] = "" # Atribui string vazia se a coluna não existe na linha
        return users
    except FileNotFoundError:
        QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo do banco de dados não foi encontrado: {DB_EXCEL_PATH}")
        return {}
    except KeyError as ke:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha esperada não foi encontrada ou erro de chave: {ke} em {DB_EXCEL_PATH}")
        return {}
    except Exception as e:
        QMessageBox.critical(None, "Erro de Carregamento", f"Erro ao carregar usuários: {e}")
        return {}


def register_user(username, password, role="user"):
    """Registra um novo usuário no arquivo Excel do banco de dados."""
    try:
        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        sheet = wb["users"]
        
        # Carrega cabeçalhos existentes
        headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
        header_map = {h: idx for idx, h in enumerate(headers)}

        if "username" not in header_map or "password_hash" not in header_map or "role" not in header_map or "id" not in header_map:
            raise ValueError("Cabeçalhos essenciais (id, username, password_hash, role) não encontrados na planilha 'users'.")

        # Gera o próximo ID
        next_id = 1
        existing_ids = set()
        for row_idx in range(2, sheet.max_row + 1):
            row_values = [cell.value for cell in sheet[row_idx]]
            if len(row_values) > header_map["id"] and row_values[header_map["id"]] is not None:
                existing_ids.add(row_values[header_map["id"]])
            if len(row_values) > header_map["username"] and row_values[header_map["username"]] == username:
                raise ValueError("Nome de usuário já existe.")
        while next_id in existing_ids:
            next_id += 1

        password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        
        # Prepara a nova linha, garantindo que todos os cabeçalhos sejam preenchidos
        new_row_data = [""] * len(headers)
        new_row_data[header_map["id"]] = next_id
        new_row_data[header_map["username"]] = username
        new_row_data[header_map["password_hash"]] = password_hash
        new_row_data[header_map["role"]] = role
        # Outros campos (full_name, email, phone, department) ficam vazios por padrão no registro
        
        sheet.append(new_row_data)
        wb.save(DB_EXCEL_PATH)
    except FileNotFoundError:
        QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo do banco de dados não foi encontrado em: {DB_EXCEL_PATH}. Não é possível registrar o usuário.")
    except KeyError:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha 'users' não foi encontrada em {DB_EXCEL_PATH}. Não é possível registrar o usuário.")
    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Ocorreu um erro durante o registro: {e}")

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
        
        headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
        header_map = {h: idx for idx, h in enumerate(headers)}

        required_cols = ["id", "name", "description", "path"]
        if not all(col in header_map for col in required_cols):
            QMessageBox.critical(None, "Erro de Cabeçalhos", 
                                 f"Planilha 'tools' em '{os.path.basename(TOOLS_EXCEL_PATH)}' não contém todos os cabeçalhos obrigatórios: {', '.join(required_cols)}")
            return {}

        for row_idx in range(2, sheet.max_row + 1):
            row_values = [cell.value for cell in sheet[row_idx]]
            
            if len(row_values) > max(header_map[col] for col in required_cols):
                tool_id = row_values[header_map["id"]]
                tools[tool_id] = {
                    "id": tool_id,
                    "name": row_values[header_map["name"]],
                    "description": row_values[header_map["description"]],
                    "path": row_values[header_map["path"]]
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
        headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
        header_map = {h: idx for idx, h in enumerate(headers)}

        if "role" not in header_map or "allowed_tools" not in header_map:
             QMessageBox.critical(None, "Erro de Cabeçalhos", 
                                 f"Planilha 'access' em '{os.path.basename(DB_EXCEL_PATH)}' não contém os cabeçalhos obrigatórios: 'role', 'allowed_tools'.")
             return {}

        for row_idx in range(2, sheet.max_row + 1):
            row_values = [cell.value for cell in sheet[row_idx]]
            
            if len(row_values) > max(header_map["role"], header_map["allowed_tools"]):
                role = row_values[header_map["role"]]
                allowed_tools_str = str(row_values[header_map["allowed_tools"]]).strip() if row_values[header_map["allowed_tools"]] is not None else ""
                
                if allowed_tools_str.lower() == "all":
                    perms[role] = "all"
                elif allowed_tools_str:
                    perms[role] = allowed_tools_str.split(",")
                else:
                    perms[role] = [] # Nenhuma ferramenta permitida se vazio

            else:
                print(f"Aviso: Ignorando linha malformada na planilha 'access': {', '.join(str(c.value) for c in row_values)}")
        return perms
    except FileNotFoundError:
        QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo do banco de dados não foi encontrado em: {DB_EXCEL_PATH}")
        return {}
    except KeyError:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha 'access' não foi encontrada em {DB_EXCEL_PATH}. Não é possível carregar o usuário.")
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

# === CLASSE PRINCIPAL DA GUI ===
class TeamcenterStyleGUI(QMainWindow):
    def __init__(self, user_data):
        super().__init__()
        self.user_data = user_data
        self.current_user_role = user_data["role"]
        self.access_permissions = load_role_permissions()
        self.available_tools_metadata = load_tools_from_excel()
        self.setWindowTitle("5revolution ERP")
        self.setGeometry(100, 100, 1200, 800) # Tamanho padrão da janela
        self._init_ui()

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
            
            # Ação para criar engenharia.xlsx
            create_engenharia_action = QAction("Criar/Reinicializar Engenharia.xlsx", self)
            create_engenharia_action.setToolTip("Cria ou reinicializa o arquivo 'engenharia.xlsx' na pasta 'user_sheets' com as estruturas 'Estrutura' e 'Workflows' e dados de exemplo definidos.")
            create_engenharia_action.triggered.connect(self._run_create_engenharia_xlsx_script)
            admin_menu.addAction(create_engenharia_action)

            # Ação para sincronizar o esquema do banco de dados (db_db)
            sync_db_schema_action = QAction("Sincronizar Esquema do Banco de Dados (db_db)", self)
            sync_db_schema_action.setToolTip("Coleta os cabeçalhos de todas as planilhas do projeto (user_sheets e app_sheets, exceto o próprio db.xlsx) e os registra na planilha 'db_db' em 'db.xlsx'. Esta ação reconstrói o dicionário de dados do sistema, que é a base para a validação de consistência.")
            sync_db_schema_action.triggered.connect(self._run_update_db_db_schema) 
            admin_menu.addAction(sync_db_schema_action)

            # Ação para validar consistência do DB
            validate_db_consistency_action = QAction("Validar Consistência do DB", self)
            validate_db_consistency_action.setToolTip("Compara a estrutura real dos cabeçalhos das planilhas do projeto com o esquema registrado na planilha 'db_db' em 'db.xlsx', identificando inconsistências e erros.")
            validate_db_consistency_action.triggered.connect(self._run_validate_db_consistency)
            admin_menu.addAction(validate_db_consistency_action)

            admin_menu_btn.setMenu(admin_menu)
            toolbar.addWidget(admin_menu_btn)


    def _populate_tools_menu(self, menu):
        """Popula o menu de ferramentas com base nas permissões do usuário."""
        user_allowed_modules = self.access_permissions.get(self.current_user_role, [])
        if user_allowed_modules == "all":
            # Se for "all", o usuário tem acesso a todas as ferramentas
            allowed_tool_ids = self.available_tools_metadata.keys()
        else:
            # Caso contrário, filtra pelas IDs dos módulos permitidos
            # Mapeia IDs de módulos para IDs de ferramentas se necessário (supondo 1:1 por enquanto)
            allowed_tool_ids = user_allowed_modules 
        
        for tool_id, tool_info in self.available_tools_metadata.items():
            if tool_id in allowed_tool_ids:
                action = QAction(tool_info["name"], self)
                # Conecta a ação para abrir a ferramenta correspondente
                action.triggered.connect(lambda checked, t_id=tool_id: self._open_tool(t_id))
                menu.addAction(action)

    def _open_tool(self, tool_id):
        """Abre a ferramenta selecionada em uma nova aba."""
        tool_info = self.available_tools_metadata.get(tool_id)
        if not tool_info:
            QMessageBox.warning(self, "Ferramenta Não Encontrada", f"A ferramenta com ID '{tool_id}' não foi encontrada.")
            return

        tool_name = tool_info["name"]
        
        # Mapeamento de IDs de ferramentas para classes de ferramentas
        tool_classes = {
            "mod1": BomManagerTool,
            "mod3": ColaboradoresTool,
            "modX": ConfiguradorTool,
            "mod4": EngenhariaDataTool, # Esta é a ferramenta de dados (tabular)
            "mod_workflow": EngenhariaWorkflowTool, # Esta é a ferramenta de diagrama de workflow
            "mod7": EstoqueTool, # Mapeamento para EstoqueTool
            "mod8": FinanceiroTool, # Mapeamento para FinanceiroTool
            "mod11": ItemsTool, # Mapeamento para ItemsTool
            "mod5": ManufacturingTool, # Mapeamento para ManufacturingTool
            "mod10": ManutencaoTool, # Mapeamento para ManutencaoTool
            "mod9": PedidosTool, # Mapeamento para PedidosTool
            "mod6": PcpTool, # Mapeamento para PcpTool
            "mod_product_data": ProductDataTool, # Mapeamento para ProductDataTool
            "mod_rpi": RpiTool, # Mapeamento para RpiTool
            "mod_structure_view": StructureViewTool, # Mapeamento para StructureViewTool
            "mod_excel_viewer": ExcelViewerTool, # Mapeamento para ExcelViewerTool
            "mod_user_settings": UserSettingsTool, # Mapeamento para UserSettingsTool
        }

        # Dicionário de caminhos para as classes de ferramentas que precisam de um caminho de arquivo
        tool_file_paths = {
            "mod4": ENGENHARIA_EXCEL_PATH, # EngenhariaDataTool usa engenharia.xlsx
            "mod_workflow": ENGENHARIA_EXCEL_PATH, # EngenhariaWorkflowTool TAMBÉM usa engenharia.xlsx (para abas diferentes)
            "mod3": COLABORADORES_EXCEL_PATH,
            "mod7": ESTOQUE_EXCEL_PATH, # Caminho do arquivo para EstoqueTool
            "mod8": FINANCEIRO_EXCEL_PATH,
            "mod11": ITEMS_DATA_EXCEL_PATH, # Caminho do arquivo para ItemsTool
            "mod10": MANUTENCAO_EXCEL_PATH,
            "mod9": PEDIDOS_EXCEL_PATH,
            "mod6": PROGRAMACAO_EXCEL_PATH, # Assumindo PCPTool usa programacao.xlsx
            "mod1": BOM_DATA_EXCEL_PATH, # BomManagerTool usa bom_data.xlsx por padrão, mas pode usar engenharia.xlsx
            "modX": CONFIGURADOR_EXCEL_PATH, # ConfiguradorTool usa configurador.xlsx
            "mod5": MANUFACTURING_DATA_EXCEL_PATH, # Caminho para ManufacturingTool
            "mod_product_data": OUTPUT_EXCEL_PATH, # Caminho para ProductDataTool
            "mod_rpi": RPI_EXCEL_PATH, # Caminho para RpiTool
            "mod_structure_view": ENGENHARIA_EXCEL_PATH, # Caminho padrão para StructureViewTool, pode ser ajustado
            "mod_excel_viewer": None, # ExcelViewerTool não precisa de um path inicial fixo, será passado dinamicamente
            "mod_user_settings": None, # UserSettingsTool não precisa de um path de arquivo Excel no construtor
        }

        ToolClass = tool_classes.get(tool_id)
        if ToolClass:
            try:
                tool_instance = None
                if tool_id in tool_file_paths and tool_file_paths[tool_id] is not None:
                    # Lógica para ferramentas que precisam de um caminho de arquivo
                    # Para EngenhariaWorkflowTool, passamos o caminho e a sheet padrão para workflows
                    if tool_id == "mod_workflow":
                        tool_instance = ToolClass(tool_file_paths[tool_id], sheet_name="Workflows")
                    # Para ItemsTool ou RpiTool, verificamos se é para ser somente leitura (se o arquivo for engenharia.xlsx)
                    elif tool_id in ["mod11", "mod_rpi"] and os.path.basename(tool_file_paths[tool_id]).lower() == "engenharia.xlsx":
                        tool_instance = ToolClass(tool_file_paths[tool_id], read_only=True)
                    # Para StructureViewTool, podemos passar um sheet_name específico se necessário, ou deixar para o default
                    elif tool_id == "mod_structure_view":
                        tool_instance = ToolClass(tool_file_paths[tool_id], sheet_name="Estrutura") # Exemplo: assume sheet 'Estrutura' para engenharia.xlsx
                    else:
                        tool_instance = ToolClass(tool_file_paths[tool_id])
                elif tool_id == "mod_user_settings":
                    # Passa o dicionário user_data completo para UserSettingsTool
                    tool_instance = UserSettingsTool(self.user_data) 
                else:
                    # Para ferramentas que não precisam de um file_path no construtor
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
        self.tree_widget.setHeaderLabels(["Nome", "Tipo"]) # Define cabeçalhos
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
        main_splitter.setSizes([300, 900]) # 300px para a árvore, 900px para as abas

    def _setup_initial_content(self):
        """Popula a árvore de navegação e pode abrir abas iniciais."""
        self._populate_workspace_tree()
        self._populate_file_system_tree()
        # Não abre nenhuma aba por padrão, o usuário fará isso.

    def _populate_workspace_tree(self):
        """Popula a seção 'Espaço de Trabalho' da árvore."""
        self.tree_widget.clear() # Limpa a árvore existente
        
        # Adiciona o item raiz "Projetos/Espaço de Trabalho"
        workspace_root_item = QTreeWidgetItem(self.tree_widget, ["Projetos/Espaço de Trabalho", "Pasta"])
        workspace_root_item.setExpanded(True) # Expande o item raiz
        
        # Adiciona itens de exemplo codificados
        for item_name in WORKSPACE_ITEMS:
            QTreeWidgetItem(workspace_root_item, [item_name, "Item"])

    def _populate_file_system_tree(self):
        """Popula as seções 'Arquivos do Usuário' e 'Arquivos do Sistema' da árvore."""
        # Seções de arquivos
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
                if filename.endswith(".xlsx") and not filename.startswith('~$'): # Ignora arquivos temporários
                    file_path = os.path.join(directory, filename)
                    file_info = QFileInfo(file_path)
                    item = QTreeWidgetItem(parent_item, [file_info.fileName(), "Arquivo Excel"])
                    item.setData(0, Qt.UserRole, file_path) # Armazena o caminho completo
        except Exception as e:
            QMessageBox.warning(self, "Erro ao Listar Arquivos", f"Não foi possível listar arquivos em {directory}: {e}")

    def _on_tree_item_double_clicked(self, item, column):
        """Lida com o clique duplo em um item da árvore."""
        file_path = item.data(0, Qt.UserRole) # Obtém o caminho do arquivo armazenado
        if file_path and os.path.exists(file_path):
            self._open_excel_file_in_viewer(file_path)
        elif not file_path:
            # Caso seja um nó de pasta ou um item não-arquivo
            pass
        else:
            QMessageBox.warning(self, "Arquivo Não Encontrado", f"O arquivo '{os.path.basename(file_path)}' não existe ou o caminho está incorreto.")

    def _open_excel_file_in_viewer(self, file_path):
        """Abre um arquivo Excel usando o ExcelViewerTool."""
        tool_name = f"Viewer: {os.path.basename(file_path)}"
        
        # Verifica se a aba já está aberta para evitar duplicatas
        for i in range(self.central_widget.count()):
            if self.central_widget.tabText(i) == tool_name:
                self.central_widget.setCurrentIndex(i)
                return

        try:
            # Instancia e abre a ferramenta ExcelViewerTool (agora um visualizador puro)
            excel_viewer_tool = ExcelViewerTool(file_path=file_path)
            self.central_widget.addTab(excel_viewer_tool, tool_name)
            self.central_widget.setCurrentWidget(excel_viewer_tool)
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Abrir Arquivo", f"Não foi possível abrir '{os.path.basename(file_path)}' no visualizador: {e}")
            
    # --- NOVAS FUNÇÕES PARA EXECUTAR SCRIPTS EXTERNOS ---
    def _run_external_python_script(self, script_path, *args):
        """
        Executa um script Python externo em um processo separado.
        Exibe uma caixa de mensagem com o resultado.
        """
        # Chamada mais robusta para PyInstaller e ambientes virtuais
        cmd = [sys.executable, script_path] + list(args)
        
        try:
            # shell=False é mais seguro e preferível
            # text=True (ou universal_newlines=True para py < 3.7) para capturar stdout/stderr como texto
            result = subprocess.run(cmd, capture_output=True, text=True, check=True)
            output = result.stdout.strip()
            error = result.stderr.strip()

            if result.returncode == 0:
                QMessageBox.information(self, "Sucesso na Execução do Script", f"Script executado com sucesso:\n\n{output}")
                print(f"Sucesso: {output}") # Imprime também no console para depuração
            else:
                QMessageBox.critical(self, "Erro na Execução do Script", f"O script retornou um erro (Código: {result.returncode}):\n\n{error}\n{output}")
                print(f"Erro: {error}\nOutput: {output}") # Imprime no console
        except FileNotFoundError:
            QMessageBox.critical(self, "Erro de Arquivo", f"O executável Python ou o script '{os.path.basename(script_path)}' não foi encontrado. Verifique o PATH ou o caminho do script.")
            print(f"Erro: Python executable or script '{script_path}' not found.")
        except subprocess.CalledProcessError as e:
            QMessageBox.critical(self, "Erro no Processo", f"O script '{os.path.basename(script_path)}' falhou:\n\n{e.stdout}\n{e.stderr}")
            print(f"Erro no processo: {e.stdout}\n{e.stderr}")
        except Exception as e:
            QMessageBox.critical(self, "Erro Inesperado", f"Ocorreu um erro inesperado ao tentar executar o script '{os.path.basename(script_path)}': {e}")
            print(f"Erro inesperado ao executar script: {e}")

    def _run_create_engenharia_xlsx_script(self):
        """Executa o script para criar/atualizar engenharia.xlsx."""
        if not os.path.exists(CREATE_ENGENHARIA_SCRIPT_PATH):
            QMessageBox.critical(self, "Erro", f"O script para criar engenharia.xlsx não foi encontrado em: {CREATE_ENGENHARIA_SCRIPT_PATH}")
            return
        self._run_external_python_script(CREATE_ENGENHARIA_SCRIPT_PATH)
        # Após a criação/atualização, recarrega a árvore de arquivos para refletir a possível criação
        self._populate_file_system_tree()

    def _run_update_user_sheets_headers(self):
        """
        OBSOLETO: Esta função foi consolidada em _run_update_db_db_schema.
        Mantida temporariamente para evitar quebras se ainda houver referências.
        """
        QMessageBox.information(self, "Ação Consolidada", "Esta função foi consolidada com 'Sincronizar Esquema do Banco de Dados (db_db)'. Por favor, use essa opção.")

    def _run_update_db_db_schema(self):
        """Executa o script para atualizar o schema db_db com os cabeçalhos reais."""
        self._run_external_python_script(UPDATE_METADATA_SCRIPT_PATH, "update_db_schema")
        # A db_db é um arquivo de sistema/metadados, não um arquivo que aparece na árvore user_sheets/app_sheets
        # Recarregar o tree_widget pode não ser necessário, mas não faz mal.
        self._populate_file_system_tree()


    def _run_validate_db_consistency(self):
        """Executa o script para validar a consistência do DB."""
        self._run_external_python_script(SHEET_VALIDATOR_SCRIPT_PATH, "validate")
        # A validação apenas gera um relatório, não altera arquivos visíveis diretamente,
        # então não há necessidade de recarregar a árvore.

# === INÍCIO DO APLICATIVO ===
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Cria a janela de login e a mostra
    login_window = LoginWindow()
    login_window.show()
    
    sys.exit(app.exec_())
