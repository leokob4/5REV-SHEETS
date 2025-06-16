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
project_root = os.path.dirname(current_dir) # Variável 'project_root' definida aqui (com 'p' minúsculo)
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
from ui.tools.engenharia_workflow_tool import EngenhariaWorkflowTool # IMPORTANTE: Importar a nova tool


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

# --- Caminho para o script de atualização de metadados ---
UPDATE_METADATA_SCRIPT_PATH = os.path.join(APP_SHEETS_DIR, "tools", "update_user_sheets_metadata.py")
# --- Caminho para o script de validação de sheets ---
SHEET_VALIDATOR_SCRIPT_PATH = os.path.join(project_root, "sheet validator", "sheet_validator.py")
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
    os.path.basename(DB_EXCEL_PATH), # db.xlsx é protegido
    os.path.basename(ENGENHARIA_EXCEL_PATH),
    os.path.basename(TOOLS_EXCEL_PATH),
    os.path.basename(MODULES_EXCEL_PATH),
    os.path.basename(PERMISSIONS_EXCEL_PATH),
    os.path.basename(ROLES_TOOLS_EXCEL_PATH),
    os.path.basename(USERS_EXCEL_PATH), # Redundante se db.xlsx for protegido, mas mantido para clareza
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
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
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
            QMessageBox.critical(self, "Erro de Carregamento", f"Erro ao carregar dados existentes de db_db: {e}")
        return existing_data

    def _update_db_headers(self):
        # Esta função agora é um placeholder ou pode ser removida/revisada
        # pois a lógica principal será movida para funções externas.
        QMessageBox.information(self, "Ação", "Esta função será executada via Ferramentas Admin no menu principal.")
        pass

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
            create_engenharia_action = QAction("Criar engenharia.xlsx", self)
            create_engenharia_action.setToolTip("Cria ou reinicializa o arquivo 'engenharia.xlsx' na pasta 'user_sheets' com um schema padrão e dados de exemplo.")
            create_engenharia_action.triggered.connect(self._run_create_engenharia_xlsx_script)
            admin_menu.addAction(create_engenharia_action)

            # Ação para atualizar cabeçalhos das User Sheets
            update_user_sheets_action = QAction("Atualizar Cabeçalhos das User Sheets", self)
            update_user_sheets_action.setToolTip("Este script varre a pasta 'user_sheets' para coletar os cabeçalhos atuais de todas as planilhas. Esses cabeçalhos são então salvos na planilha 'db_db' em 'db.xlsx', servindo como a 'fonte da verdade' para o esquema de dados do sistema.")
            update_user_sheets_action.triggered.connect(self._run_update_user_sheets_headers)
            admin_menu.addAction(update_user_sheets_action)

            # Ação para atualizar o schema db_db
            update_db_db_schema_action = QAction("Atualizar Schema db_db", self)
            update_db_db_schema_action.setToolTip("Atualiza os cabeçalhos e metadados na planilha 'db_db' em 'db.xlsx' com base nos esquemas reais das planilhas em 'user_sheets' e 'app_sheets'. Essencial após alterações na estrutura das planilhas.")
            update_db_db_schema_action.triggered.connect(self._run_update_db_db_schema)
            admin_menu.addAction(update_db_db_schema_action)

            # Ação para validar consistência do DB
            validate_db_consistency_action = QAction("Validar Consistência do DB", self)
            validate_db_consistency_action.setToolTip("Executa uma verificação para garantir que a estrutura e os dados das planilhas no sistema estejam consistentes e sigam os esquemas definidos, ajudando a identificar erros ou anomalias.")
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
            # Adicione outros mapeamentos aqui conforme suas ferramentas em ui/tools/
            # Ex: "mod2": TwinSyncTool,
            # "mod5": ManufacturingTool,
            # "mod6": PcpTool,
            # "mod7": ItemsTool, # Ou EstoqueTool, depende de como mapeou
            # "mod8": FinanceiroTool,
            # "mod9": PedidosTool,
            # "mod10": ManutencaoTool,
        }

        # Dicionário de caminhos para as classes de ferramentas que precisam de um caminho de arquivo
        tool_file_paths = {
            "mod4": ENGENHARIA_EXCEL_PATH, # EngenhariaDataTool usa engenharia.xlsx
            "mod_workflow": ENGENHARIA_EXCEL_PATH, # EngenhariaWorkflowTool TAMBÉM usa engenharia.xlsx (para abas diferentes)
            # Adicione outros mapeamentos aqui se alguma ferramenta específica precisar de um caminho de arquivo
            "mod3": COLABORADORES_EXCEL_PATH,
            "mod7": ESTOQUE_EXCEL_PATH,
            "mod8": FINANCEIRO_EXCEL_PATH,
            "mod10": MANUTENCAO_EXCEL_PATH,
            "mod9": PEDIDOS_EXCEL_PATH,
            "mod6": PROGRAMACAO_EXCEL_PATH, # Assumindo PCPTool usa programacao.xlsx
            "mod1": BOM_DATA_EXCEL_PATH, # BomManagerTool usa bom_data.xlsx por padrão, mas pode usar engenharia.xlsx
            "modX": CONFIGURADOR_EXCEL_PATH # ConfiguradorTool usa configurador.xlsx
        }

        ToolClass = tool_classes.get(tool_id)
        if ToolClass:
            try:
                # Verifica se a ferramenta espera um 'file_path' no construtor
                if tool_id in tool_file_paths:
                    # Para EngenhariaWorkflowTool, passamos o caminho e a sheet padrão para workflows
                    if tool_id == "mod_workflow":
                        tool_instance = ToolClass(tool_file_paths[tool_id], sheet_name="Workflows")
                    else:
                        tool_instance = ToolClass(tool_file_paths[tool_id])
                else:
                    tool_instance = ToolClass()

                self.central_widget.addTab(tool_instance, tool_name)
                self.central_widget.setCurrentWidget(tool_instance)
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
            # Instancia e abre a ferramenta ExcelViewerTool
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
        """Executa o script para atualizar os cabeçalhos das user_sheets."""
        self._run_external_python_script(UPDATE_METADATA_SCRIPT_PATH, "update_user_sheets")
        # Após a atualização, recarrega a árvore de arquivos para refletir possíveis mudanças
        self._populate_file_system_tree()

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

