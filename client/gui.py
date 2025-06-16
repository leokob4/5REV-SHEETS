import sys
import os
import bcrypt
import openpyxl
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QAction, QTabWidget, QMenu, QToolButton,
    QWidget, QVBoxLayout, QSplitter, QTreeWidget, QTreeWidgetItem,
    QLabel, QLineEdit, QPushButton, QHBoxLayout, QMessageBox, QGraphicsView,
    QGraphicsScene, QGraphicsRectItem, QGraphicsLineItem, QDialog, QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView, QComboBox
)
from PyQt5.QtCore import Qt, QPointF
from PyQt5.QtGui import QBrush, QPen, QColor, QFont

# --- Configuração de Paths para Importação de Módulos ---
# Obtém o diretório atual do gui.py
current_dir = os.path.dirname(os.path.abspath(__file__))
# Navega para o diretório raiz do projeto (assumindo que gui.py está em client/)
project_root = os.path.dirname(current_dir)

# Adiciona o diretório raiz do projeto ao sys.path para que módulos como 'ui' e 'client' possam ser encontrados
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# --- Importação dos Módulos das Ferramentas ---
# Certifique-se de que esses arquivos existem nas pastas corretas e que os __init__.py estão presentes.
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
from ui.tools.structure_view_tool import StructureViewTool
from ui.tools.user_settings_tool import UserSettingsTool
from ui.tools.engenharia_workflow_tool import EngenhariaWorkflowTool

# Importa o novo diálogo de adição de item (agora de client.add_item_dialog)
from client.add_item_dialog import AddItemDialog

# --- Configuração de Caminhos de Arquivos ---
# Define paths padrão para consistência.
USER_SHEETS_DIR = os.path.join(project_root, "user_sheets")
APP_SHEETS_DIR = os.path.join(project_root, "app_sheets")
DB_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "db.xlsx")
TOOLS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "tools.xlsx")
WORKSPACE_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "workspace_data.xlsx")

# Garante que os diretórios existam
os.makedirs(USER_SHEETS_DIR, exist_ok=True)
os.makedirs(APP_SHEETS_DIR, exist_ok=True)

# === FUNÇÕES DE AJUDA PARA PLANILHAS ===
def load_users_from_excel():
    """Carrega dados de usuários do arquivo Excel do banco de dados."""
    try:
        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        users_sheet = wb["users"]
        users = {}
        for row in users_sheet.iter_rows(min_row=2):
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
        next_id = sheet.max_row
        for row in sheet.iter_rows(min_row=2):
            if row[1].value == username:
                raise ValueError("Nome de usuário já existe.")

        password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        sheet.append([next_id, username, password_hash, role])
        wb.save(DB_EXCEL_PATH)
    except FileNotFoundError:
        QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo do banco de dados não foi encontrado em: {DB_EXCEL_PATH}. Não é possível registrar o usuário.")
    except KeyError:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha 'users' não foi encontrada em {DB_EXCEL_PATH}. Não é possível registrar o usuário.")
    except Exception as e:
        QMessageBox.critical(None, "Erro", f"Ocorreu um erro durante o registro do usuário: {e}")

def load_tools_from_excel():
    """
    Carrega dados de ferramentas do arquivo Excel dedicado.
    Caminho corrigido para 'app_sheets/tools.xlsx' e tratamento de erros adicionado.
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
    """Carrega permissões de função do arquivo Excel do banco de dados."""
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

def load_workspace_items_from_excel():
    """Carrega itens do espaço de trabalho de workspace_data.xlsx, criando-o se necessário."""
    items = []
    try:
        if not os.path.exists(WORKSPACE_EXCEL_PATH):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "items"
            ws.append(["ID", "Name", "Type", "ParentID", "Description"])
            
            ws_structure = wb.create_sheet("structure")
            ws_structure.append(["ParentID", "ComponentID", "ComponentName", "Quantity", "Unit", "Type"])

            wb.save(WORKSPACE_EXCEL_PATH)
            QMessageBox.information(None, "Arquivo de Workspace Criado", f"O arquivo '{WORKSPACE_EXCEL_PATH}' foi criado.")
        
        wb = openpyxl.load_workbook(WORKSPACE_EXCEL_PATH)
        if "items" not in wb.sheetnames:
            QMessageBox.warning(None, "Planilha 'items' Não Encontrada", f"A planilha 'items' não foi encontrada em '{WORKSPACE_EXCEL_PATH}'. Criando uma nova.")
            ws = wb.create_sheet("items")
            ws.append(["ID", "Name", "Type", "ParentID", "Description"])
            wb.save(WORKSPACE_EXCEL_PATH)

        sheet = wb["items"]
        headers = [cell.value for cell in sheet[1]]
        
        header_map = {header: idx for idx, header in enumerate(headers)}
        
        for row_idx in range(2, sheet.max_row + 1):
            row_values = [cell.value for cell in sheet[row_idx]]
            item_data = {}
            for col_name, col_idx in header_map.items():
                item_data[col_name] = row_values[col_idx] if col_idx < len(row_values) else None
            items.append(item_data)
            
    except Exception as e:
        QMessageBox.critical(None, "Erro ao Carregar Itens do Workspace", f"Erro: {e}")
    return items

def add_workspace_item_to_excel(item_data):
    """Adiciona um novo item à planilha 'items' em workspace_data.xlsx."""
    try:
        wb = openpyxl.load_workbook(WORKSPACE_EXCEL_PATH)
        sheet = wb["items"]

        for row in sheet.iter_rows(min_row=2):
            if row[0].value == item_data["ID"]:
                QMessageBox.warning(None, "ID Duplicado", f"Um item com o ID '{item_data['ID']}' já existe. Por favor, use um ID único.")
                return False

        sheet.append([
            item_data.get("ID"),
            item_data.get("Name"),
            item_data.get("Type"),
            item_data.get("ParentID"),
            item_data.get("Description")
        ])
        wb.save(WORKSPACE_EXCEL_PATH)
        QMessageBox.information(None, "Item Adicionado", f"Item '{item_data['Name']}' adicionado com sucesso ao Espaço de Trabalho.")
        return True
    except FileNotFoundError:
        QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo '{WORKSPACE_EXCEL_PATH}' não foi encontrado.")
        return False
    except KeyError:
        QMessageBox.critical(None, "Erro de Planilha", f"A planilha 'items' não foi encontrada em '{WORKSPACE_EXCEL_PATH}'.")
        return False
    except Exception as e:
        QMessageBox.critical(None, "Erro ao Adicionar Item", f"Erro ao adicionar item ao workspace: {e}")
        return False


# === JANELA DE LOGIN ===
class LoginWindow(QWidget):
    """
    A janela de login para a aplicação.
    Gerencia a autenticação e o registro de usuários.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("5revolution Login")
        self.setGeometry(400, 200, 300, 180)
        self.users = load_users_from_excel()

        self._init_ui()

    def _init_ui(self):
        """Inicializa os elementos da interface de usuário para a janela de login."""
        layout = QVBoxLayout()

        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Nome de Usuário")
        self.username_input.returnPressed.connect(self.authenticate)
        
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
        self.main.show()
        self.close()

    def handle_register(self):
        """Gerencia o registro de usuário."""
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

# === GUI PRINCIPAL ===
class TeamcenterStyleGUI(QMainWindow):
    """
    A GUI principal da aplicação, estilizada para se assemelhar ao Teamcenter.
    Fornece uma visualização em árvore do espaço de trabalho, área de conteúdo com abas e uma barra de ferramentas.
    """
    def __init__(self, user):
        super().__init__()
        self.setWindowTitle("Plataforma 5revolution")
        self.setGeometry(100, 100, 1280, 800)

        self.username = user["username"]
        self.role = user["role"]
        self.tools = load_tools_from_excel()
        self.permissions = load_role_permissions()
        self.workspace_items_data = [] # Será recarregado ao popular a árvore

        self._create_toolbar()
        self._create_main_layout()

        self.statusBar().showMessage(f"Logado como: {self.username} | Papel: {self.role}")

    def _create_toolbar(self):
        """Cria a barra de ferramentas principal da aplicação."""
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
                
                # Conecta dinamicamente ações aos widgets das ferramentas corretas
                if tool["id"] == "mod4": # Engenharia (Workflow)
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, EngenhariaWorkflowTool()))
                elif tool["id"] == "mes_pcp": # MES (Apontamento Fábrica)
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, self._create_mes_pcp_tool_widget()))
                elif tool["id"] == "prod_data":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ProductDataTool()))
                elif tool["id"] == "bom_manager":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, BomManagerTool()))
                elif tool["id"] == "configurador":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ConfiguradorTool()))
                elif tool["id"] == "colab":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ColaboradoresTool()))
                elif tool["id"] == "items_tool":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ItemsTool()))
                elif tool["id"] == "manuf":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ManufacturingTool()))
                elif tool["id"] == "pcp_tool":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, PcpTool()))
                elif tool["id"] == "estoque_tool":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, EstoqueTool()))
                elif tool["id"] == "financeiro":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, FinanceiroTool()))
                elif tool["id"] == "pedidos":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, PedidosTool()))
                elif tool["id"] == "manutencao":
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ManutencaoTool()))
                else:
                    action.triggered.connect(lambda chk=False, title=tool["name"], desc=tool["description"]: self._open_tab(title, QLabel(desc)))
        self.tools_btn.setMenu(tools_menu)
        self.toolbar.addWidget(self.tools_btn)

        self.profile_btn = QToolButton()
        self.profile_btn.setText(f"👤 {self.username}")
        self.profile_btn.setPopupMode(QToolButton.InstantPopup)
        profile_menu = QMenu()
        profile_menu.addAction("⚙️ Configurações", lambda: self._open_tab("Configurações do Usuário", UserSettingsTool(self.username, self.role)))
        profile_menu.addSeparator()
        profile_menu.addAction("🔒 Sair", self._logout)
        self.profile_btn.setMenu(profile_menu)
        self.toolbar.addWidget(self.profile_btn)

    def _create_main_layout(self):
        """Cria o layout dividido principal com a visualização em árvore e as abas."""
        self.splitter = QSplitter()

        left_pane_widget = QWidget()
        left_pane_layout = QVBoxLayout(left_pane_widget)
        left_pane_layout.setContentsMargins(0, 0, 0, 0)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Espaço de Trabalho")
        self._populate_workspace_tree()
        self.tree.expandAll()
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._show_tree_context_menu)
        self.tree.itemDoubleClicked.connect(self._open_structure_view_for_item)

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

    def _populate_workspace_tree(self):
        """Popula a árvore com dados carregados de workspace_data.xlsx."""
        self.tree.clear()
        self.workspace_items_data = load_workspace_items_from_excel()

        item_map = {item['ID']: {'data': item, 'children': []} for item in self.workspace_items_data}
        
        root_items = []
        for item_id, item_info in item_map.items():
            parent_id = item_info['data'].get('ParentID')
            if parent_id == 'ROOT':
                root_items.append(item_info['data'])
            elif parent_id in item_map:
                item_map[parent_id]['children'].append(item_info['data'])

        def add_items_to_tree(parent_qtree_item, items_list):
            for item_data in items_list:
                icon_map = {
                    "Project": "📁", "Assembly": "📦", "Part": "📄",
                    "Component": "🧩", "Document": "📜", "Variant": "💡"
                }
                icon = icon_map.get(item_data.get('Type', ''), '❓')
                
                q_item = QTreeWidgetItem([f"{icon} {item_data.get('Name', 'N/A')}"])
                q_item.setData(0, Qt.UserRole, item_data.get('ID'))
                q_item.setData(1, Qt.UserRole, item_data.get('Name'))
                q_item.setData(2, Qt.UserRole, item_data.get('Type'))
                q_item.setData(3, Qt.UserRole, item_data.get('ParentID'))
                q_item.setData(4, Qt.UserRole, item_data.get('Description'))
                
                parent_qtree_item.addChild(q_item)
                
                if item_data.get('ID') in item_map and item_map[item_data['ID']]['children']:
                    add_items_to_tree(q_item, item_map[item_data['ID']]['children'])

        for item_data in root_items:
            icon_map = {
                "Project": "📁", "Assembly": "📦", "Part": "📄",
                "Component": "🧩", "Document": "📜", "Variant": "💡"
            }
            icon = icon_map.get(item_data.get('Type', ''), '❓')
            
            root_q_item = QTreeWidgetItem([f"{icon} {item_data.get('Name', 'N/A')}"])
            root_q_item.setData(0, Qt.UserRole, item_data.get('ID'))
            root_q_item.setData(1, Qt.UserRole, item_data.get('Name'))
            root_q_item.setData(2, Qt.UserRole, item_data.get('Type'))
            root_q_item.setData(3, Qt.UserRole, item_data.get('ParentID'))
            root_q_item.setData(4, Qt.UserRole, item_data.get('Description'))
            self.tree.addTopLevelItem(root_q_item)

            if item_data.get('ID') in item_map and item_map[item_data['ID']]['children']:
                add_items_to_tree(root_q_item, item_map[item_data['ID']]['children'])
        
        self.tree.expandAll()

    def _open_structure_view_for_item(self, item, column):
        """
        Abre uma nova aba com a StructureViewTool para o item clicado duas vezes.
        """
        item_id = item.data(0, Qt.UserRole)
        if not item_id:
            QMessageBox.warning(self, "Erro", "Não foi possível obter o ID do item para visualização da estrutura.")
            return

        item_name = item.text(0).split(' ', 1)[1].strip()

        tab_title = f"Estrutura: {item_name}"
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == tab_title:
                self.tabs.setCurrentIndex(i)
                QMessageBox.information(self, "Informação", f"A guia para a estrutura de '{item_name}' já está aberta.")
                return

        structure_tool = StructureViewTool(item_id, item_name)
        self._open_tab(tab_title, structure_tool)

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
        """Gerencia o envio de dados MES (placeholder)."""
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
        Realiza uma busca nos itens do espaço de trabalho e exibe os resultados em um diálogo.
        """
        search_term = self.item_search_bar.text().strip().lower()
        if not search_term:
            QMessageBox.information(self, "Pesquisar", "Por favor, digite um termo de pesquisa.")
            return

        results = [item for item in self.workspace_items_data if search_term in item.get('Name', '').lower() or search_term in item.get('ID', '').lower()]
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
                list_item_text = f"{item.get('Name', 'N/A')} ({item.get('ID', 'N/A')})"
                list_item = QListWidgetItem(list_item_text)
                list_item.setData(Qt.UserRole, item.get('ID'))
                list_item.setData(Qt.UserRole + 1, item.get('Name'))
                list_widget.addItem(list_item)

            list_widget.itemDoubleClicked.connect(
                lambda item_list_widget: self._open_structure_view_for_item_by_data(item_list_widget.data(Qt.UserRole), item_list_widget.data(Qt.UserRole + 1)) or dialog.accept()
            ) 
            layout.addWidget(list_widget)

        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)

        dialog.exec_()

    def _open_structure_view_for_item_by_data(self, item_id, item_name):
        """Função auxiliar para abrir a visualização da estrutura diretamente do ID e Nome."""
        tab_title = f"Estrutura: {item_name}"
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == tab_title:
                self.tabs.setCurrentIndex(i)
                QMessageBox.information(self, "Informação", f"A guia para a estrutura de '{item_name}' já está aberta.")
                return

        structure_tool = StructureViewTool(item_id, item_name)
        self._open_tab(tab_title, structure_tool)

    def find_qtree_item_by_id(self, item_id):
        """
        Função auxiliar para encontrar um QTreeWidgetItem pelo seu ID armazenado (Qt.UserRole).
        Realiza uma busca recursiva na árvore.
        """
        def search_item(parent_item, target_id):
            for i in range(parent_item.childCount()):
                child_item = parent_item.child(i)
                if child_item.data(0, Qt.UserRole) == target_id:
                    return child_item
                found_in_children = search_item(child_item, target_id)
                if found_in_children:
                    return found_in_children
            return None

        for i in range(self.tree.topLevelItemCount()):
            top_item = self.tree.topLevelItem(i)
            if top_item.data(0, Qt.UserRole) == item_id:
                return top_item
            found = search_item(top_item, item_id)
            if found:
                return found
        return None

    def _open_tab(self, title, widget_instance):
        """
        Abre uma nova aba ou alterna para uma existente.
        Aceita uma instância de widget diretamente.
        """
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == title:
                self.tabs.setCurrentIndex(i)
                return
        self.tabs.addTab(widget_instance, title)
        self.tabs.setCurrentIndex(self.tabs.count() - 1)

    def _open_options(self):
        """Abre o diálogo de opções/configurações do usuário."""
        self._open_tab("Configurações do Usuário", UserSettingsTool(self.username, self.role))

    def _logout(self):
        """Desconecta o usuário atual e retorna para a tela de login."""
        confirm_logout = QMessageBox.question(self, "Confirmação de Saída", "Tem certeza de que deseja sair?",
                                              QMessageBox.Yes | QMessageBox.No)
        if confirm_logout == QMessageBox.Yes:
            self.close()
            self.login = LoginWindow()
            self.login.show()

    def _show_tree_context_menu(self, pos):
        """Exibe um menu de contexto para itens na visualização em árvore."""
        item = self.tree.itemAt(pos)
        
        menu = QMenu()

        add_root_item_action = menu.addAction("➕ Adicionar Novo Item Raiz")
        add_root_item_action.triggered.connect(self._add_new_root_item)

        if item:
            item_id = item.data(0, Qt.UserRole)
            item_name = item.data(1, Qt.UserRole) or item.text(0).split(' ', 1)[1].strip()

            menu.addSeparator()
            
            menu.addAction("🔍 Ver Detalhes/Estrutura", lambda: self._open_structure_view_for_item(item, 0))
            menu.addAction("✏️ Editar Propriedades (Simulado)", lambda: QMessageBox.information(self, "Ação Simulada", f"Editando propriedades para: {item_name} (ação simulada)"))
            menu.addAction("❌ Excluir Item (Simulado)", lambda: QMessageBox.warning(self, "Ação Simulada", f"Excluindo: {item_name} (ação simulada)"))
            
            add_subitem_action = menu.addAction("➕ Adicionar Subitem")
            add_subitem_action.triggered.connect(lambda: self._add_new_subitem(item_id, item_name))

        menu.exec_(self.tree.viewport().mapToGlobal(pos))

    def _add_new_root_item(self):
        """Abre um diálogo para adicionar um novo item de nível superior ao espaço de trabalho."""
        dialog = AddItemDialog(parent_id="ROOT", parent_name="ROOT", parent=self)
        if dialog.exec_() == QDialog.Accepted:
            item_data = dialog.item_data
            if add_workspace_item_to_excel(item_data):
                self._populate_workspace_tree()
    
    def _add_new_subitem(self, parent_id, parent_name):
        """Abre um diálogo para adicionar um novo subitem a um pai selecionado no espaço de trabalho."""
        dialog = AddItemDialog(parent_id=parent_id, parent_name=parent_name, parent=self)
        if dialog.exec_() == QDialog.Accepted:
            item_data = dialog.item_data
            if add_workspace_item_to_excel(item_data):
                self._populate_workspace_tree()

    def _show_tab_context_menu(self, pos):
        """Exibe um menu de contexto para as abas no widget de abas."""
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
        """Fecha todas as abas, exceto a do índice 'keep_index'."""
        for i in reversed(range(self.tabs.count())):
            if i != keep_index:
                self.tabs.removeTab(i)

# === PONTO DE ENTRADA ===
if __name__ == "__main__":
    app = QApplication(sys.argv)
    login = LoginWindow()
    login.show()
    sys.exit(app.exec_())
