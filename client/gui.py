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
from PyQt5.QtGui import QBrush, QPen, QColor, QFont # Import QFont for EngenhariaWorkflowTool (even if not explicitly used here directly)

# --- Fix for ModuleNotFoundError: No module named 'ui' ---
# Get the absolute path of the directory containing gui.py
current_dir = os.path.dirname(os.path.abspath(__file__))
# Navigate up to the project root (assuming gui.py is in client/, and client/ is in project_root/)
project_root = os.path.dirname(current_dir)
# Add the project root to sys.path so Python can find 'ui' and 'user_sheets' etc.
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# --- Import New Tool Modules ---
# Ensure these files exist in client/ui/tools/
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
from ui.tools.structure_view_tool import StructureViewTool # New tool for structure view
from ui.tools.user_settings_tool import UserSettingsTool # Make UserSettingsTool a separate importable module
from ui.tools.engenharia_workflow_tool import EngenhariaWorkflowTool # New import for refactored tool

# Import the new AddItemDialog
from client.add_item_dialog import AddItemDialog

# --- File Paths Configuration ---
# Define standard paths for consistency.
USER_SHEETS_DIR = os.path.join(project_root, "user_sheets")
APP_SHEETS_DIR = os.path.join(project_root, "app_sheets")
DB_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "db.xlsx")
TOOLS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "tools.xlsx")
WORKSPACE_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "workspace_data.xlsx") # New path for workspace data

# Ensure directories exist
os.makedirs(USER_SHEETS_DIR, exist_ok=True)
os.makedirs(APP_SHEETS_DIR, exist_ok=True)

# === SHEET HELPERS ===
def load_users_from_excel():
    """Loads user data from the database Excel file."""
    try:
        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        users_sheet = wb["users"]
        users = {}
        # Iterate from the second row to skip headers
        for row in users_sheet.iter_rows(min_row=2):
            # Check if row has enough cells before accessing
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
    """Registers a new user into the database Excel file."""
    try:
        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        sheet = wb["users"]
        next_id = sheet.max_row # Get the next available row number for ID
        # Ensure unique username
        for row in sheet.iter_rows(min_row=2):
            if row[1].value == username:
                raise ValueError("Nome de usuário já existe.")

        password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        # Append new user data to the sheet
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
    Loads tool data from the dedicated tools Excel file.
    Corrected path to 'app_sheets/tools.xlsx' and added error handling.
    """
    tools = {}
    try:
        if not os.path.exists(TOOLS_EXCEL_PATH):
            QMessageBox.critical(None, "Arquivo Não Encontrado", f"O arquivo de ferramentas não foi encontrado em: {TOOLS_EXCEL_PATH}. Por favor, certifique-se de que ele exista.")
            return {}

        wb = openpyxl.load_workbook(TOOLS_EXCEL_PATH)
        sheet = wb["tools"] # Corrected to read from 'tools' sheet
        
        # Check if sheet has enough rows (at least header + one data row)
        if sheet.max_row < 2:
            QMessageBox.warning(None, "Planilha Vazia", f"A planilha 'tools' em {TOOLS_EXCEL_PATH} parece estar vazia ou conter apenas cabeçalhos.")
            return {}

        for row in sheet.iter_rows(min_row=2):
            # Ensure enough cells are present to avoid IndexError
            if len(row) >= 4 and all(cell.value is not None for cell in row[:4]): # Ensure ID, Name, Desc, Path exist
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
    """Loads role permissions from the database Excel file."""
    perms = {}
    try:
        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        sheet = wb["access"]
        perms = {}
        # Iterate from the second row to skip headers
        for row in sheet.iter_rows(min_row=2):
            # Check if row has enough cells and value is not None
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
    """Loads workspace items from workspace_data.xlsx, creating it if necessary."""
    items = []
    try:
        if not os.path.exists(WORKSPACE_EXCEL_PATH):
            # Create the file and sheet with default data
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "items"
            ws.append(["ID", "Name", "Type", "ParentID", "Description"])
            ws.append(["PROJ-001", "Demo Project - Rev A", "Project", "ROOT", "Main project for demonstration"])
            ws.append(["PART-001", "Part-001", "Part", "PROJ-001", "A manufactured part"])
            ws.append(["ASSY-001", "Assembly-001", "Assembly", "PROJ-001", "An assembly of multiple parts"])
            ws.append(["COMP-001", "Component-XYZ", "Component", "ASSY-001", "A standard component"])
            ws.append(["VAR-001", "Sample Variant - V1.0", "Variant", "ROOT", "A product variant"])
            ws.append(["DRAW-001", "Drawing-CAD-001", "Document", "PART-001", "CAD drawing for Part-001"])
            
            # Create structure sheet
            ws_structure = wb.create_sheet("structure")
            ws_structure.append(["ParentID", "ComponentID", "ComponentName", "Quantity", "Unit", "Type"])
            ws_structure.append(["ASSY-001", "PART-001", "Part-001", 2, "PCS", "Part"])
            ws_structure.append(["ASSY-001", "COMP-001", "Component-XYZ", 1, "PCS", "Component"])
            ws_structure.append(["PROJ-001", "ASSY-001", "Assembly-001", 1, "EA", "Assembly"])
            ws_structure.append(["PROJ-001", "DRAW-001", "Drawing-CAD-001", 1, "EA", "Document"])
            ws_structure.append(["VAR-001", "PART-001", "Part-001", 1, "PCS", "Part"])
            ws_structure.append(["VAR-001", "DRAW-001", "Drawing-CAD-001", 1, "EA", "Document"])

            wb.save(WORKSPACE_EXCEL_PATH)
            QMessageBox.information(None, "Arquivo de Workspace Criado", f"O arquivo '{WORKSPACE_EXCEL_PATH}' foi criado com dados de exemplo.")
        
        wb = openpyxl.load_workbook(WORKSPACE_EXCEL_PATH)
        if "items" not in wb.sheetnames:
            QMessageBox.warning(None, "Planilha 'items' Não Encontrada", f"A planilha 'items' não foi encontrada em '{WORKSPACE_EXCEL_PATH}'. Criando uma nova.")
            ws = wb.create_sheet("items")
            ws.append(["ID", "Name", "Type", "ParentID", "Description"])
            wb.save(WORKSPACE_EXCEL_PATH)

        sheet = wb["items"]
        headers = [cell.value for cell in sheet[1]] # Get headers
        
        # Map header names to column indices for robust access
        header_map = {header: idx for idx, header in enumerate(headers)}
        
        for row_idx in range(2, sheet.max_row + 1): # Start from row 2 for data
            row_values = [cell.value for cell in sheet[row_idx]]
            item_data = {}
            for col_name, col_idx in header_map.items():
                item_data[col_name] = row_values[col_idx] if col_idx < len(row_values) else None
            items.append(item_data)
            
    except Exception as e:
        QMessageBox.critical(None, "Erro ao Carregar Itens do Workspace", f"Erro: {e}")
    return items

def add_workspace_item_to_excel(item_data):
    """Adds a new item to the 'items' sheet in workspace_data.xlsx."""
    try:
        wb = openpyxl.load_workbook(WORKSPACE_EXCEL_PATH)
        sheet = wb["items"]

        # Check for duplicate ID
        for row in sheet.iter_rows(min_row=2):
            if row[0].value == item_data["ID"]:
                QMessageBox.warning(None, "ID Duplicado", f"Um item com o ID '{item_data['ID']}' já existe. Por favor, use um ID único.")
                return False

        # Append the new item data
        # Ensure the order matches the headers: "ID", "Name", "Type", "ParentID", "Description"
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


# === LOGIN WINDOW ===
class LoginWindow(QWidget):
    """
    The login window for the application.
    Handles user authentication and registration.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("5revolution Login")
        self.setGeometry(400, 200, 300, 180) # x, y, width, height
        self.users = load_users_from_excel() # Load users on initialization

        self._init_ui()

    def _init_ui(self):
        """Initializes the UI elements for the login window."""
        layout = QVBoxLayout()

        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Nome de Usuário")
        # Connect returnPressed to authenticate
        self.username_input.returnPressed.connect(self.authenticate)
        
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Senha")
        self.password_input.setEchoMode(QLineEdit.Password)
        # Connect returnPressed to authenticate
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
        """Authenticates the user based on provided credentials."""
        uname = self.username_input.text().strip()
        pwd = self.password_input.text().strip()

        if not uname or not pwd:
            QMessageBox.warning(self, "Falha no Login", "Nome de usuário e senha não podem estar vazios.")
            return

        user = self.users.get(uname)

        if not user or not bcrypt.checkpw(pwd.encode(), user["password_hash"].encode()):
            QMessageBox.warning(self, "Falha no Login", "Nome de usuário ou senha inválidos.")
            return

        # If authentication is successful, launch the main application
        self.main = TeamcenterStyleGUI(user)
        self.main.show()
        self.close() # Close the login window

    def handle_register(self):
        """Handles user registration."""
        uname = self.username_input.text().strip()
        pwd = self.password_input.text().strip()

        if not uname or not pwd:
            QMessageBox.warning(self, "Erro de Validação", "Nome de usuário e senha são obrigatórios para o registro.")
            return

        try:
            register_user(uname, pwd)
            QMessageBox.information(self, "Registrado", f"Usuário '{uname}' registrado com sucesso com o papel 'user'.")
            self.users = load_users_from_excel() # Reload users after registration
            self.username_input.clear()
            self.password_input.clear()
        except ValueError as ve:
            QMessageBox.warning(self, "Falha no Registro", str(ve))
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro durante o registro: {e}")

# === MAIN GUI ===
class TeamcenterStyleGUI(QMainWindow):
    """
    The main application GUI, styled to resemble Teamcenter.
    Provides a workspace tree view, tabbed content area, and a toolbar.
    """
    def __init__(self, user):
        super().__init__()
        self.setWindowTitle("Plataforma 5revolution")
        self.setGeometry(100, 100, 1280, 800) # x, y, width, height

        self.username = user["username"]
        self.role = user["role"]
        self.tools = load_tools_from_excel() # Load tools using the updated function
        self.permissions = load_role_permissions()
        # Initial load of workspace data; will be reloaded on tree population
        self.workspace_items_data = [] 

        self._create_toolbar()
        self._create_main_layout()

        # Display user information in status bar
        self.statusBar().showMessage(f"Logado como: {self.username} | Papel: {self.role}")

    def _create_toolbar(self):
        """Creates the main application toolbar."""
        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setMovable(False) # Make toolbar fixed
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)

        # 🛠 Tools Menu Button
        self.tools_btn = QToolButton()
        self.tools_btn.setText("🛠 Ferramentas")
        self.tools_btn.setPopupMode(QToolButton.InstantPopup) # Shows menu instantly on click
        tools_menu = QMenu()

        allowed_tools = self.permissions.get(self.role, []) # Get allowed tools for the user's role
        for tid, tool in self.tools.items():
            # Check if user has permission for this tool or if role is 'all'
            if allowed_tools == "all" or tid in allowed_tools:
                action = tools_menu.addAction(tool["name"])
                
                # Dynamically connect actions to the correct tool widgets
                if tool["id"] == "mod4": # Engenharia (Workflow)
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, EngenhariaWorkflowTool()))
                elif tool["id"] == "mes_pcp": # MES (Apontamento Fábrica)
                    # For MES, we will create a dedicated widget, perhaps using a placeholder class for now
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
                else: # Generic tool
                    action.triggered.connect(lambda chk=False, title=tool["name"], desc=tool["description"]: self._open_tab(title, QLabel(desc)))
        self.tools_btn.setMenu(tools_menu)
        self.toolbar.addWidget(self.tools_btn)

        # 👤 Profile Menu Button
        self.profile_btn = QToolButton()
        self.profile_btn.setText(f"👤 {self.username}") # Display username in profile button
        self.profile_btn.setPopupMode(QToolButton.InstantPopup)
        profile_menu = QMenu()
        # Connect "⚙️ Configurações" to open the new UserSettingsTool
        profile_menu.addAction("⚙️ Configurações", lambda: self._open_tab("Configurações do Usuário", UserSettingsTool(self.username, self.role)))
        profile_menu.addSeparator() # Add a separator for better visual grouping
        profile_menu.addAction("🔒 Sair", self._logout)
        self.profile_btn.setMenu(profile_menu)
        self.toolbar.addWidget(self.profile_btn)

    def _create_main_layout(self):
        """Creates the main split layout with tree view and tabs."""
        self.splitter = QSplitter() # Allows resizing of sub-widgets

        # Left Pane Widget Container (to add search bar easily)
        left_pane_widget = QWidget()
        left_pane_layout = QVBoxLayout(left_pane_widget)
        left_pane_layout.setContentsMargins(0, 0, 0, 0) # Remove margins for cleaner look

        # 🌳 Tree View (Left Pane)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Espaço de Trabalho")
        self._populate_workspace_tree() # Populate with data from Excel
        self.tree.expandAll() # Expand all tree items by default
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu) # Enable custom context menu
        self.tree.customContextMenuRequested.connect(self._show_tree_context_menu)
        # Connect double-click to open structure view
        self.tree.itemDoubleClicked.connect(self._open_structure_view_for_item)


        # Search Bar
        search_layout = QHBoxLayout()
        self.item_search_bar = QLineEdit()
        self.item_search_bar.setPlaceholderText("Pesquisar itens...")
        self.item_search_bar.returnPressed.connect(self.handle_item_search) # Connect Enter key
        self.search_items_btn = QPushButton("🔍")
        self.search_items_btn.clicked.connect(self.handle_item_search)

        search_layout.addWidget(self.item_search_bar)
        search_layout.addWidget(self.search_items_btn)

        # Add search bar and tree to the left pane layout
        left_pane_layout.addWidget(QLabel("Espaço de Trabalho")) # Label above search bar
        left_pane_layout.addLayout(search_layout)
        left_pane_layout.addWidget(self.tree)


        # 📑 Tabs (Right Pane)
        self.tabs = QTabWidget()
        self.tabs.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tabs.customContextMenuRequested.connect(self._show_tab_context_menu)
        self.tabs.setTabsClosable(True) # Make tabs closable by default
        self.tabs.tabCloseRequested.connect(self.tabs.removeTab) # Connect close button to remove tab

        # Welcome/Home Tab
        welcome_widget = QWidget()
        welcome_layout = QVBoxLayout()
        welcome_layout.addWidget(QLabel(f"Bem-vindo {self.username} – Papel: {self.role}"))
        welcome_widget.setLayout(welcome_layout)
        self.tabs.addTab(welcome_widget, "Início")

        # Add widgets to the splitter
        self.splitter.addWidget(left_pane_widget) # Add the container widget to the splitter
        self.splitter.addWidget(self.tabs)
        self.splitter.setStretchFactor(1, 4) # Give more space to the tabs

        # Set splitter as the central widget
        container = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(self.splitter)
        container.setLayout(layout)
        self.setCentralWidget(container)

    def _populate_workspace_tree(self):
        """Populates the tree with data loaded from workspace_data.xlsx."""
        self.tree.clear() # Clear existing items
        self.workspace_items_data = load_workspace_items_from_excel() # Reload data

        # Build a dictionary for quick lookup of items by ID and to store children
        item_map = {item['ID']: {'data': item, 'children': []} for item in self.workspace_items_data}
        
        # Create root items (those with 'ROOT' as ParentID) and establish hierarchy
        root_items = []
        for item_id, item_info in item_map.items():
            parent_id = item_info['data'].get('ParentID') # Use .get() for safety
            if parent_id == 'ROOT':
                root_items.append(item_info['data'])
            elif parent_id in item_map:
                item_map[parent_id]['children'].append(item_info['data'])

        # Recursive function to add items to the tree
        def add_items_to_tree(parent_qtree_item, items_list):
            for item_data in items_list:
                icon_map = {
                    "Project": "📁", "Assembly": "📦", "Part": "📄",
                    "Component": "🧩", "Document": "📜", "Variant": "💡"
                }
                icon = icon_map.get(item_data.get('Type', ''), '❓') # Default icon
                
                # Store the full item data directly in the QTreeWidgetItem
                q_item = QTreeWidgetItem([f"{icon} {item_data.get('Name', 'N/A')}"])
                q_item.setData(0, Qt.UserRole, item_data.get('ID')) # Store ID for later lookup
                q_item.setData(1, Qt.UserRole, item_data.get('Name')) # Store Name for easier retrieval
                q_item.setData(2, Qt.UserRole, item_data.get('Type')) # Store Type
                q_item.setData(3, Qt.UserRole, item_data.get('ParentID')) # Store ParentID
                q_item.setData(4, Qt.UserRole, item_data.get('Description')) # Store Description
                
                parent_qtree_item.addChild(q_item)
                
                # Recursively add children
                if item_data.get('ID') in item_map and item_map[item_data['ID']]['children']:
                    add_items_to_tree(q_item, item_map[item_data['ID']]['children'])

        # Add root items
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
        
        self.tree.expandAll() # Expand all nodes by default for visibility


    def _open_structure_view_for_item(self, item, column):
        """
        Opens a new tab with the StructureViewTool for the double-clicked item.
        """
        item_id = item.data(0, Qt.UserRole) # Retrieve the stored ID
        if not item_id:
            QMessageBox.warning(self, "Erro", "Não foi possível obter o ID do item para visualização da estrutura.")
            return

        item_name = item.text(0).split(' ', 1)[1].strip() # Get the name from the display text

        tab_title = f"Estrutura: {item_name}"
        # Check if tab is already open
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == tab_title:
                self.tabs.setCurrentIndex(i)
                QMessageBox.information(self, "Informação", f"A guia para a estrutura de '{item_name}' já está aberta.")
                return

        # Create and open the StructureViewTool in a new tab
        structure_tool = StructureViewTool(item_id, item_name)
        self._open_tab(tab_title, structure_tool)


    def _create_mes_pcp_tool_widget(self):
        """Creates the widget for the MES (Apontamento Fábrica) tool."""
        mes_widget = QWidget()
        mes_layout = QVBoxLayout()
        mes_layout.addWidget(QLabel("<h2>MES (Apontamento Fábrica)</h2>"))
        mes_layout.addWidget(QLabel("Inserir dados de produção, acompanhar progresso e gerenciar operações de chão de fábrica."))

        form_layout = QVBoxLayout()
        # Using self.mes_ prefixes to avoid naming conflicts with other modules if they were directly in GUI
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
        submit_btn.clicked.connect(self._submit_mes_data) # Connect to a submission handler

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
        mes_layout.addStretch() # Push content to top
        mes_widget.setLayout(mes_layout)
        return mes_widget

    def _submit_mes_data(self):
        """Handles submission of MES data (placeholder)."""
        order_id = self.mes_order_id_input.text()
        item_code = self.mes_item_code_input.text()
        quantity = self.mes_quantity_input.text()
        start_time = self.mes_start_time_input.text()
        end_time = self.mes_end_time_input.text()

        if not all([order_id, item_code, quantity, start_time, end_time]):
            QMessageBox.warning(self, "Erro de Entrada", "Todos os campos MES devem ser preenchidos.")
            return

        # In a real application, you would save this data to a database or file
        QMessageBox.information(self, "Dados MES Enviados",
                                f"Dados de Produção Enviados:\n"
                                f"ID da Ordem: {order_id}\n"
                                f"Código do Item: {item_code}\n"
                                f"Quantidade: {quantity}\n"
                                f"Início: {start_time}\n"
                                f"Término: {end_time}")
        # Clear fields after submission
        self.mes_order_id_input.clear()
        self.mes_item_code_input.clear()
        self.mes_quantity_input.clear()
        self.mes_start_time_input.clear()
        self.mes_end_time_input.clear()

    def handle_item_search(self):
        """
        Performs a search on workspace items and displays results in a dialog.
        """
        search_term = self.item_search_bar.text().strip().lower()
        if not search_term:
            QMessageBox.information(self, "Pesquisar", "Por favor, digite um termo de pesquisa.")
            return

        # Filter self.workspace_items_data
        results = [item for item in self.workspace_items_data if search_term in item.get('Name', '').lower() or search_term in item.get('ID', '').lower()]
        self.display_search_results_dialog(results)

    def display_search_results_dialog(self, results):
        """
        Displays search results in a new QDialog window.
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("Resultados da Pesquisa")
        dialog.setGeometry(self.x() + 200, self.y() + 100, 400, 300) # Position relative to main window

        layout = QVBoxLayout(dialog)
        
        if not results:
            layout.addWidget(QLabel("Nenhum item encontrado correspondente à sua pesquisa."))
        else:
            list_widget = QListWidget()
            for item in results:
                # Display Name (ID) for clarity
                list_item_text = f"{item.get('Name', 'N/A')} ({item.get('ID', 'N/A')})"
                list_item = QListWidgetItem(list_item_text)
                list_item.setData(Qt.UserRole, item.get('ID')) # Store ID in UserRole
                list_item.setData(Qt.UserRole + 1, item.get('Name')) # Store Name for easier retrieval
                list_widget.addItem(list_item)

            list_widget.itemDoubleClicked.connect(
                # Use the ID and Name from the stored data to open structure view
                lambda item_list_widget: self._open_structure_view_for_item_by_data(item_list_widget.data(Qt.UserRole), item_list_widget.data(Qt.UserRole + 1)) or dialog.accept()
            ) 
            layout.addWidget(list_widget)

        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(dialog.accept) # Close dialog on button click
        layout.addWidget(close_btn)

        dialog.exec_() # Show dialog modally

    def _open_structure_view_for_item_by_data(self, item_id, item_name):
        """Helper to open structure view directly from ID and Name."""
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
        Helper to find a QTreeWidgetItem by its stored ID (Qt.UserRole).
        Performs a recursive search through the tree.
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

        # Start search from top-level items
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
        Opens a new tab or switches to an existing one.
        Accepts a widget instance directly.
        """
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == title:
                self.tabs.setCurrentIndex(i)
                return
        # If tab doesn't exist, add it
        self.tabs.addTab(widget_instance, title)
        self.tabs.setCurrentIndex(self.tabs.count() - 1) # Switch to the newly opened tab

    def _open_options(self):
        """Opens the user options/settings dialog."""
        # Use the already defined UserSettingsTool
        self._open_tab("Configurações do Usuário", UserSettingsTool(self.username, self.role))

    def _logout(self):
        """Logs out the current user and returns to the login screen."""
        confirm_logout = QMessageBox.question(self, "Confirmação de Saída", "Tem certeza de que deseja sair?",
                                              QMessageBox.Yes | QMessageBox.No)
        if confirm_logout == QMessageBox.Yes:
            self.close() # Close the main application window
            self.login = LoginWindow() # Create a new login window instance
            self.login.show() # Show the login window

    def _show_tree_context_menu(self, pos):
        """Displays a context menu for items in the tree view."""
        item = self.tree.itemAt(pos)
        
        menu = QMenu()

        # Action to add a new top-level item (e.g., a new project or variant)
        # This will appear regardless of whether an item is selected or not
        add_root_item_action = menu.addAction("➕ Adicionar Novo Item Raiz")
        add_root_item_action.triggered.connect(self._add_new_root_item)

        if item: # If an item was clicked
            item_id = item.data(0, Qt.UserRole)
            item_name = item.data(1, Qt.UserRole) or item.text(0).split(' ', 1)[1].strip() # Fallback to text if name not set

            menu.addSeparator() # Separator for item-specific actions
            
            # Context actions for any clicked item
            menu.addAction("🔍 Ver Detalhes/Estrutura", lambda: self._open_structure_view_for_item(item, 0))
            menu.addAction("✏️ Editar Propriedades (Simulado)", lambda: QMessageBox.information(self, "Ação Similada", f"Editando propriedades para: {item_name} (ação simulada)"))
            menu.addAction("❌ Excluir Item (Simulado)", lambda: QMessageBox.warning(self, "Ação Similada", f"Excluindo: {item_name} (ação simulada)"))
            
            # Action to add a subitem, specific to the clicked item
            add_subitem_action = menu.addAction("➕ Adicionar Subitem")
            add_subitem_action.triggered.connect(lambda: self._add_new_subitem(item_id, item_name))

        menu.exec_(self.tree.viewport().mapToGlobal(pos)) # Show menu at mouse position

    def _add_new_root_item(self):
        """Opens a dialog to add a new top-level item to the workspace."""
        dialog = AddItemDialog(parent_id="ROOT", parent_name="ROOT", parent=self)
        if dialog.exec_() == QDialog.Accepted:
            item_data = dialog.item_data
            if add_workspace_item_to_excel(item_data):
                self._populate_workspace_tree() # Refresh tree after adding
    
    def _add_new_subitem(self, parent_id, parent_name):
        """Opens a dialog to add a new subitem to a selected parent in the workspace."""
        dialog = AddItemDialog(parent_id=parent_id, parent_name=parent_name, parent=self)
        if dialog.exec_() == QDialog.Accepted:
            item_data = dialog.item_data
            if add_workspace_item_to_excel(item_data):
                self._populate_workspace_tree() # Refresh tree after adding


    def _show_tab_context_menu(self, pos):
        """Displays a context menu for tabs in the tab widget."""
        index = self.tabs.tabBar().tabAt(pos)
        if index < 0: return # No tab clicked

        menu = QMenu()
        menu.addAction("❌ Fechar Guia", lambda: self.tabs.removeTab(index))
        # Ensure "Fechar Outras Guias" doesn't close the current tab if it's the only one
        if self.tabs.count() > 1:
            menu.addAction("🔁 Fechar Outras Guias", lambda: self._close_other_tabs(index))
        if self.tabs.count() > 0: # Only show "Fechar Todas as Guias" if there are tabs
            menu.addAction("🧹 Fechar Todas as Guias", self.tabs.clear)
        menu.exec_(self.tabs.tabBar().mapToGlobal(pos))

    def _close_other_tabs(self, keep_index):
        """Closes all tabs except the one at 'keep_index'."""
        # Iterate in reverse to avoid index issues when removing tabs
        for i in reversed(range(self.tabs.count())):
            if i != keep_index:
                self.tabs.removeTab(i)

# === ENTRYPOINT ===
if __name__ == "__main__":
    app = QApplication(sys.argv)
    login = LoginWindow()
    login.show()
    sys.exit(app.exec_())
