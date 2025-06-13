import sys
import os
import bcrypt
import openpyxl
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QAction, QTabWidget, QMenu, QToolButton,
    QWidget, QVBoxLayout, QSplitter, QTreeWidget, QTreeWidgetItem,
    QLabel, QLineEdit, QPushButton, QHBoxLayout, QMessageBox, QGraphicsView,
    QGraphicsScene, QGraphicsRectItem, QGraphicsLineItem, QDialog, QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt5.QtCore import Qt, QPointF
from PyQt5.QtGui import QBrush, QPen, QColor

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

# --- File Paths Configuration ---
# Define standard paths for consistency.
# These paths are now relative to the project root, which is in sys.path
USER_SHEETS_DIR = os.path.join(project_root, "user_sheets")
APP_SHEETS_DIR = os.path.join(project_root, "app_sheets")
DB_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "db.xlsx")
TOOLS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "tools.xlsx") # New path for tools.xlsx

# Ensure directories exist
os.makedirs(USER_SHEETS_DIR, exist_ok=True)
os.makedirs(APP_SHEETS_DIR, exist_ok=True)

# Sample hardcoded workspace items for search (in a real app, these would come from a backend/database)
WORKSPACE_ITEMS = [
    "Demo Project - Rev A",
    "Part-001",
    "Assembly-001",
    "Sample Variant - V1.0",
    "Component-XYZ",
    "Specification-005",
    "Drawing-CAD-001"
]

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

# === NEW TOOL: ENGENHARIA WORKFLOW DIAGRAM ===
class EngenhariaWorkflowTool(QWidget):
    """
    A placeholder widget for the Engenharia Workflow Diagram tool.
    Provides a basic QGraphicsView for diagramming.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Engenharia (Workflow) Tool")
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        self.layout.addWidget(self.view)

        self._add_sample_diagram_elements()

        # Add control buttons
        control_layout = QHBoxLayout()
        add_node_btn = QPushButton("Adicionar Nó de Tarefa")
        add_node_btn.clicked.connect(self._add_task_node)
        add_link_btn = QPushButton("Adicionar Ligação de Dependência")
        add_link_btn.clicked.connect(self._add_dependency_link)
        clear_btn = QPushButton("Limpar Diagrama")
        clear_btn.clicked.connect(self._clear_diagram)

        control_layout.addWidget(add_node_btn)
        control_layout.addWidget(add_link_btn)
        control_layout.addWidget(clear_btn)
        self.layout.addLayout(control_layout)

        self.nodes = [] # To keep track of added nodes

    def _add_sample_diagram_elements(self):
        """Adds some sample elements to the diagram scene."""
        # Task nodes - CORRECTED COLOR USAGE
        node1 = self.scene.addRect(50, 50, 100, 50, QPen(Qt.black), QBrush(QColor("lightblue")))
        node2 = self.scene.addRect(200, 150, 100, 50, QPen(Qt.black), QBrush(QColor("lightgreen")))
        node3 = self.scene.addRect(350, 50, 100, 50, QPen(Qt.black), QBrush(QColor("lightcoral")))

        self.scene.addText("Fase de Design", QPointF(55, 65))
        self.scene.addText("Revisão (Aprovado)", QPointF(205, 165))
        self.scene.addText("Preparação da Produção", QPointF(355, 65))

        # Links/Arrows
        pen = QPen(Qt.darkGray)
        pen.setWidth(2)
        self.scene.addLine(node1.x() + node1.rect().width(), node1.y() + node1.rect().height() / 2,
                           node2.x(), node2.y() + node2.rect().height() / 2, pen)
        self.scene.addLine(node2.x() + node2.rect().width(), node2.y() + node2.rect().height() / 2,
                           node3.x(), node3.y() + node3.rect().height() / 2, pen)

    def _add_task_node(self):
        """Adds a new generic task node to the diagram."""
        x = 10 + len(self.nodes) * 120 # Offset for new nodes
        y = 10 + (len(self.nodes) % 3) * 70
        node = self.scene.addRect(x, y, 100, 50, QPen(Qt.black), QBrush(QColor("#FFD700"))) # Gold color
        self.scene.addText(f"Nova Tarefa {len(self.nodes) + 1}", QPointF(x + 5, y + 15))
        self.nodes.append(node)
        self.view.centerOn(node)

    def _add_dependency_link(self):
        """Prompts user to select two nodes to link. (Conceptual, requires selection logic)."""
        QMessageBox.information(self, "Adicionar Ligação", "Clique em dois nós de tarefa para criar uma ligação. (Lógica de seleção a ser implementada)")
        # In a real implementation, you'd need selection mechanisms (e.g., click listeners on QGraphicsRectItem)
        # to get two nodes and then draw a QGraphicsLineItem between their centroids or edges.

    def _clear_diagram(self):
        """Clears all elements from the diagram."""
        self.scene.clear()
        self.nodes = [] # Reset nodes list
        QMessageBox.information(self, "Diagrama Limpo", "O diagrama foi limpo.")


# === NEW TOOL: User Settings ===
class UserSettingsTool(QWidget):
    """
    A widget to display user profile and personal information in read-only mode.
    """
    def __init__(self, username, role):
        super().__init__()
        self.setWindowTitle("Configurações do Usuário")
        self.layout = QVBoxLayout(self)

        self.layout.addWidget(QLabel("<h2>Informações do Perfil</h2>"))

        # Username (read-only)
        username_layout = QHBoxLayout()
        username_layout.addWidget(QLabel("Nome de Usuário:"))
        self.username_display = QLineEdit(username)
        self.username_display.setReadOnly(True)
        username_layout.addWidget(self.username_display)
        self.layout.addLayout(username_layout)

        # Role (read-only)
        role_layout = QHBoxLayout()
        role_layout.addWidget(QLabel("Cargo/Função:"))
        self.role_display = QLineEdit(role)
        self.role_display.setReadOnly(True)
        role_layout.addWidget(self.role_display)
        self.layout.addLayout(role_layout)

        # Add some stretch to push content to the top
        self.layout.addStretch()

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
        self._populate_sample_tree() # Populate with sample data
        self.tree.expandAll() # Expand all tree items by default
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu) # Enable custom context menu
        self.tree.customContextMenuRequested.connect(self._show_tree_context_menu)

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

    def _populate_sample_tree(self):
        """Populates the tree with sample project/variant data."""
        root = QTreeWidgetItem(["Projetos"])
        project1 = QTreeWidgetItem(["Projeto Demo - Rev A"])
        project1.addChild(QTreeWidgetItem(["Peça-001"]))
        project1.addChild(QTreeWidgetItem(["Montagem-001"]))
        self.tree.addTopLevelItem(root)

        project2 = QTreeWidgetItem(["Variante Amostra - V1.0"])
        project2.addChild(QTreeWidgetItem(["Componente-XYZ"]))
        self.tree.addTopLevelItem(project2) # Added directly to root for testing different structures

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

        results = [item for item in WORKSPACE_ITEMS if search_term in item.lower()]
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
                list_widget.addItem(item)
            list_widget.itemDoubleClicked.connect(
                lambda item: self.open_selected_item_tab(item.text()) or dialog.accept()
            ) # Close dialog on double click
            layout.addWidget(list_widget)

        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(dialog.accept) # Close dialog on button click
        layout.addWidget(close_btn)

        dialog.exec_() # Show dialog modally

    def open_selected_item_tab(self, item_name):
        """
        Opens a new tab in the main GUI to display details of the selected item.
        """
        tab_id = f"item-details-{item_name.replace(' ', '-')}"
        tab_title = f"Detalhes: {item_name}"

        # Check if tab is already open
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == tab_title:
                self.tabs.setCurrentIndex(i)
                QMessageBox.information(self, "Informação", f"A guia para '{item_name}' já está aberta.")
                return

        # Create a widget for item details
        item_details_widget = QWidget()
        item_details_layout = QVBoxLayout()
        item_details_layout.addWidget(QLabel(f"<h2>Detalhes do Item: {item_name}</h2>"))
        item_details_layout.addWidget(QLabel(f"Exibindo detalhes abrangentes para <b>{item_name}</b>."))
        item_details_layout.addWidget(QLabel("Esta seção carregaria dados reais: propriedades, revisões, arquivos associados, etc."))
        item_details_layout.addStretch() # Push content to top
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
        if not item: return

        menu = QMenu()
        # Actions for root items (e.g., "Projetos")
        if item.parent() is None:
            menu.addAction("🔁 Atualizar Projeto", lambda: QMessageBox.information(self, "Ação Similada", "Projeto atualizado (ação simulada)"))
            menu.addAction("➕ Adicionar Novo Item", lambda: QMessageBox.information(self, "Ação Similada", "Adicionar novo item (ação simulada)"))
        # Actions for child items (e.g., "Projeto Demo", "Peça-001")
        else:
            menu.addAction("🔍 Ver Detalhes", lambda: QMessageBox.information(self, "Ação Similada", f"Visualizando detalhes para: {item.text(0)} (ação simulada)"))
            menu.addAction("✏️ Editar Propriedades", lambda: QMessageBox.information(self, "Ação Similada", f"Editando propriedades para: {item.text(0)} (ação simulada)"))
            menu.addAction("❌ Excluir Item", lambda: QMessageBox.warning(self, "Ação Similada", f"Excluído: {item.text(0)} (ação simulada)"))

        menu.exec_(self.tree.viewport().mapToGlobal(pos)) # Show menu at mouse position

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
