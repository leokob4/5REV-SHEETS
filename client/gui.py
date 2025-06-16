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
from ui.tools.excel_viewer_tool import ExcelViewerTool # Ferramenta genérica de visualização de Excel
from ui.tools.rpi_tool import RpiTool # Nova ferramenta para RPI.xlsx

# Importa o diálogo de adição de item (manter se ainda for útil para outras funções)
from client.add_item_dialog import AddItemDialog 

# --- Configuração de Caminhos de Arquivos ---
# Define paths padrão para consistência.
USER_SHEETS_DIR = os.path.join(project_root, "user_sheets")
APP_SHEETS_DIR = os.path.join(project_root, "app_sheets")
DB_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "db.xlsx")
TOOLS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "tools.xlsx")
# WORKSPACE_EXCEL_PATH foi tornado obsoleto e removido.

# Garante que os diretórios existam
os.makedirs(USER_SHEETS_DIR, exist_ok=True)
os.makedirs(APP_SHEETS_DIR, exist_ok=True)

# --- Mapeamento de Nomes de Arquivo Excel para Ferramentas Específicas ---
# Adicione mais mapeamentos conforme você implementa cada ferramenta
# O nome do arquivo Excel (minúsculas) deve corresponder à chave.
EXCEL_TOOL_MAP = {
    "output.xlsx": ProductDataTool,
    "bom_data.xlsx": BomManagerTool,
    "configurador.xlsx": ConfiguradorTool, # Assumindo o nome do arquivo, confirme se é este
    "colaboradores.xlsx": ColaboradoresTool, # Assumindo o nome do arquivo, confirme se é este
    "items_data.xlsx": ItemsTool,
    "manufacturing_data.xlsx": ManufacturingTool,
    "programacao.xlsx": PcpTool, # Mapeando programacao.xlsx para PcpTool
    "estoque_data.xlsx": EstoqueTool,
    "financeiro.xlsx": FinanceiroTool, # Assumindo o nome do arquivo, confirme se é este
    "pedidos_data.xlsx": PedidosTool,
    "manutencao_data.xlsx": ManutencaoTool,
    "rpi.xlsx": RpiTool, # Mapeando RPI.xlsx para a nova RpiTool
    # 'db.xlsx' e 'tools.xlsx' são arquivos do sistema, não devem ser abertos diretamente por este mapeamento.
    # 'engenharia_workflow_tool' e 'user_settings_tool' não são diretamente baseadas em arquivos Excel.
}


# === FUNÇÕES DE AJUDA PARA PLANILHAS ===
def load_users_from_excel():
    """Carrega dados de usuários do arquivo Excel do banco de dados."""
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
    """Registra um novo usuário no arquivo Excel do banco de dados."""
    try:
        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        sheet = wb["users"]
        next_id = sheet.max_row
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
    """Carrega permissões de função do arquivo Excel do banco de dados."""
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

# === JANELA DE LOGIN ===
class LoginWindow(QWidget):
    """
    A janela de login para a aplicação.
    Gerencia a autenticação e o registro de usuários.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("5revolution Login")
        self.setGeometry(400, 200, 300, 180) # x, y, width, height
        self.users = load_users_from_excel() # Load users on initialization

        self._init_ui()

    def _init_ui(self):
        """Inicializa os elementos da interface de usuário para a janela de login."""
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

        # If authentication is successful, launch the main application
        self.main = TeamcenterStyleGUI(user)
        self.main.show()
        self.close() # Close the login window

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
            self.users = load_users_from_excel() # Reload users after registration
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
        """Cria a barra de ferramentas principal da aplicação."""
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
                
                # Conecta dinamicamente ações aos widgets das ferramentas corretas
                if tool["id"] == "mod4": # Engenharia (Workflow)
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, EngenhariaWorkflowTool()))
                elif tool["id"] == "mes_pcp": # MES (Apontamento Fábrica)
                    # Para MES, vamos criar um widget dedicado, talvez usando uma classe placeholder por agora
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, self._create_mes_pcp_tool_widget()))
                elif tool["id"] == "prod_data": # Dados do Produto
                    # ProductDataTool agora espera um file_path; vamos passar o default output.xlsx
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ProductDataTool(file_path=os.path.join(USER_SHEETS_DIR, "output.xlsx"))))
                elif tool["id"] == "bom_manager": # Gerenciador de BOM
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, BomManagerTool(file_path=os.path.join(USER_SHEETS_DIR, "bom_data.xlsx"))))
                elif tool["id"] == "configurador": # Configurador
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ConfiguradorTool(file_path=os.path.join(USER_SHEETS_DIR, "configurador.xlsx"))))
                elif tool["id"] == "colab": # Colaboradores
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ColaboradoresTool(file_path=os.path.join(USER_SHEETS_DIR, "colaboradores.xlsx"))))
                elif tool["id"] == "items_tool": # Itens
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ItemsTool(file_path=os.path.join(USER_SHEETS_DIR, "items_data.xlsx"))))
                elif tool["id"] == "manuf": # Fabricação
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ManufacturingTool(file_path=os.path.join(USER_SHEETS_DIR, "manufacturing_data.xlsx"))))
                elif tool["id"] == "pcp_tool": # PCP
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, PcpTool(file_path=os.path.join(USER_SHEETS_DIR, "programacao.xlsx"))))
                elif tool["id"] == "estoque_tool": # Estoque
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, EstoqueTool(file_path=os.path.join(USER_SHEETS_DIR, "estoque_data.xlsx"))))
                elif tool["id"] == "financeiro": # Financeiro
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, FinanceiroTool(file_path=os.path.join(USER_SHEETS_DIR, "financeiro.xlsx"))))
                elif tool["id"] == "pedidos": # Pedidos
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, PedidosTool(file_path=os.path.join(USER_SHEETS_DIR, "pedidos_data.xlsx"))))
                elif tool["id"] == "manutencao": # Manutenção
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, ManutencaoTool(file_path=os.path.join(USER_SHEETS_DIR, "manutencao_data.xlsx"))))
                else: # Ferramenta Genérica (fallback)
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
        """Cria o layout dividido principal com a visualização em árvore e as abas."""
        self.splitter = QSplitter() # Allows resizing of sub-widgets

        # Left Pane Widget Container (to add search bar easily)
        left_pane_widget = QWidget()
        left_pane_layout = QVBoxLayout(left_pane_widget)
        left_pane_layout.setContentsMargins(0, 0, 0, 0) # Remove margins for cleaner look

        # 🌳 Tree View (Left Pane) - Agora lista arquivos .xlsx
        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Arquivos do Usuário (.xlsx)")
        self._populate_workspace_tree() # Populate with data from Excel
        self.tree.expandAll() # Expand all tree items by default
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu) # Enable custom context menu
        self.tree.customContextMenuRequested.connect(self._show_tree_context_menu)
        # Conecta duplo-clique para abrir o arquivo com a ferramenta apropriada
        self.tree.itemDoubleClicked.connect(self._open_file_from_tree)


        # Search Bar
        search_layout = QHBoxLayout()
        self.item_search_bar = QLineEdit()
        self.item_search_bar.setPlaceholderText("Pesquisar arquivos...")
        self.item_search_bar.returnPressed.connect(self.handle_file_search) # Connect Enter key
        self.search_items_btn = QPushButton("🔍")
        self.search_items_btn.clicked.connect(self.handle_file_search)

        search_layout.addWidget(self.item_search_bar)
        search_layout.addWidget(self.search_items_btn)

        # Add search bar and tree to the left pane layout
        left_pane_layout.addWidget(QLabel("Arquivos do Usuário")) # Label above search bar
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
        """Popula a árvore com os arquivos .xlsx encontrados em USER_SHEETS_DIR."""
        self.tree.clear()
        
        # Filtra para incluir apenas arquivos .xlsx e excluir db.xlsx e tools.xlsx (arquivos do sistema)
        excel_files = [f for f in os.listdir(USER_SHEETS_DIR) if f.endswith('.xlsx') and f not in ['db.xlsx', 'tools.xlsx']]
        
        if not excel_files:
            # Não exibe QMessageBox se a pasta user_sheets estiver vazia, apenas mostra uma árvore vazia
            print(f"Nenhum arquivo .xlsx encontrado na pasta: {USER_SHEETS_DIR}.")
            return

        for filename in sorted(excel_files): # Ordena para exibição consistente
            file_item = QTreeWidgetItem([filename])
            file_item.setData(0, Qt.UserRole, os.path.join(USER_SHEETS_DIR, filename)) # Armazena o caminho completo do arquivo
            self.tree.addTopLevelItem(file_item)
        
        self.tree.expandAll()

    def _open_file_from_tree(self, item, column):
        """
        Abre o arquivo Excel selecionado na árvore com a ferramenta apropriada.
        """
        file_path = item.data(0, Qt.UserRole)
        if not file_path:
            return # Não é um item de arquivo

        file_name = os.path.basename(file_path)
        tool_class = EXCEL_TOOL_MAP.get(file_name.lower()) # Usa .lower() para corresponder ao mapeamento

        tab_title = f"Arquivo: {file_name}"
        # Se a aba já está aberta, apenas ativa-a
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == tab_title:
                self.tabs.setCurrentIndex(i)
                return

        if tool_class:
            # Se existe um mapeamento, instancie a ferramenta específica com o caminho do arquivo
            tool_instance = tool_class(file_path=file_path) # Passa o caminho do arquivo
            self._open_tab(tab_title, tool_instance)
        else:
            # Se não houver mapeamento, abra com o visualizador genérico de Excel
            excel_viewer = ExcelViewerTool(file_path=file_path)
            self._open_tab(tab_title, excel_viewer)


    def _create_mes_pcp_tool_widget(self):
        """Cria o widget para a ferramenta MES (Apontamento Fábrica). (Placeholder)"""
        mes_widget = QWidget()
        mes_layout = QVBoxLayout()
        mes_layout.addWidget(QLabel("<h2>MES (Apontamento Fábrica)</h2>"))
        mes_layout.addWidget(QLabel("Esta é uma ferramenta placeholder para MES. Dados de entrada/saída serão implementados aqui."))

        form_layout = QVBoxLayout()
        # Exemplo de campos para MES, você os implementará de verdade depois
        mes_layout.addWidget(QLabel("ID da Ordem de Produção:"))
        mes_layout.addWidget(QLineEdit("ORDEM-001"))
        mes_layout.addWidget(QLabel("Quantidade Produzida:"))
        mes_layout.addWidget(QLineEdit("100"))
        mes_layout.addWidget(QPushButton("Registrar Produção (Simulado)"))
        
        mes_layout.addLayout(form_layout)
        mes_layout.addStretch() # Push content to top
        mes_widget.setLayout(mes_layout)
        return mes_widget

    def handle_file_search(self):
        """
        Realiza uma busca nos nomes dos arquivos Excel e exibe os resultados.
        """
        search_term = self.item_search_bar.text().strip().lower()
        if not search_term:
            QMessageBox.information(self, "Pesquisar", "Por favor, digite um termo de pesquisa.")
            return

        excel_files = [f for f in os.listdir(USER_SHEETS_DIR) if f.endswith('.xlsx') and f not in ['db.xlsx', 'tools.xlsx']]
        results = [f for f in excel_files if search_term in f.lower()]
        self.display_search_results_dialog(results)

    def display_search_results_dialog(self, results):
        """
        Exibe os resultados da pesquisa em uma nova janela QDialog.
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("Resultados da Pesquisa de Arquivos")
        dialog.setGeometry(self.x() + 200, self.y() + 100, 400, 300)

        layout = QVBoxLayout(dialog)
        
        if not results:
            layout.addWidget(QLabel("Nenhum arquivo encontrado correspondente à sua pesquisa."))
        else:
            list_widget = QListWidget()
            for file_name in results:
                list_item = QListWidgetItem(file_name)
                list_item.setData(Qt.UserRole, os.path.join(USER_SHEETS_DIR, file_name)) # Store full path
                list_widget.addItem(list_item)

            list_widget.itemDoubleClicked.connect(
                lambda item_list_widget: self._open_file_from_tree(item_list_widget, 0) or dialog.accept()
            ) 
            layout.addWidget(list_widget)

        close_btn = QPushButton("Fechar")
        close_btn.clicked.connect(dialog.accept) # Close dialog on button click
        layout.addWidget(close_btn)

        dialog.exec_() # Show dialog modally

    def _open_tab(self, title, widget_instance):
        """
        Abre uma nova aba ou alterna para uma existente.
        Aceita uma instância de widget diretamente.
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
        """Displays a context menu for items (files) in the tree view."""
        item = self.tree.itemAt(pos)
        
        menu = QMenu()

        # Opção para criar um novo arquivo .xlsx vazio (se o usuário quiser iniciar um do zero)
        create_new_file_action = menu.addAction("➕ Criar Novo Arquivo .xlsx")
        create_new_file_action.triggered.connect(self._create_new_excel_file)


        if item: # If a file was clicked
            file_path = item.data(0, Qt.UserRole)
            file_name = os.path.basename(file_path)
            
            menu.addSeparator()
            menu.addAction("🔍 Abrir Arquivo", lambda: self._open_file_from_tree(item, 0))
            
            # Opções para visualizar a estrutura se a planilha 'structure' existir no arquivo
            try:
                wb = openpyxl.load_workbook(file_path)
                if "structure" in wb.sheetnames:
                    menu.addAction("📊 Visualizar Estrutura", lambda: self._open_tab(f"Estrutura: {file_name}", StructureViewTool(file_path=file_path, sheet_name="structure")))
            except Exception:
                pass # Ignora se o arquivo não puder ser lido ou não tiver a aba
            
            # Estas são ações simuladas para Renomear e Excluir.
            # A implementação real exigiria lógica de manipulação de arquivos no disco.
            menu.addAction("✏️ Renomear Arquivo (Simulado)", lambda: QMessageBox.information(self, "Ação Simulada", f"Renomeando: {file_name} (ação simulada)"))
            menu.addAction("❌ Excluir Arquivo (Simulado)", lambda: QMessageBox.warning(self, "Ação Simulada", f"Excluindo: {file_name} (ação simulada)"))
            
        menu.exec_(self.tree.viewport().mapToGlobal(pos)) # Show menu at mouse position

    def _create_new_excel_file(self):
        """Abre um diálogo para criar um novo arquivo Excel vazio."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Criar Novo Arquivo Excel")
        dialog_layout = QVBoxLayout(dialog)

        file_name_label = QLabel("Nome do novo arquivo (.xlsx):")
        file_name_input = QLineEdit()
        file_name_input.setPlaceholderText("ex: meu_novo_arquivo.xlsx")
        
        dialog_layout.addWidget(file_name_label)
        dialog_layout.addWidget(file_name_input)

        button_box = QHBoxLayout()
        create_btn = QPushButton("Criar")
        cancel_btn = QPushButton("Cancelar")
        
        button_box.addWidget(create_btn)
        button_box.addWidget(cancel_btn)
        
        dialog_layout.addLayout(button_box)

        create_btn.clicked.connect(dialog.accept)
        cancel_btn.clicked.connect(dialog.reject)

        if dialog.exec_() == QDialog.Accepted:
            new_file_name = file_name_input.text().strip()
            if not new_file_name.endswith(".xlsx"):
                new_file_name += ".xlsx"
            
            new_file_path = os.path.join(USER_SHEETS_DIR, new_file_name)
            
            if os.path.exists(new_file_path):
                QMessageBox.warning(self, "Arquivo Já Existe", f"O arquivo '{new_file_name}' já existe. Por favor, escolha outro nome.")
                return

            try:
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = "Sheet1" # Default first sheet
                wb.save(new_file_path)
                QMessageBox.information(self, "Arquivo Criado", f"Arquivo '{new_file_name}' criado com sucesso em '{USER_SHEETS_DIR}'.")
                self._populate_workspace_tree() # Refresh the tree view
            except Exception as e:
                QMessageBox.critical(self, "Erro ao Criar Arquivo", f"Erro ao criar o arquivo Excel: {e}")


    def _show_tab_context_menu(self, pos):
        """Exibe um menu de contexto para as abas no widget de abas."""
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
        """Fecha todas as abas, exceto a do índice 'keep_index'."""
        # Iterate in reverse to avoid index issues when removing tabs
        for i in reversed(range(self.tabs.count())):
            if i != keep_index:
                self.tabs.removeTab(i)

# === PONTO DE ENTRADA ===
if __name__ == "__main__":
    app = QApplication(sys.argv)
    login = LoginWindow()
    login.show()
    sys.exit(app.exec_())
