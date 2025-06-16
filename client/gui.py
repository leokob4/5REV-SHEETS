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
from ui.tools.excel_viewer_tool import ExcelViewerTool # Novo import para a ferramenta genérica de visualização de Excel

# Importa o novo diálogo de adição de item (agora de client.add_item_dialog)
# Note: As funções de adição de item ao 'workspace_data.xlsx' serão removidas,
# mas o diálogo pode ser reutilizado para outras finalidades no futuro se necessário.
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
EXCEL_TOOL_MAP = {
    "output.xlsx": ProductDataTool,
    "bom_data.xlsx": BomManagerTool,
    "configurador_data.xlsx": ConfiguradorTool,
    "colaboradores_data.xlsx": ColaboradoresTool,
    "items_data.xlsx": ItemsTool,
    "manufacturing_data.xlsx": ManufacturingTool,
    "pcp_data.xlsx": PcpTool,
    "estoque_data.xlsx": EstoqueTool,
    "financeiro_data.xlsx": FinanceiroTool,
    "pedidos_data.xlsx": PedidosTool,
    "manutencao_data.xlsx": ManutencaoTool,
    # Atenção: 'db.xlsx' e 'tools.xlsx' são arquivos do sistema, não devem ser abertos diretamente por aqui.
    # Engenharia Workflow Tool (mod4) não interage diretamente com um arquivo .xlsx de usuário de forma padronizada para este mapeamento
}


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
        self.tree.setHeaderLabel("Arquivos do Usuário (.xlsx)")
        self._populate_workspace_tree() # Agora lista arquivos .xlsx em user_sheets
        self.tree.expandAll()
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._show_tree_context_menu)
        self.tree.itemDoubleClicked.connect(self._open_file_from_tree) # Abre o arquivo selecionado

        search_layout = QHBoxLayout()
        self.item_search_bar = QLineEdit()
        self.item_search_bar.setPlaceholderText("Pesquisar arquivos...")
        self.item_search_bar.returnPressed.connect(self.handle_file_search)
        self.search_items_btn = QPushButton("🔍")
        self.search_items_btn.clicked.connect(self.handle_file_search)

        search_layout.addWidget(self.item_search_bar)
        search_layout.addWidget(self.search_items_btn)

        left_pane_layout.addWidget(QLabel("Arquivos do Usuário"))
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
        """Popula a árvore com os arquivos .xlsx encontrados em USER_SHEETS_DIR."""
        self.tree.clear()
        
        # Filtra para incluir apenas arquivos .xlsx e excluir db.xlsx
        excel_files = [f for f in os.listdir(USER_SHEETS_DIR) if f.endswith('.xlsx') and f != 'db.xlsx']
        
        # Opcional: Adicionar uma pasta "Sistema" para db.xlsx e tools.xlsx se quisermos mostrá-los
        # root_system = QTreeWidgetItem(["Arquivos do Sistema"])
        # self.tree.addTopLevelItem(root_system)
        # root_system.addChild(QTreeWidgetItem(["db.xlsx"])) # Placeholder
        # root_system.addChild(QTreeWidgetItem(["tools.xlsx"])) # Placeholder


        if not excel_files:
            QMessageBox.information(self, "Nenhum Arquivo Encontrado", f"Nenhum arquivo .xlsx encontrado na pasta: {USER_SHEETS_DIR}. Crie arquivos para começar!")
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
        tool_class = EXCEL_TOOL_MAP.get(file_name)

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

    def handle_file_search(self):
        """
        Realiza uma busca nos nomes dos arquivos Excel e exibe os resultados.
        """
        search_term = self.item_search_bar.text().strip().lower()
        if not search_term:
            QMessageBox.information(self, "Pesquisar", "Por favor, digite um termo de pesquisa.")
            return

        excel_files = [f for f in os.listdir(USER_SHEETS_DIR) if f.endswith('.xlsx') and f != 'db.xlsx']
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
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)

        dialog.exec_()

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
        """Exibe um menu de contexto para itens (arquivos) na visualização em árvore."""
        item = self.tree.itemAt(pos)
        
        menu = QMenu()

        # Adicionar uma opção para criar um novo arquivo .xlsx vazio, se desejado
        # Exemplo: menu.addAction("➕ Criar Novo Arquivo .xlsx (Simulado)")

        if item: # Se um arquivo foi clicado
            file_path = item.data(0, Qt.UserRole)
            file_name = os.path.basename(file_path)
            
            menu.addSeparator()
            menu.addAction("🔍 Abrir Arquivo", lambda: self._open_file_from_tree(item, 0))
            
            # Opções para o StructureViewTool
            # Se o arquivo contiver uma planilha "structure", oferecer para abrir a estrutura
            try:
                wb = openpyxl.load_workbook(file_path)
                if "structure" in wb.sheetnames:
                    menu.addAction("📊 Visualizar Estrutura", lambda: self._open_tab(f"Estrutura: {file_name}", StructureViewTool(file_path=file_path, sheet_name="structure")))
            except Exception:
                pass # Ignora se o arquivo não puder ser lido ou não tiver a aba

            menu.addAction("✏️ Renomear Arquivo (Simulado)", lambda: QMessageBox.information(self, "Ação Simulada", f"Renomeando: {file_name} (ação simulada)"))
            menu.addAction("❌ Excluir Arquivo (Simulado)", lambda: QMessageBox.warning(self, "Ação Simulada", f"Excluindo: {file_name} (ação simulada)"))
            
        menu.exec_(self.tree.viewport().mapToGlobal(pos))

    # _add_new_root_item e _add_new_subitem são removidos ou alterados drasticamente
    # pois não há mais um 'workspace_data.xlsx' para adicionar items de forma genérica.
    # A adição de dados agora ocorrerá dentro das ferramentas específicas.

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
