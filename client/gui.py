import sys
import os
import bcrypt
import openpyxl
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QAction, QTabWidget, QMenu, QToolButton,
    QWidget, QVBoxLayout, QSplitter, QTreeWidget, QTreeWidgetItem,
    QLabel, QLineEdit, QPushButton, QHBoxLayout, QMessageBox, QGraphicsView, QGraphicsScene, QGraphicsRectItem, QGraphicsLineItem
)
from PyQt5.QtCore import Qt, QPointF
from PyQt5.QtGui import QBrush, QPen, QColor

# --- File Paths Configuration ---
# Define standard paths for consistency.
USER_SHEETS_DIR = "user_sheets"
APP_SHEETS_DIR = "app_sheets"
DB_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "db.xlsx")
TOOLS_EXCEL_PATH = os.path.join(APP_SHEETS_DIR, "tools.xlsx") # New path for tools.xlsx

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
        QMessageBox.critical(None, "File Not Found", f"Database file not found at: {DB_EXCEL_PATH}")
        return {}
    except KeyError:
        QMessageBox.critical(None, "Sheet Error", f"Sheet 'users' not found in {DB_EXCEL_PATH}")
        return {}
    except Exception as e:
        QMessageBox.critical(None, "Loading Error", f"Error loading users: {e}")
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
                raise ValueError("Username already exists.")

        password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
        # Append new user data to the sheet
        sheet.append([next_id, username, password_hash, role])
        wb.save(DB_EXCEL_PATH)
    except FileNotFoundError:
        QMessageBox.critical(None, "File Not Found", f"Database file not found at: {DB_EXCEL_PATH}. Cannot register user.")
    except KeyError:
        QMessageBox.critical(None, "Sheet Error", f"Sheet 'users' not found in {DB_EXCEL_PATH}. Cannot register user.")
    except Exception as e:
        QMessageBox.critical(None, "Registration Error", f"Error registering user: {e}")

def load_tools_from_excel():
    """
    Loads tool data from the dedicated tools Excel file.
    Corrected path to 'app_sheets/tools.xlsx' and added error handling.
    """
    try:
        wb = openpyxl.load_workbook(TOOLS_EXCEL_PATH)
        sheet = wb["tools"] # Corrected to read from 'tools' sheet
        tools = {}
        # Iterate from the second row to skip headers
        for row in sheet.iter_rows(min_row=2):
            # Ensure enough cells are present to avoid IndexError
            if len(row) >= 4:
                tools[row[0].value] = {
                    "id": row[0].value,
                    "name": row[1].value,
                    "description": row[2].value,
                    "path": row[3].value
                }
        return tools
    except FileNotFoundError:
        QMessageBox.critical(None, "File Not Found", f"Tools file not found at: {TOOLS_EXCEL_PATH}. Please ensure it exists.")
        return {}
    except KeyError:
        QMessageBox.critical(None, "Sheet Error", f"Sheet 'tools' not found in {TOOLS_EXCEL_PATH}. Please ensure the sheet name is 'tools'.")
        return {}
    except Exception as e:
        QMessageBox.critical(None, "Loading Error", f"Error loading tools: {e}")
        return {}


def load_role_permissions():
    """Loads role permissions from the database Excel file."""
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
                print(f"Warning: Skipping malformed row in 'access' sheet: {', '.join(str(c.value) for c in row)}")
        return perms
    except FileNotFoundError:
        QMessageBox.critical(None, "File Not Found", f"Database file not found at: {DB_EXCEL_PATH}")
        return {}
    except KeyError:
        QMessageBox.critical(None, "Sheet Error", f"Sheet 'access' not found in {DB_EXCEL_PATH}")
        return {}
    except Exception as e:
        QMessageBox.critical(None, "Loading Error", f"Error loading permissions: {e}")
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
        self.username_input.setPlaceholderText("Username")
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Password")
        self.password_input.setEchoMode(QLineEdit.Password)

        login_btn = QPushButton("Login")
        login_btn.clicked.connect(self.authenticate)

        register_btn = QPushButton("Register")
        register_btn.clicked.connect(self.handle_register)

        layout.addWidget(QLabel("Welcome to 5revolution"))
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
            QMessageBox.warning(self, "Login Failed", "Username and password cannot be empty.")
            return

        user = self.users.get(uname)

        if not user or not bcrypt.checkpw(pwd.encode(), user["password_hash"].encode()):
            QMessageBox.warning(self, "Login Failed", "Invalid username or password.")
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
            QMessageBox.warning(self, "Validation Error", "Username and password are required for registration.")
            return

        try:
            register_user(uname, pwd)
            QMessageBox.information(self, "Registered", f"User '{uname}' registered successfully with role 'user'.")
            self.users = load_users_from_excel() # Reload users after registration
            self.username_input.clear()
            self.password_input.clear()
        except ValueError as ve:
            QMessageBox.warning(self, "Registration Failed", str(ve))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"An error occurred during registration: {e}")

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
        add_node_btn = QPushButton("Add Task Node")
        add_node_btn.clicked.connect(self._add_task_node)
        add_link_btn = QPushButton("Add Dependency Link")
        add_link_btn.clicked.connect(self._add_dependency_link)
        clear_btn = QPushButton("Clear Diagram")
        clear_btn.clicked.connect(self._clear_diagram)

        control_layout.addWidget(add_node_btn)
        control_layout.addWidget(add_link_btn)
        control_layout.addWidget(clear_btn)
        self.layout.addLayout(control_layout)

        self.nodes = [] # To keep track of added nodes

    def _add_sample_diagram_elements(self):
        """Adds some sample elements to the diagram scene."""
        # Task nodes
        node1 = self.scene.addRect(50, 50, 100, 50, QPen(Qt.black), QBrush(Qt.lightblue))
        node2 = self.scene.addRect(200, 150, 100, 50, QPen(Qt.black), QBrush(Qt.lightgreen))
        node3 = self.scene.addRect(350, 50, 100, 50, QPen(Qt.black), QBrush(Qt.lightcoral))

        self.scene.addText("Design Phase", QPointF(55, 65))
        self.scene.addText("Review (Approved)", QPointF(205, 165))
        self.scene.addText("Production Prep", QPointF(355, 65))

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
        self.scene.addText(f"New Task {len(self.nodes) + 1}", QPointF(x + 5, y + 15))
        self.nodes.append(node)
        self.view.centerOn(node)

    def _add_dependency_link(self):
        """Prompts user to select two nodes to link. (Conceptual, requires selection logic)."""
        QMessageBox.information(self, "Add Link", "Click two task nodes to create a link. (Selection logic to be implemented)")
        # In a real implementation, you'd need selection mechanisms (e.g., click listeners on QGraphicsRectItem)
        # to get two nodes and then draw a QGraphicsLineItem between their centroids or edges.

    def _clear_diagram(self):
        """Clears all elements from the diagram."""
        self.scene.clear()
        self.nodes = [] # Reset nodes list
        QMessageBox.information(self, "Diagram Cleared", "The diagram has been cleared.")


# === MAIN GUI ===
class TeamcenterStyleGUI(QMainWindow):
    """
    The main application GUI, styled to resemble Teamcenter.
    Provides a workspace tree view, tabbed content area, and a toolbar.
    """
    def __init__(self, user):
        super().__init__()
        self.setWindowTitle("5revolution Platform")
        self.setGeometry(100, 100, 1280, 800) # x, y, width, height

        self.username = user["username"]
        self.role = user["role"]
        self.tools = load_tools_from_excel() # Load tools using the updated function
        self.permissions = load_role_permissions()

        self._create_toolbar()
        self._create_main_layout()

        # Display user information in status bar
        self.statusBar().showMessage(f"Logged in as: {self.username} | Role: {self.role}")

    def _create_toolbar(self):
        """Creates the main application toolbar."""
        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setMovable(False) # Make toolbar fixed
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)

        # 🛠 Tools Menu Button
        self.tools_btn = QToolButton()
        self.tools_btn.setText("🛠 Tools")
        self.tools_btn.setPopupMode(QToolButton.InstantPopup) # Shows menu instantly on click
        tools_menu = QMenu()

        allowed_tools = self.permissions.get(self.role, []) # Get allowed tools for the user's role
        for tid, tool in self.tools.items():
            # Check if user has permission for this tool or if role is 'all'
            if allowed_tools == "all" or tid in allowed_tools:
                action = tools_menu.addAction(tool["name"])
                # Use functools.partial for passing arguments to slot (cleaner for loops)
                # We need to map tool IDs to actual widget classes or functions
                if tool["id"] == "mod4": # Special handling for the new Engenharia tool
                    action.triggered.connect(lambda chk=False, title=tool["name"]: self._open_tab(title, EngenhariaWorkflowTool()))
                else:
                    action.triggered.connect(lambda chk=False, title=tool["name"], desc=tool["description"]: self._open_tab(title, QLabel(desc)))
        self.tools_btn.setMenu(tools_menu)
        self.toolbar.addWidget(self.tools_btn)

        # 👤 Profile Menu Button
        self.profile_btn = QToolButton()
        self.profile_btn.setText(f"👤 {self.username}") # Display username in profile button
        self.profile_btn.setPopupMode(QToolButton.InstantPopup)
        profile_menu = QMenu()
        profile_menu.addAction("⚙️ Settings", self._open_options)
        profile_menu.addSeparator() # Add a separator for better visual grouping
        profile_menu.addAction("🔒 Logout", self._logout)
        self.profile_btn.setMenu(profile_menu)
        self.toolbar.addWidget(self.profile_btn)

    def _create_main_layout(self):
        """Creates the main split layout with tree view and tabs."""
        self.splitter = QSplitter() # Allows resizing of sub-widgets

        # 🌳 Tree View (Left Pane)
        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Workspace")
        self._populate_sample_tree() # Populate with sample data
        self.tree.expandAll() # Expand all tree items by default
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu) # Enable custom context menu
        self.tree.customContextMenuRequested.connect(self._show_tree_context_menu)

        # 📑 Tabs (Right Pane)
        self.tabs = QTabWidget()
        self.tabs.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tabs.customContextMenuRequested.connect(self._show_tab_context_menu)
        self.tabs.setTabsClosable(True) # Make tabs closable by default
        self.tabs.tabCloseRequested.connect(self.tabs.removeTab) # Connect close button to remove tab

        # Welcome/Home Tab
        welcome_widget = QWidget()
        welcome_layout = QVBoxLayout()
        welcome_layout.addWidget(QLabel(f"Welcome {self.username} – Role: {self.role}"))
        welcome_widget.setLayout(welcome_layout)
        self.tabs.addTab(welcome_widget, "Home")

        # Add widgets to the splitter
        self.splitter.addWidget(self.tree)
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
        root = QTreeWidgetItem(["Projects"])
        project1 = QTreeWidgetItem(["Demo Project - Rev A"])
        project1.addChild(QTreeWidgetItem(["Part-001"]))
        project1.addChild(QTreeWidgetItem(["Assembly-001"]))
        root.addChild(project1)

        project2 = QTreeWidgetItem(["Sample Variant - V1.0"])
        project2.addChild(QTreeWidgetItem(["Component-XYZ"]))
        root.addChild(project2)

        self.tree.addTopLevelItem(root)


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
        QMessageBox.information(self, "Options", "User settings will be managed here. (Feature under development)")

    def _logout(self):
        """Logs out the current user and returns to the login screen."""
        confirm_logout = QMessageBox.question(self, "Logout Confirmation", "Are you sure you want to log out?",
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
        # Actions for root items (e.g., "Projects")
        if item.parent() is None:
            menu.addAction("🔁 Refresh Project", lambda: QMessageBox.information(self, "Mock Action", "Project refreshed (mock action)"))
            menu.addAction("➕ Add New Item", lambda: QMessageBox.information(self, "Mock Action", "Add new item (mock action)"))
        # Actions for child items (e.g., "Demo Project", "Part-001")
        else:
            menu.addAction("🔍 View Details", lambda: QMessageBox.information(self, "Mock Action", f"Viewing details for: {item.text(0)} (mock action)"))
            menu.addAction("✏️ Edit Properties", lambda: QMessageBox.information(self, "Mock Action", f"Editing properties for: {item.text(0)} (mock action)"))
            menu.addAction("❌ Delete Item", lambda: QMessageBox.warning(self, "Mock Action", f"Deleted: {item.text(0)} (mock action)"))

        menu.exec_(self.tree.viewport().mapToGlobal(pos)) # Show menu at mouse position

    def _show_tab_context_menu(self, pos):
        """Displays a context menu for tabs in the tab widget."""
        index = self.tabs.tabBar().tabAt(pos)
        if index < 0: return # No tab clicked

        menu = QMenu()
        menu.addAction("❌ Close Tab", lambda: self.tabs.removeTab(index))
        # Ensure "Close Other Tabs" doesn't close the current tab if it's the only one
        if self.tabs.count() > 1:
            menu.addAction("🔁 Close Other Tabs", lambda: self._close_other_tabs(index))
        if self.tabs.count() > 0: # Only show "Close All Tabs" if there are tabs
            menu.addAction("🧹 Close All Tabs", self.tabs.clear)
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
