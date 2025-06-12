import sys
import os
import bcrypt
import openpyxl

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QToolBar, QAction, QTabWidget, QMenu, QToolButton,
    QWidget, QVBoxLayout, QSplitter, QTreeWidget, QTreeWidgetItem,
    QLabel, QLineEdit, QPushButton, QHBoxLayout, QMessageBox
)
from PyQt5.QtCore import Qt

# 🔁 Import db builder from app_sheets
import client.build_db as build_db

# === SHEET HELPERS ===
def load_users_from_excel():
    wb = openpyxl.load_workbook("user_sheets/db.xlsx")
    users_sheet = wb["users"]
    users = {}
    for row in users_sheet.iter_rows(min_row=2):
        users[row[1].value] = {
            "id": row[0].value,
            "username": row[1].value,
            "password_hash": row[2].value,
            "role": row[3].value
        }
    return users

def register_user(username, password, role="user"):
    wb = openpyxl.load_workbook("app_sheets/users.xlsx")
    sheet = wb.active
    next_id = sheet.max_row + 1
    password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
    sheet.append([next_id, username, password_hash, role])
    wb.save("app_sheets/users.xlsx")
    build_db.build_combined_db()  # Refresh db.xlsx immediately

def load_tools_from_excel():
    wb = openpyxl.load_workbook("user_sheets/db.xlsx")
    sheet = wb["tools"]
    tools = {}
    for row in sheet.iter_rows(min_row=2):
        tools[row[0].value] = {
            "id": row[0].value,
            "name": row[1].value,
            "description": row[2].value,
            "path": row[3].value
        }
    return tools

def load_role_permissions():
    wb = openpyxl.load_workbook("user_sheets/db.xlsx")
    sheet = wb["access"]
    perms = {}
    for row in sheet.iter_rows(min_row=2):
        perms[row[0].value] = row[1].value.split(",") if row[1].value != "all" else "all"
    return perms


# === LOGIN WINDOW ===
class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("5revolution Login")
        self.setGeometry(400, 200, 300, 180)

        build_db.build_combined_db()
        self.users = load_users_from_excel()

        layout = QVBoxLayout()
        self.username = QLineEdit()
        self.username.setPlaceholderText("Username")
        self.password = QLineEdit()
        self.password.setPlaceholderText("Password")
        self.password.setEchoMode(QLineEdit.Password)

        login_btn = QPushButton("Login")
        login_btn.clicked.connect(self.authenticate)

        register_btn = QPushButton("Register")
        register_btn.clicked.connect(self.handle_register)

        layout.addWidget(QLabel("Welcome to 5revolution"))
        layout.addWidget(self.username)
        layout.addWidget(self.password)

        btns = QHBoxLayout()
        btns.addWidget(login_btn)
        btns.addWidget(register_btn)

        layout.addLayout(btns)
        self.setLayout(layout)

    def authenticate(self):
        uname = self.username.text()
        pwd = self.password.text()
        user = self.users.get(uname)

        if not user or not bcrypt.checkpw(pwd.encode(), user["password_hash"].encode()):
            QMessageBox.warning(self, "Login Failed", "Invalid username or password.")
            return

        self.main = TeamcenterStyleGUI(user)
        self.main.show()
        self.close()

    def handle_register(self):
        uname = self.username.text().strip()
        pwd = self.password.text().strip()
        if not uname or not pwd:
            QMessageBox.warning(self, "Validation Error", "Username and password required.")
            return

        try:
            register_user(uname, pwd)
            QMessageBox.information(self, "Registered", f"User '{uname}' registered.")
            self.users = load_users_from_excel()
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))


# === MAIN GUI ===
class TeamcenterStyleGUI(QMainWindow):
    def __init__(self, user):
        super().__init__()
        self.setWindowTitle("5revolution Platform")
        self.setGeometry(100, 100, 1280, 800)

        self.username = user["username"]
        self.role = user["role"]
        self.tools = load_tools_from_excel()
        self.permissions = load_role_permissions()

        self._create_toolbar()
        self._create_main_layout()

    def _create_toolbar(self):
        self.toolbar = QToolBar("Main Toolbar")
        self.toolbar.setMovable(False)
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)

        self.tools_btn = QToolButton()
        self.tools_btn.setText("🛠 Tools")
        self.tools_btn.setPopupMode(QToolButton.InstantPopup)
        tools_menu = QMenu()

        allowed = self.permissions.get(self.role, [])
        for tid, tool in self.tools.items():
            if allowed == "all" or tid in allowed:
                tools_menu.addAction(tool["name"], lambda chk=False, t=tool: self._load_tool(t))

        self.tools_btn.setMenu(tools_menu)
        self.toolbar.addWidget(self.tools_btn)

        self.profile_btn = QToolButton()
        self.profile_btn.setText("👤 Profile")
        self.profile_btn.setPopupMode(QToolButton.InstantPopup)
        profile_menu = QMenu()
        profile_menu.addAction("⚙️ Settings", self._open_options)
        profile_menu.addAction("🔒 Logout", self._logout)
        self.profile_btn.setMenu(profile_menu)
        self.toolbar.addWidget(self.profile_btn)

    def _create_main_layout(self):
        self.splitter = QSplitter()

        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("Workspace")
        root = QTreeWidgetItem(["Projects"])
        root.addChild(QTreeWidgetItem(["Demo Project"]))
        root.addChild(QTreeWidgetItem(["Sample Variant"]))
        self.tree.addTopLevelItem(root)
        self.tree.expandAll()
        self.tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self._show_tree_context_menu)

        self.tabs = QTabWidget()
        self.tabs.setContextMenuPolicy(Qt.CustomContextMenu)
        self.tabs.customContextMenuRequested.connect(self._show_tab_context_menu)

        welcome = QWidget()
        wl = QVBoxLayout()
        wl.addWidget(QLabel(f"Welcome {self.username} – Role: {self.role}"))
        welcome.setLayout(wl)
        self.tabs.addTab(welcome, "Home")

        self.splitter.addWidget(self.tree)
        self.splitter.addWidget(self.tabs)
        self.splitter.setStretchFactor(1, 4)

        container = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(self.splitter)
        container.setLayout(layout)
        self.setCentralWidget(container)

    def _load_tool(self, tool):
        try:
            path = tool["path"]
            if not os.path.exists(path):
                QMessageBox.warning(self, "Missing", f"Path not found: {path}")
                return

            import importlib.util
            spec = importlib.util.spec_from_file_location("ToolModule", path)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)

            widget = mod.SheetEditorWidget() if hasattr(mod, "SheetEditorWidget") else QLabel(tool["description"])
            self._open_tab(tool["name"], widget)

        except Exception as e:
            QMessageBox.critical(self, "Tool Error", str(e))

    def _open_tab(self, title, widget):
        for i in range(self.tabs.count()):
            if self.tabs.tabText(i) == title:
                self.tabs.setCurrentIndex(i)
                return
        self.tabs.addTab(widget, title)
        self.tabs.setCurrentIndex(self.tabs.count() - 1)

    def _open_options(self):
        QMessageBox.information(self, "Options", "User settings placeholder.")

    def _logout(self):
        self.close()
        self.login = LoginWindow()
        self.login.show()

    def _show_tree_context_menu(self, pos):
        item = self.tree.itemAt(pos)
        if not item: return
        menu = QMenu()
        menu.addAction("🔁 Refresh", lambda: QMessageBox.information(self, "Mock", "Refreshed."))
        menu.addAction("➕ Add Item", lambda: self._open_tab("New Tool", QLabel("New Item")))
        menu.exec_(self.tree.viewport().mapToGlobal(pos))

    def _show_tab_context_menu(self, pos):
        idx = self.tabs.tabBar().tabAt(pos)
        if idx < 0: return
        menu = QMenu()
        menu.addAction("❌ Close Tab", lambda: self.tabs.removeTab(idx))
        menu.addAction("🧹 Close All", self.tabs.clear)
        menu.exec_(self.tabs.tabBar().mapToGlobal(pos))


# === LAUNCH ===
if __name__ == "__main__":
    app = QApplication(sys.argv)
    login = LoginWindow()
    login.show()
    sys.exit(app.exec_())
