import sys
import openpyxl
import bcrypt
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout,
    QLineEdit, QPushButton, QListWidget, QMessageBox,
    QDialog, QTextEdit
)

# === Excel DB Loader ===
def load_excel():
    wb_main = openpyxl.load_workbook("app_sheets/main.xlsx")
    refs_sheet = wb_main["refs"]
    refs = {row[1].value: row[0].value for row in refs_sheet.iter_rows(min_row=2)}

    wb_users = openpyxl.load_workbook(f"app_sheets/{refs['users']}")
    users = {}
    for row in wb_users["users"].iter_rows(min_row=2):
        users[row[1].value] = {
            "username": row[1].value,
            "password_hash": row[2].value,
            "role": row[3].value,
        }

    wb_modules = openpyxl.load_workbook(f"app_sheets/{refs['modules']}")
    modules = {}
    for row in wb_modules["modules"].iter_rows(min_row=2):
        modules[row[0].value] = {
            "id": row[0].value,
            "name": row[1].value,
            "description": row[2].value,
        }

    wb_perms = openpyxl.load_workbook(f"app_sheets/{refs['permissions']}")
    perms = {}
    for row in wb_perms["permissions"].iter_rows(min_row=2):
        perms[row[0].value] = (
            row[1].value.split(",") if row[1].value != "all" else "all"
        )

    return users, modules, perms

# === GUI Classes ===
class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🔐 PLM ERP Login")
        self.resize(300, 150)
        self.users, self.modules, self.permissions = load_excel()

        layout = QVBoxLayout()
        self.username_input = QLineEdit(placeholderText="Username")
        self.password_input = QLineEdit(placeholderText="Password")
        self.password_input.setEchoMode(QLineEdit.Password)
        self.login_btn = QPushButton("Login")
        self.login_btn.clicked.connect(self.check_login)

        layout.addWidget(QLabel("🔐 Enter your credentials"))
        layout.addWidget(self.username_input)
        layout.addWidget(self.password_input)
        layout.addWidget(self.login_btn)

        self.setLayout(layout)

    def check_login(self):
        uname = self.username_input.text()
        pwd = self.password_input.text()
        user = self.users.get(uname)
        if not user or not bcrypt.checkpw(pwd.encode(), user["password_hash"].encode()):
            QMessageBox.critical(self, "Login Failed", "Invalid credentials")
            return
        self.close()
        self.dashboard = Dashboard(user, self.modules, self.permissions)
        self.dashboard.show()

class Dashboard(QWidget):
    def __init__(self, user, modules, permissions):
        super().__init__()
        self.setWindowTitle("📦 Dashboard")
        self.resize(500, 400)

        self.user = user
        self.modules = modules
        self.permissions = permissions

        layout = QVBoxLayout()
        layout.addWidget(QLabel(f"👤 {user['username']} ({user['role']})"))

        self.mod_list = QListWidget()
        allowed = permissions[user["role"]]
        shown = modules.values() if allowed == "all" else [modules[m] for m in allowed if m in modules]
        self.shown_modules = list(shown)
        for mod in shown:
            self.mod_list.addItem(mod["name"])

        self.mod_list.itemClicked.connect(self.show_module_popup)

        layout.addWidget(QLabel("🧰 Available Modules"))
        layout.addWidget(self.mod_list)

        if user["role"] == "admin":
            reload_btn = QPushButton("🔁 Reload Sheets")
            reload_btn.clicked.connect(self.reload_sheets)
            layout.addWidget(reload_btn)

        self.setLayout(layout)

    def show_module_popup(self, item):
        mod_name = item.text()
        mod = next((m for m in self.shown_modules if m["name"] == mod_name), None)
        if mod:
            dlg = QDialog(self)
            dlg.setWindowTitle(mod["name"])
            dlg.resize(400, 300)
            layout = QVBoxLayout()
            text = QTextEdit()
            text.setText(f"📦 {mod['name']}\n\n📝 {mod['description']}")
            text.setReadOnly(True)
            layout.addWidget(text)
            btn = QPushButton("Close")
            btn.clicked.connect(dlg.close)
            layout.addWidget(btn)
            dlg.setLayout(layout)
            dlg.exec_()

    def reload_sheets(self):
        self.users, self.modules, self.permissions = load_excel()
        QMessageBox.information(self, "Reloaded", "Excel files reloaded. Restart required for updates.")

# === Launcher ===
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = LoginWindow()
    win.show()
    sys.exit(app.exec_())
