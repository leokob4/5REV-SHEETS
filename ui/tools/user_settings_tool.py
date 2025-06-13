import sys # Added import sys
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QLineEdit, QHBoxLayout, QApplication

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

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Example usage for standalone testing
    window = UserSettingsTool("testuser", "admin")
    window.show()
    sys.exit(app.exec_())
