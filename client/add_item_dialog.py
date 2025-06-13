import sys
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QComboBox, QMessageBox
from PyQt5.QtCore import Qt

class AddItemDialog(QDialog):
    """
    A dialog for adding new items to the workspace.
    Allows specifying ID, Name, Type, ParentID, and Description.
    """
    def __init__(self, parent_id=None, parent_name="Nenhum", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Adicionar Novo Item ao Espaço de Trabalho")
        self.setModal(True) # Make it a modal dialog
        self.item_data = None # To store collected data

        self.layout = QVBoxLayout(self)

        # Item ID
        id_layout = QHBoxLayout()
        id_layout.addWidget(QLabel("ID do Item:"))
        self.id_input = QLineEdit()
        self.id_input.setPlaceholderText("Ex: PROJ-002, PART-005")
        # Restrição para IDs: Apenas letras maiúsculas, números e hífen, max 15 caracteres.
        self.id_input.setMaxLength(15)
        self.id_input.textEdited.connect(self._validate_id_input) # Validate on text change
        id_layout.addWidget(self.id_input)
        self.layout.addLayout(id_layout)

        # Item Name
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("Nome do Item:"))
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Nome Descritivo do Item")
        name_layout.addWidget(self.name_input)
        self.layout.addLayout(name_layout)

        # Item Type
        type_layout = QHBoxLayout()
        type_layout.addWidget(QLabel("Tipo:"))
        self.type_selector = QComboBox()
        self.type_selector.addItems(["Project", "Assembly", "Part", "Component", "Document", "Variant", "Other"])
        type_layout.addWidget(self.type_selector)
        self.layout.addLayout(type_layout)

        # Parent ID (Optional, pre-filled if adding as subitem)
        parent_layout = QHBoxLayout()
        parent_layout.addWidget(QLabel("ID do Item Pai:"))
        self.parent_id_input = QLineEdit(parent_id if parent_id else "")
        self.parent_id_input.setPlaceholderText(f"Herdado de '{parent_name}' ou 'ROOT'")
        # If a parent is given, make it read-only
        if parent_id:
            self.parent_id_input.setReadOnly(True)
        parent_layout.addWidget(self.parent_id_input)
        self.layout.addLayout(parent_layout)

        # Description
        desc_layout = QHBoxLayout()
        desc_layout.addWidget(QLabel("Descrição:"))
        self.description_input = QLineEdit()
        self.description_input.setPlaceholderText("Breve descrição do item")
        desc_layout.addWidget(self.description_input)
        self.layout.addLayout(desc_layout)

        # Buttons
        button_layout = QHBoxLayout()
        self.add_btn = QPushButton("Adicionar")
        self.add_btn.clicked.connect(self.accept_data)
        self.cancel_btn = QPushButton("Cancelar")
        self.cancel_btn.clicked.connect(self.reject)
        button_layout.addWidget(self.add_btn)
        button_layout.addWidget(self.cancel_btn)
        self.layout.addLayout(button_layout)

    def _validate_id_input(self, text):
        """Ensures ID input contains only uppercase letters, numbers, and hyphens."""
        cleaned_text = "".join(c for c in text if c.isalnum() or c == '-').upper()
        if text != cleaned_text:
            self.id_input.setText(cleaned_text) # Update text to filtered version
            QMessageBox.warning(self, "Entrada Inválida", "O ID do Item só pode conter letras maiúsculas (A-Z), números (0-9) e hífens (-). Caracteres inválidos foram removidos.")


    def accept_data(self):
        """Collects data and accepts the dialog."""
        item_id = self.id_input.text().strip().upper() # IDs usually uppercase
        item_name = self.name_input.text().strip()
        item_type = self.type_selector.currentText()
        parent_id = self.parent_id_input.text().strip().upper() if self.parent_id_input.text().strip() else "ROOT"
        description = self.description_input.text().strip()

        if not item_id or not item_name:
            QMessageBox.warning(self, "Entrada Inválida", "ID e Nome do Item são campos obrigatórios.")
            return

        self.item_data = {
            "ID": item_id,
            "Name": item_name,
            "Type": item_type,
            "ParentID": parent_id,
            "Description": description
        }
        self.accept() # Close dialog with QDialog.Accepted

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Exemplo de uso para teste autônomo
    dialog = AddItemDialog(parent_id="PROJ-001", parent_name="Demo Project")
    if dialog.exec_() == QDialog.Accepted:
        print("Item adicionado:", dialog.item_data)
    else:
        print("Adição de item cancelada.")
    sys.exit(app.exec_())
