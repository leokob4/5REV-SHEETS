import os
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton, QMessageBox, QTreeWidget, QTreeWidgetItem
from PyQt5.QtCore import Qt, QTimer

class SearchBarWidget(QWidget):
    """
    Um widget reutilizável que fornece uma barra de pesquisa e lógica
    para filtrar itens em um QTreeWidget alvo.
    """
    def __init__(self, target_tree_widget: QTreeWidget, parent=None):
        super().__init__(parent)
        self.target_tree_widget = target_tree_widget
        self._init_ui()

    def _init_ui(self):
        """Inicializa os elementos da interface do usuário da barra de pesquisa."""
        main_layout = QVBoxLayout(self)
        search_layout = QHBoxLayout()

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Buscar no Espaço de Trabalho...")
        # Conecta o Enter para executar a busca
        self.search_input.returnPressed.connect(self.execute_search) 
        
        self.search_button = QPushButton("Buscar")
        self.search_button.clicked.connect(self.execute_search)
        
        self.clear_button = QPushButton("Limpar")
        self.clear_button.clicked.connect(self.clear_search)

        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_button)
        search_layout.addWidget(self.clear_button)
        
        main_layout.addLayout(search_layout)
        self.setLayout(main_layout)

    def execute_search(self):
        """
        Executa a busca com base no texto do input e filtra os itens no QTreeWidget alvo.
        Os itens que não correspondem são ocultados.
        """
        search_term = self.search_input.text().strip().lower()
        
        # Limpa qualquer filtro anterior antes de aplicar o novo
        self.clear_search(clear_input=False)

        if not search_term:
            QMessageBox.information(self, "Busca vazia", "Digite algo para buscar.")
            return

        # Assumimos que o primeiro item top-level é a raiz do workspace
        workspace_root = self.target_tree_widget.topLevelItem(0) 
        if not workspace_root or workspace_root.text(0) != "Projetos/Espaço de Trabalho":
            QMessageBox.warning(self, "Erro na Árvore", "Não foi possível encontrar a raiz 'Projetos/Espaço de Trabalho'.")
            return

        found_items = []
        for i in range(workspace_root.childCount()):
            item = workspace_root.child(i)
            item_name = item.text(0).lower()
            item_type = item.text(1).lower() if item.childCount() == 0 else "pasta" 

            if search_term in item_name or search_term in item_type:
                found_items.append(item)
                item.setHidden(False) # Garante que o item correspondente é visível
                item.setSelected(True)
                item.setExpanded(True)
            else:
                item.setHidden(True) # Oculta itens que não correspondem

        if not found_items:
            QMessageBox.information(self, "Sem resultados", f"Nenhum item encontrado para: '{search_term}' no Espaço de Trabalho.")
            return

        # Foca no primeiro resultado
        self.target_tree_widget.setCurrentItem(found_items[0])

    def clear_search(self, clear_input=True):
        """
        Limpa a barra de busca e reexibe todos os itens da raiz do workspace.
        """
        if clear_input:
            self.search_input.clear()
        self.target_tree_widget.clearSelection()

        # Reexibe todos os itens sob a raiz do workspace
        workspace_root = self.target_tree_widget.topLevelItem(0)
        if workspace_root and workspace_root.text(0) == "Projetos/Espaço de Trabalho":
            for i in range(workspace_root.childCount()):
                item = workspace_root.child(i)
                item.setHidden(False) # Reexibe o item
                item.setExpanded(False) # Colapsa-o para uma visualização limpa após limpar
        
        # Opcional: Colapsar a raiz do workspace após limpar (se ela não for um filtro ativo)
        if workspace_root:
            workspace_root.setExpanded(False)

