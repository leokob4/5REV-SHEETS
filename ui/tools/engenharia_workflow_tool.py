import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QHBoxLayout, QMessageBox, QGraphicsView, QGraphicsScene
from PyQt5.QtCore import Qt, QPointF
from PyQt5.QtGui import QBrush, QPen, QColor, QFont # Import QFont

class EngenhariaWorkflowTool(QWidget):
    """
    Um widget placeholder para a ferramenta de Diagrama de Fluxo de Trabalho de Engenharia.
    Fornece uma QGraphicsView básica para diagramação.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Engenharia (Fluxo de Trabalho)")
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        self.layout.addWidget(self.view)

        self._add_sample_diagram_elements()

        # Botões de controle
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

        self.nodes = [] # Para rastrear os nós adicionados

    def _add_sample_diagram_elements(self):
        """Adiciona alguns elementos de exemplo à cena do diagrama."""
        # Nós de tarefa
        node1 = self.scene.addRect(50, 50, 100, 50, QPen(Qt.black), QBrush(QColor("lightblue")))
        node2 = self.scene.addRect(200, 150, 100, 50, QPen(Qt.black), QBrush(QColor("lightgreen")))
        node3 = self.scene.addRect(350, 50, 100, 50, QPen(Qt.black), QBrush(QColor("lightcoral")))

        # Uso corrigido de addText: addText retorna um QGraphicsTextItem, depois define sua posição
        text_item1 = self.scene.addText("Fase de Design")
        text_item1.setPos(55, 65)
        
        text_item2 = self.scene.addText("Revisão (Aprovado)")
        text_item2.setPos(205, 165)
        
        text_item3 = self.scene.addText("Preparação da Produção")
        text_item3.setPos(355, 65)

        # Ligações/Setas
        pen = QPen(Qt.darkGray)
        pen.setWidth(2)
        self.scene.addLine(node1.x() + node1.rect().width(), node1.y() + node1.rect().height() / 2,
                           node2.x(), node2.y() + node2.rect().height() / 2, pen)
        self.scene.addLine(node2.x() + node2.rect().width(), node2.y() + node2.rect().height() / 2,
                           node3.x(), node3.y() + node3.rect().height() / 2, pen)

    def _add_task_node(self):
        """Adiciona um novo nó de tarefa genérico ao diagrama."""
        x = 10 + len(self.nodes) * 120 # Deslocamento para novos nós
        y = 10 + (len(self.nodes) % 3) * 70
        node = self.scene.addRect(x, y, 100, 50, QPen(Qt.black), QBrush(QColor("#FFD700"))) # Cor ouro
        
        text_item = self.scene.addText(f"Nova Tarefa {len(self.nodes) + 1}")
        text_item.setPos(x + 5, y + 15)
        
        self.nodes.append(node)
        self.view.centerOn(node)

    def _add_dependency_link(self):
        """Solicita ao usuário que selecione dois nós para ligar. (Conceitual, requer lógica de seleção)."""
        QMessageBox.information(self, "Adicionar Ligação", "Clique em dois nós de tarefa para criar uma ligação. (Lógica de seleção a ser implementada)")

    def _clear_diagram(self):
        """Limpa todos os elementos do diagrama."""
        self.scene.clear()
        self.nodes = [] # Reinicia a lista de nós
        QMessageBox.information(self, "Diagrama Limpo", "O diagrama foi limpo.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = EngenhariaWorkflowTool()
    window.show()
    sys.exit(app.exec_())
