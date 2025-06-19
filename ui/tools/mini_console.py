from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QTextEdit, QLineEdit, QLabel, QSizePolicy
from PyQt5.QtCore import Qt, QTimer, pyqtSignal

class MiniConsoleWidget(QWidget):
    """
    Um mini-console para exibir saída e receber entrada, útil para depuração
    ou para interações simples via linha de comando.
    """
    # Sinal emitido quando o usuário pressiona Enter no campo de input
    command_entered = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_ui()

    def _init_ui(self):
        """Inicializa os elementos da interface do usuário do mini-console."""
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0) # Remove margens internas

        # Área de saída do console
        self.output_area = QTextEdit()
        self.output_area.setReadOnly(True)
        self.output_area.setPlaceholderText("Saída do Console...")
        self.output_area.setFontPointSize(9) # Tamanho da fonte menor
        # Define uma política de tamanho para que o QTextEdit possa ser redimensionado
        self.output_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        main_layout.addWidget(self.output_area)

        # Barra de input do console
        input_layout = QHBoxLayout()
        input_layout.setContentsMargins(0, 0, 0, 0)

        self.input_label = QLabel(">>>")
        self.input_area = QLineEdit()
        self.input_area.setPlaceholderText("Digite um comando aqui...")
        self.input_area.returnPressed.connect(self._handle_command_input) # Conecta Enter

        input_layout.addWidget(self.input_label)
        input_layout.addWidget(self.input_area)
        
        main_layout.addLayout(input_layout)
        self.setLayout(main_layout)

        # Inicia com uma mensagem de boas-vindas
        self.append_output("Mini-Console Inicializado.")

    def _handle_command_input(self):
        """Processa o comando digitado pelo usuário."""
        command = self.input_area.text().strip()
        self.input_area.clear() # Limpa o input após o comando

        if command:
            self.append_output(f">>> {command}") # Mostra o comando no output
            self.command_entered.emit(command) # Emite o sinal com o comando para o GUI principal


    def append_output(self, text: str):
        """
        Adiciona texto à área de saída do console.
        """
        self.output_area.append(text)
        self.output_area.verticalScrollBar().setValue(self.output_area.verticalScrollBar().maximum()) # Rola para o final
    
    def clear_output(self):
        """Limpa todo o texto da área de saída do console."""
        self.output_area.clear()

