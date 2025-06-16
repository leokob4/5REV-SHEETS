import sys
import os
import openpyxl
import json # Para serializar/desserializar a lista de conexões
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QHBoxLayout, QMessageBox, QGraphicsView, QGraphicsScene, QComboBox, QLabel, QInputDialog
from PyQt5.QtCore import Qt, QPointF
from PyQt5.QtGui import QBrush, QPen, QColor, QFont, QGraphicsRectItem, QGraphicsLineItem, QGraphicsTextItem

# Definindo caminhos de forma dinâmica a partir da localização do script
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.dirname(os.path.dirname(current_dir)) # Navega de ui/tools para a raiz do projeto
user_sheets_dir = os.path.join(project_root, 'user_sheets')

DEFAULT_DATA_EXCEL_FILENAME = "engenharia.xlsx"
DEFAULT_SHEET_NAME = "Workflows" # Nome da planilha padrão para salvar/carregar workflows

class EngenhariaWorkflowTool(QWidget):
    """
    GUI para criar, visualizar, salvar e carregar diagramas de fluxo de trabalho.
    Permite adicionar nós de tarefa e ligações de dependência.
    Os dados do diagrama (posições, textos, etc.) são salvos/carregados de engenharia.xlsx.
    A estrutura da planilha é implícita pelos dados salvos.
    """
    def __init__(self, file_path=None, sheet_name=None):
        super().__init__()
        self.setWindowTitle("Engenharia (Fluxo de Trabalho)")
        
        self.file_path = file_path if file_path else os.path.join(
            user_sheets_dir, DEFAULT_DATA_EXCEL_FILENAME
        )
        self.sheet_name = sheet_name if sheet_name else DEFAULT_SHEET_NAME

        self.layout = QVBoxLayout(self)

        # Controles de arquivo e planilha
        file_sheet_layout = QHBoxLayout()
        self.file_name_label = QLabel(f"<b>Arquivo:</b> {os.path.basename(self.file_path)}")
        file_sheet_layout.addWidget(self.file_name_label)
        file_sheet_layout.addStretch()

        file_sheet_layout.addWidget(QLabel("Planilha:"))
        self.sheet_selector = QComboBox()
        self.sheet_selector.setMinimumWidth(150)
        self.sheet_selector.currentIndexChanged.connect(self._load_workflow_from_selected_sheet)
        file_sheet_layout.addWidget(self.sheet_selector)

        self.refresh_sheets_btn = QPushButton("Atualizar Abas")
        self.refresh_sheets_btn.clicked.connect(self._populate_sheet_selector)
        file_sheet_layout.addWidget(self.refresh_sheets_btn)
        self.layout.addLayout(file_sheet_layout)

        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        self.layout.addWidget(self.view)

        # Botões de controle de diagrama
        control_layout = QHBoxLayout()
        add_node_btn = QPushButton("Adicionar Nó de Tarefa")
        add_node_btn.clicked.connect(self._add_task_node)
        add_link_btn = QPushButton("Adicionar Ligação de Dependência")
        add_link_btn.clicked.connect(self._add_dependency_link)
        clear_btn = QPushButton("Limpar Diagrama")
        clear_btn.clicked.connect(self._clear_diagram)
        save_btn = QPushButton("Salvar Workflow")
        save_btn.clicked.connect(self._save_workflow_to_excel)
        load_btn = QPushButton("Recarregar Workflow") 
        load_btn.clicked.connect(self._load_workflow_from_selected_sheet)

        control_layout.addWidget(add_node_btn)
        control_layout.addWidget(add_link_btn)
        control_layout.addWidget(clear_btn)
        control_layout.addWidget(save_btn)
        control_layout.addWidget(load_btn)
        self.layout.addLayout(control_layout)

        self.nodes = [] # Para rastrear os nós adicionados (QGraphicsRectItem)
        self.node_properties = {} # Para armazenar propriedades adicionais dos nós (texto, ID, etc.)
        self.links = [] # Para rastrear as ligações (QGraphicsLineItem)
        self.next_node_id = 1 # Contador para IDs de nós

        # Popula o seletor de planilhas e carrega os dados iniciais
        self._populate_sheet_selector()

    # Removido _get_workflow_schema_headers() pois os headers serão dinâmicos da planilha.

    def _populate_sheet_selector(self):
        """Popula o QComboBox com os nomes das planilhas do arquivo Excel."""
        self.sheet_selector.clear()
        os.makedirs(os.path.dirname(self.file_path), exist_ok=True)

        # Se o arquivo não existe, adiciona apenas a sheet padrão e configura um estado inicial
        if not os.path.exists(self.file_path):
            QMessageBox.warning(self, "Arquivo Não Encontrado", 
                                f"O arquivo de dados não foi encontrado: {os.path.basename(self.file_path)}. "
                                f"Ele será criado com a aba padrão '{self.sheet_name}' ao salvar.")
            self.sheet_selector.addItem(self.sheet_name)
            self.sheet_selector.setCurrentText(self.sheet_name) # Define o texto atual
            self._load_workflow_from_selected_sheet() # Chama para inicializar a tabela mesmo sem arquivo
            return

        try:
            wb = openpyxl.load_workbook(self.file_path, read_only=True)
            sheet_names = wb.sheetnames
            
            if not sheet_names:
                self.sheet_selector.addItem(self.sheet_name)
                QMessageBox.warning(self, "Nenhuma Planilha Encontrada", 
                                    f"Nenhuma planilha encontrada em '{os.path.basename(self.file_path)}'. "
                                    f"Adicionando a aba padrão '{self.sheet_name}'.")
            else:
                for sheet_name in sheet_names:
                    self.sheet_selector.addItem(sheet_name)
                
                # Tenta selecionar a aba que a ferramenta estava aberta
                default_index = self.sheet_selector.findText(self.sheet_name)
                if default_index != -1:
                    self.sheet_selector.setCurrentIndex(default_index)
                elif sheet_names:
                    self.sheet_selector.setCurrentIndex(0) # Seleciona a primeira sheet disponível
                
            self._load_workflow_from_selected_sheet() # Carrega os dados da aba selecionada

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Listar Planilhas", f"Erro ao listar planilhas em '{os.path.basename(self.file_path)}': {e}")
            self.sheet_selector.addItem(self.sheet_name) # Fallback
            self._clear_diagram() # Limpa o diagrama em caso de erro grave


    def _save_workflow_to_excel(self):
        """
        Salva o estado atual do diagrama para a planilha Excel selecionada.
        Cada nó e link é salvo como uma linha.
        """
        current_sheet_name = self.sheet_selector.currentText()
        if not current_sheet_name:
            QMessageBox.warning(self, "Erro", "Selecione uma planilha para salvar.")
            return

        try:
            wb = None
            if not os.path.exists(self.file_path):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = current_sheet_name
            else:
                wb = openpyxl.load_workbook(self.file_path)
                if current_sheet_name not in wb.sheetnames:
                    ws = wb.create_sheet(current_sheet_name)
                else:
                    ws = wb[current_sheet_name]
            
            # Limpa todas as linhas existentes, exceto a primeira (cabeçalhos)
            # ou todas as linhas se não houver cabeçalhos ainda.
            for row_idx in range(ws.max_row, 0, -1): # Começa do fim, apaga tudo
                ws.delete_rows(row_idx)

            # Cabeçalhos fixos para o formato de salvamento do workflow
            # Estes são internos à ferramenta e não vêm de db.xlsx
            workflow_headers = ["Tipo", "ID", "X", "Y", "Largura", "Altura", "Texto", "Cor", "Conexões"]
            ws.append(workflow_headers) 

            # Salvar Nós
            for node_item in self.nodes:
                node_props = self.node_properties.get(node_item, {})
                node_id = node_props.get("id")
                node_text = node_props.get("text", "") # Pega o texto armazenado
                
                node_x = node_item.rect().x()
                node_y = node_item.rect().y()
                node_width = node_item.rect().width()
                node_height = node_item.rect().height()
                node_color = node_item.brush().color().name() 
                
                connections = [] # Nós não têm "conexões" diretas armazenadas aqui, mas podemos usar para atributos futuros
                
                row_data = [
                    "Node",
                    node_id,
                    node_x,
                    node_y,
                    node_width,
                    node_height,
                    node_text,
                    node_color,
                    json.dumps(connections)
                ]
                ws.append(row_data)

            # Salvar Links
            for link_item in self.links:
                link_connections = {"source": getattr(link_item, 'source_node_id', "N/A"), 
                                    "target": getattr(link_item, 'target_node_id', "N/A")}
                
                row_data = [
                    "Link",
                    "", # Links não têm ID próprio neste esquema simplificado
                    "", "", "", "", "", "", # Campos vazios para links
                    json.dumps(link_connections)
                ]
                ws.append(row_data)

            wb.save(self.file_path)
            QMessageBox.information(self, "Sucesso", f"Workflow salvo em '{current_sheet_name}' em '{os.path.basename(self.file_path)}'.")
        except Exception as e:
            QMessageBox.critical(self, "Erro ao Salvar", f"Não foi possível salvar o workflow: {e}")

    def _load_workflow_from_selected_sheet(self):
        """
        Carrega um diagrama de fluxo de trabalho da planilha Excel selecionada.
        """
        self._clear_diagram() # Limpa o diagrama antes de carregar
        current_sheet_name = self.sheet_selector.currentText()
        
        if not current_sheet_name or not os.path.exists(self.file_path):
            self._add_sample_diagram_elements_if_empty()
            return

        try:
            wb = openpyxl.load_workbook(self.file_path)
            if current_sheet_name not in wb.sheetnames:
                QMessageBox.warning(self, "Planilha Não Encontrada", f"A planilha '{current_sheet_name}' não foi encontrada em '{os.path.basename(self.file_path)}'.")
                self._add_sample_diagram_elements_if_empty() # Adiciona exemplos se a sheet não existir
                return

            sheet = wb[current_sheet_name]
            
            # Cabeçalhos são lidos da primeira linha da planilha, mas para o workflow,
            # esperamos um formato específico. Se a primeira linha não se parece com os cabeçalhos esperados,
            # ou está vazia, podemos considerar a planilha como "sem dados de workflow".
            headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
            # Não há necessidade de mapeamento complexo, apenas garante que os índices esperados existam
            # Se a planilha é recém-criada ou vazia, headers estará vazio.

            # Mapa para acesso fácil às colunas por nome
            header_col_map = {h: idx for idx, h in enumerate(headers)}

            loaded_nodes = {} # Mapeia IDs de nós para os objetos QGraphicsRectItem
            max_id = 0

            for row_idx in range(2, sheet.max_row + 1): # Começa da segunda linha para pular cabeçalhos
                row_values = [cell.value for cell in sheet[row_idx]]
                
                # Acessa os valores por índice do mapa
                def get_val(header_name):
                    col_idx = header_col_map.get(header_name)
                    return row_values[col_idx] if col_idx is not None and col_idx < len(row_values) else None

                row_type = get_val("Tipo")

                if row_type == "Node":
                    node_id = get_val("ID")
                    x = get_val("X") or 0
                    y = get_val("Y") or 0
                    width = get_val("Largura") or 100
                    height = get_val("Altura") or 50
                    text = str(get_val("Texto") or "")
                    color_name = get_val("Cor") or "lightblue"

                    node_rect = self.scene.addRect(x, y, width, height, QPen(Qt.black), QBrush(QColor(color_name)))
                    node_text_item = self.scene.addText(text) 
                    node_text_item.setPos(x + 5, y + 15) # Posição do texto dentro do nó
                    
                    self.nodes.append(node_rect)
                    self.node_properties[node_rect] = {"id": node_id, "text": text, "text_item": node_text_item}
                    loaded_nodes[node_id] = node_rect
                    
                    try:
                        if isinstance(node_id, str) and node_id.startswith("node_"):
                            num_part = int(node_id.split('_')[1])
                            max_id = max(max_id, num_part)
                    except ValueError:
                        pass # Ignora IDs inválidos

                elif row_type == "Link":
                    link_data_str = get_val("Conexões")
                    try:
                        link_data = json.loads(link_data_str)
                        source_id = link_data.get("source")
                        target_id = link_data.get("target")

                        source_node = loaded_nodes.get(source_id)
                        target_node = loaded_nodes.get(target_id)

                        if source_node and target_node:
                            pen = QPen(Qt.darkGray, 2)
                            line = self.scene.addLine(
                                source_node.rect().x() + source_node.rect().width(), source_node.rect().y() + source_node.rect().height() / 2,
                                target_node.rect().x(), target_node.rect().y() + target_node.rect().height() / 2,
                                pen
                            )
                            # Armazena os IDs dos nós conectados para referência futura (se necessário para edição de links)
                            line.source_node_id = source_id
                            line.target_node_id = target_id
                            self.links.append(line)
                    except json.JSONDecodeError:
                        print(f"Aviso: Dados de conexão inválidos para link: {link_data_str}")

            self.next_node_id = max_id + 1 if max_id > 0 else 1 # Atualiza o next_node_id

            if not self.nodes and not self.links: # Se nada foi carregado, adiciona exemplos
                self._add_sample_diagram_elements_if_empty()
            else:
                QMessageBox.information(self, "Sucesso", f"Workflow carregado de '{current_sheet_name}' em '{os.path.basename(self.file_path)}'.")


        except Exception as e:
            QMessageBox.critical(self, "Erro ao Carregar", f"Não foi possível carregar o workflow da aba '{current_sheet_name}': {e}")
            self._clear_diagram() # Limpa em caso de erro no carregamento
            self._add_sample_diagram_elements_if_empty() # E adiciona exemplos

    def _add_sample_diagram_elements_if_empty(self):
        """Adiciona alguns elementos de exemplo à cena do diagrama SOMENTE se ela estiver vazia."""
        if not self.nodes and not self.links:
            # Garante que o next_node_id começa em 1 ao adicionar amostras
            self.next_node_id = 1

            # Nó 1
            node1_rect = self.scene.addRect(50, 50, 100, 50, QPen(Qt.black), QBrush(QColor("lightblue")))
            node1_text = self.scene.addText("Fase de Design")
            node1_text.setPos(55, 65)
            node1_id = f"node_{self.next_node_id}"
            self.nodes.append(node1_rect)
            self.node_properties[node1_rect] = {"id": node1_id, "text": "Fase de Design", "text_item": node1_text}
            self.next_node_id += 1

            # Nó 2
            node2_rect = self.scene.addRect(200, 150, 100, 50, QPen(Qt.black), QBrush(QColor("lightgreen")))
            node2_text = self.scene.addText("Revisão (Aprovado)")
            node2_text.setPos(205, 165)
            node2_id = f"node_{self.next_node_id}"
            self.nodes.append(node2_rect)
            self.node_properties[node2_rect] = {"id": node2_id, "text": "Revisão (Aprovado)", "text_item": node2_text}
            self.next_node_id += 1
            
            # Nó 3
            node3_rect = self.scene.addRect(350, 50, 100, 50, QPen(Qt.black), QBrush(QColor("lightcoral")))
            node3_text = self.scene.addText("Preparação da Produção")
            node3_text.setPos(355, 65)
            node3_id = f"node_{self.next_node_id}"
            self.nodes.append(node3_rect)
            self.node_properties[node3_rect] = {"id": node3_id, "text": "Preparação da Produção", "text_item": node3_text}
            self.next_node_id += 1

            # Ligações/Setas
            pen = QPen(Qt.darkGray)
            pen.setWidth(2)
            
            link1 = self.scene.addLine(node1_rect.x() + node1_rect.rect().width(), node1_rect.y() + node1_rect.rect().height() / 2,
                                       node2_rect.x(), node2_rect.y() + node2_rect.rect().height() / 2, pen)
            link1.source_node_id = node1_id # Adiciona IDs aos links para salvar
            link1.target_node_id = node2_id
            self.links.append(link1)

            link2 = self.scene.addLine(node2_rect.x() + node2_rect.rect().width(), node2_rect.y() + node2_rect.rect().height() / 2,
                                       node3_rect.x(), node3_rect.y() + node3_rect.rect().height() / 2, pen)
            link2.source_node_id = node2_id
            link2.target_node_id = node3_id
            self.links.append(link2)


    def _add_task_node(self):
        """Adiciona um novo nó de tarefa genérico ao diagrama."""
        node_text, ok = QInputDialog.getText(self, "Novo Nó de Tarefa", "Nome da Tarefa:")
        if not ok or not node_text:
            return

        x = 10 + (len(self.nodes) % 5) * 150 # Deslocamento horizontal
        y = 10 + (len(self.nodes) // 5) * 80  # Deslocamento vertical

        node_rect = self.scene.addRect(x, y, 100, 50, QPen(Qt.black), QBrush(QColor("#FFD700"))) # Cor ouro
        
        node_text_item = self.scene.addText(node_text)
        node_text_item.setPos(x + 5, y + 15) # Posição do texto dentro do nó
        
        new_node_id = f"node_{self.next_node_id}"
        self.nodes.append(node_rect)
        self.node_properties[node_rect] = {"id": new_node_id, "text": node_text, "text_item": node_text_item}
        self.next_node_id += 1

        self.view.centerOn(node_rect)
        QMessageBox.information(self, "Nó Adicionado", f"Nó '{node_text}' adicionado com ID: {new_node_id}.")


    def _add_dependency_link(self):
        """
        Lógica para adicionar uma ligação entre dois nós selecionados.
        Este é um exemplo simplificado, a seleção visual exigiria manipulação de eventos do mouse.
        Por ora, usará caixas de diálogo para obter IDs de nó.
        """
        if len(self.nodes) < 2:
            QMessageBox.warning(self, "Adicionar Ligação", "Você precisa de pelo menos dois nós para criar uma ligação.")
            return

        node_ids = [self.node_properties[n].get("id") for n in self.nodes if self.node_properties.get(n) and self.node_properties[n].get("id")]
        if not node_ids:
            QMessageBox.warning(self, "Erro", "Nenhum ID de nó válido encontrado. Adicione nós primeiro.")
            return

        source_id, ok1 = QInputDialog.getItem(self, "Adicionar Ligação", "Selecione o nó de ORIGEM:", node_ids, 0, False)
        if not ok1: return

        target_id, ok2 = QInputDialog.getItem(self, "Adicionar Ligação", "Selecione o nó de DESTINO:", node_ids, 0, False)
        if not ok2: return

        if source_id == target_id:
            QMessageBox.warning(self, "Erro", "Nós de origem e destino não podem ser os mesmos.")
            return

        source_node_obj = next((n for n in self.nodes if self.node_properties.get(n, {}).get("id") == source_id), None)
        target_node_obj = next((n for n in self.nodes if self.node_properties.get(n, {}).get("id") == target_id), None)

        if not source_node_obj or not target_node_obj:
            QMessageBox.critical(self, "Erro", "Um ou ambos os nós selecionados não foram encontrados.")
            return
        
        # Desenha a linha
        pen = QPen(Qt.darkGray, 2)
        line = self.scene.addLine(
            source_node_obj.rect().x() + source_node_obj.rect().width(), source_node_obj.rect().y() + source_node_obj.rect().height() / 2,
            target_node_obj.x(), target_node_obj.y() + target_node_obj.rect().height() / 2,
            pen
        )
        # Armazena os IDs dos nós conectados no próprio objeto de linha para salvar/carregar
        line.source_node_id = source_id
        line.target_node_id = target_id
        self.links.append(line)
        QMessageBox.information(self, "Ligação Adicionada", f"Ligação criada de '{source_id}' para '{target_id}'.")


    def _clear_diagram(self):
        """Limpa todos os elementos do diagrama e reinicia o contador de IDs."""
        self.scene.clear()
        self.nodes = [] 
        self.node_properties = {}
        self.links = []
        self.next_node_id = 1
        QMessageBox.information(self, "Diagrama Limpo", "O diagrama foi limpo.")

# Exemplo de uso (para testar este módulo individualmente)
if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Configura um caminho de teste para o ambiente de teste da tool
    project_root_test = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    user_sheets_dir_test = os.path.join(project_root_test, 'user_sheets')
    os.makedirs(user_sheets_dir_test, exist_ok=True)
    
    # Cria/Atualiza um arquivo engenharia.xlsx de teste
    test_file_path = os.path.join(user_sheets_dir_test, DEFAULT_DATA_EXCEL_FILENAME)
    if os.path.exists(test_file_path):
        os.remove(test_file_path) # Garante que começamos com um arquivo limpo
    
    # Cria um novo workbook e salva-o com as sheets padrão para engenharia
    wb_eng = openpyxl.Workbook()
    ws_estrutura = wb_eng.active
    ws_estrutura.title = "Estrutura"
    ws_estrutura.append(["part_number", "parent_part_number", "quantidade", "materia_prima"])
    ws_estrutura.append(["PROD-001", "", 1, "Não"])
    ws_estrutura.append(["COMP-001", "PROD-001", 2, "Não"])
    
    ws_workflow = wb_eng.create_sheet(DEFAULT_SHEET_NAME) # Cria a sheet 'Workflows' vazia inicialmente
    # A ferramenta adicionará os cabeçalhos do workflow na primeira vez que salvar.
    
    wb_eng.save(test_file_path)
    print(f"Arquivo de teste '{DEFAULT_DATA_EXCEL_FILENAME}' criado/atualizado com abas de exemplo.")

    window = EngenhariaWorkflowTool(file_path=test_file_path, sheet_name=DEFAULT_SHEET_NAME)
    window.show()
    sys.exit(app.exec_())
