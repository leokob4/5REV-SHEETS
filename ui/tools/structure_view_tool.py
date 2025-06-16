import os
import openpyxl
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTreeWidget, QTreeWidgetItem, QLabel, QMessageBox, QHeaderView
from PyQt5.QtCore import Qt

class StructureViewTool(QWidget):
    """
    GUI para visualizar a estrutura hierárquica (e.g., BOM ou estrutura de arquivo)
    de um item selecionado ou de uma planilha específica em um arquivo Excel.
    Os cabeçalhos da árvore são dinamicamente carregados do arquivo Excel.
    """
    def __init__(self, file_path, sheet_name="structure"):
        super().__init__()
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.setWindowTitle(f"Estrutura: {os.path.basename(self.file_path)} ({self.sheet_name})")
        self.layout = QVBoxLayout(self)

        self.layout.addWidget(QLabel(f"<h2>Estrutura do Arquivo: {os.path.basename(self.file_path)}</h2>"))
        self.layout.addWidget(QLabel(f"Exibindo estrutura da planilha: <b>{self.sheet_name}</b>"))

        self.structure_tree = QTreeWidget()
        # Cabeçalhos serão definidos dinamicamente em _load_structure_data
        self.structure_tree.header().setSectionResizeMode(QHeaderView.Interactive)
        self.structure_tree.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.layout.addWidget(self.structure_tree)

        self._load_structure_data()

    def _load_structure_data(self):
        """Carrega e exibe a estrutura hierárquica do arquivo e planilha especificados."""
        self.structure_tree.clear()
        
        if not self.file_path or not os.path.exists(self.file_path):
            QMessageBox.warning(self, "Arquivo Não Encontrado", f"O arquivo de estrutura não foi encontrado: {self.file_path}.")
            self.structure_tree.addTopLevelItem(QTreeWidgetItem(["N/A", "Arquivo não encontrado.", "", "", ""]))
            return

        try:
            wb = openpyxl.load_workbook(self.file_path)
            if self.sheet_name not in wb.sheetnames:
                QMessageBox.warning(self, "Planilha Não Encontrada", f"A planilha '{self.sheet_name}' não foi encontrada em '{os.path.basename(self.file_path)}'.")
                self.structure_tree.addTopLevelItem(QTreeWidgetItem(["N/A", "Planilha não encontrada.", "", "", ""]))
                return

            sheet = wb[self.sheet_name]
            
            # Carrega os cabeçalhos da primeira linha da planilha
            headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
            
            # Define cabeçalhos padrão se a planilha estiver vazia, ou usa os do arquivo
            if not headers:
                headers = ["ParentID", "ComponentID", "ComponentName", "Quantity", "Unit", "Type"]
            
            self.structure_tree.setHeaderLabels(headers) # Define os cabeçalhos da árvore dinamicamente
            
            header_map = {header: idx for idx, header in enumerate(headers)}
            
            children_map = {}
            for row_idx in range(2, sheet.max_row + 1):
                row_values = [cell.value for cell in sheet[row_idx]]
                
                # Garante que row_values tenha tamanho suficiente antes de acessar índices
                parent_id = row_values[header_map.get("ParentID")] if "ParentID" in header_map and header_map.get("ParentID") < len(row_values) else None
                component_id = row_values[header_map.get("ComponentID")] if "ComponentID" in header_map and header_map.get("ComponentID") < len(row_values) else None
                component_name = row_values[header_map.get("ComponentName")] if "ComponentName" in header_map and header_map.get("ComponentName") < len(row_values) else None
                quantity = row_values[header_map.get("Quantity")] if "Quantity" in header_map and header_map.get("Quantity") < len(row_values) else None
                unit = row_values[header_map.get("Unit")] if "Unit" in header_map and header_map.get("Unit") < len(row_values) else None
                item_type = row_values[header_map.get("Type")] if "Type" in header_map and header_map.get("Type") < len(row_values) else None

                if parent_id and component_id and component_name:
                    if parent_id not in children_map:
                        children_map[parent_id] = []
                    children_map[parent_id].append({
                        "ID": component_id,
                        "Name": component_name,
                        "Type": item_type,
                        "Quantity": quantity,
                        "Unit": unit,
                        "OriginalRowValues": row_values # Armazena os valores originais da linha para manter ordem
                    })
            
            root_item_id = None
            all_parent_ids = set(children_map.keys())
            all_child_ids = set(c_data["ID"] for children_list in children_map.values() for c_data in children_list)
            
            possible_roots = list(all_parent_ids - all_child_ids)
            
            if possible_roots:
                root_item_id = possible_roots[0]
            elif children_map:
                root_item_id = list(children_map.keys())[0] # Fallback if no clear root (e.g., cyclic or fragmented data)
            
            if root_item_id:
                # Recupera o "root" item (primeiro item da estrutura que não é filho de ninguém)
                root_row_values = None
                for row_idx in range(2, sheet.max_row + 1):
                    row_values = [cell.value for cell in sheet[row_idx]]
                    if row_values[header_map.get("ComponentID")] == root_item_id:
                        root_row_values = row_values
                        break

                if root_row_values:
                    root_q_item = QTreeWidgetItem([str(v) if v is not None else "" for v in root_row_values])
                    self.structure_tree.addTopLevelItem(root_q_item)
                    self._add_items_to_tree(root_q_item, children_map, root_item_id, headers)
                else: # Fallback for cases where root_item_id is found in children_map but not explicitly in a row (e.g., first row is a child)
                    QMessageBox.information(self, "Aviso", f"Item raiz '{root_item_id}' encontrado, mas sua linha correspondente não pôde ser identificada para preenchimento completo. Exibindo apenas a subestrutura.")
                    self._add_items_to_tree(self.structure_tree.invisibleRootItem(), children_map, root_item_id, headers)

                self.structure_tree.expandAll()
            else:
                 QMessageBox.information(self, "Nenhuma Estrutura Encontrada", f"Nenhum dado de estrutura hierárquica válido encontrado na planilha '{self.sheet_name}' do arquivo '{os.path.basename(self.file_path)}'. Verifique se 'ParentID' e 'ComponentID' estão presentes e corretos.")
                 self.structure_tree.addTopLevelItem(QTreeWidgetItem(["N/A", "Nenhum dado de estrutura.", "", "", ""]))

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Carregar Estrutura", f"Erro ao carregar dados da estrutura de '{os.path.basename(self.file_path)}' ({self.sheet_name}): {e}")
            self.structure_tree.addTopLevelItem(QTreeWidgetItem(["Erro", "Erro ao carregar dados.", "", "", ""]))

    def _add_items_to_tree(self, parent_qtree_item, children_map, current_item_id, headers):
        """Adiciona recursivamente itens ao QTreeWidget com base nas relações pai-filho."""
        if current_item_id in children_map:
            for child_data in children_map[current_item_id]:
                # Prepara os dados para o QTreeWidgetItem na ordem dos cabeçalhos
                item_values = []
                for header in headers:
                    if header == "ParentID": # ParentID não é exibido como coluna para o filho direto
                        item_values.append("") 
                    elif header == "ComponentID":
                        item_values.append(str(child_data.get('ID', '')))
                    elif header == "ComponentName":
                        item_values.append(str(child_data.get('Name', '')))
                    elif header == "Quantity":
                        item_values.append(str(child_data.get('Quantity', '')))
                    elif header == "Unit":
                        item_values.append(str(child_data.get('Unit', '')))
                    elif header == "Type":
                        item_values.append(str(child_data.get('Type', '')))
                    else: # Adiciona outros campos se existirem na linha original do Excel
                        # Isso tenta preservar a ordem e incluir todas as colunas
                        original_values = child_data.get('OriginalRowValues', [])
                        header_map_for_child = {h: idx for idx, h in enumerate(headers)}
                        if header in header_map_for_child and header_map_for_child[header] < len(original_values):
                             item_values.append(str(original_values[header_map_for_child[header]]))
                        else:
                             item_values.append("") # Coluna não encontrada nos dados ou valor nulo
                
                # Se houver menos valores do que cabeçalhos, preenche o restante com vazios
                while len(item_values) < len(headers):
                    item_values.append("")

                q_item = QTreeWidgetItem([str(v) if v is not None else "" for v in item_values])
                parent_qtree_item.addChild(q_item)
                self._add_items_to_tree(q_item, children_map, child_data['ID'], headers) # Passa os headers para a recursão


if __name__ == "__main__":
    app = QApplication(sys.argv)
    project_root_test = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
    test_file_dir = os.path.join(project_root_test, 'user_sheets')
    os.makedirs(test_file_dir, exist_ok=True)
    test_file_path = os.path.join(test_file_dir, "bom_data.xlsx") # Exemplo de arquivo com estrutura

    # Criar um arquivo de exemplo com uma planilha 'structure' se não existir
    if not os.path.exists(test_file_path):
        wb_test = openpyxl.Workbook()
        ws_test = wb_test.active
        ws_test.title = "BOM"
        ws_test.append(["ID do BOM", "ID do Componente", "Nome do Componente", "Quantidade", "Unidade", "Ref Designator"])
        
        ws_structure_test = wb_test.create_sheet("structure")
        ws_structure_test.append(["ParentID", "ComponentID", "ComponentName", "Quantity", "Unit", "Type", "Notes"]) # Adicionado "Notes" para teste dinâmico
        ws_structure_test.append(["", "PROD-001", "Produto Completo", 1, "EA", "Assembly", "Notas do produto principal"]) # Root item
        ws_structure_test.append(["PROD-001", "ASSY-001", "Sub-Montagem X", 1, "EA", "Assembly", "Montagem de teste"])
        ws_structure_test.append(["PROD-001", "PART-002", "Parafuso M5", 10, "PCS", "Part", "Parafuso padrão"])
        ws_structure_test.append(["ASSY-001", "PART-003", "Placa Controladora", 1, "EA", "Part", "Placa versão 2.0"])
        ws_structure_test.append(["ASSY-001", "PART-004", "Sensor de Temperatura", 2, "PCS", "Part", "Sensor para ambiente X"])
        ws_structure_test.append(["PART-003", "RES-001", "Resistor 10k", 5, "PCS", "Component", "Componente eletrônico"])
        wb_test.save(test_file_path)
        print(f"Arquivo de teste '{test_file_path}' criado com dados de exemplo.")

    window = StructureViewTool(file_path=test_file_path, sheet_name="structure")
    window.show()
    sys.exit(app.exec_())
