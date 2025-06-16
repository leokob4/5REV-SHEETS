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
    Possui suporte específico para o arquivo 'engenharia.xlsx'.
    """
    def __init__(self, file_path, sheet_name="Estrutura"): # Default sheet for engenharia
        super().__init__()
        self.file_path = file_path
        self.sheet_name = sheet_name
        
        self.is_engenharia_file = (os.path.basename(self.file_path) == "engenharia.xlsx")

        self.setWindowTitle(f"Estrutura: {os.path.basename(self.file_path)} ({self.sheet_name})")
        if self.is_engenharia_file:
            self.setWindowTitle(self.windowTitle() + " (Engenharia)")

        self.layout = QVBoxLayout(self)

        self.layout.addWidget(QLabel(f"<h2>Estrutura do Arquivo: {os.path.basename(self.file_path)}</h2>"))
        self.layout.addWidget(QLabel(f"Exibindo estrutura da planilha: <b>{self.sheet_name}</b>"))

        self.structure_tree = QTreeWidget()
        self.structure_tree.header().setSectionResizeMode(QHeaderView.Interactive)
        self.structure_tree.verticalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.layout.addWidget(self.structure_tree)

        self._load_structure_data()

    def _load_structure_data(self):
        """Carrega e exibe a estrutura hierárquica do arquivo e planilha especificados."""
        self.structure_tree.clear()
        
        if not self.file_path or not os.path.exists(self.file_path):
            QMessageBox.warning(self, "Arquivo Não Encontrado", f"O arquivo de estrutura não foi encontrado: {self.file_path}.")
            self.structure_tree.addTopLevelItem(QTreeWidgetItem(["N/A", "Arquivo não encontrado."]))
            return

        try:
            wb = openpyxl.load_workbook(self.file_path)
            if self.sheet_name not in wb.sheetnames:
                QMessageBox.warning(self, "Planilha Não Encontrada", f"A planilha '{self.sheet_name}' não foi encontrada em '{os.path.basename(self.file_path)}'.")
                self.structure_tree.addTopLevelItem(QTreeWidgetItem(["N/A", "Planilha não encontrada."]))
                return

            sheet = wb[self.sheet_name]
            
            headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
            
            if not headers:
                # Fallback headers based on file type
                if self.is_engenharia_file:
                    headers = ["part_number", "parent_part_number", "quantidade", "materia_prima"]
                else:
                    headers = ["ParentID", "ComponentID", "ComponentName", "Quantity", "Unit", "Type"]
            
            self.structure_tree.setHeaderLabels(headers)
            
            # Dynamically determine the parent and component ID columns
            parent_id_col = -1
            component_id_col = -1
            
            # Prioritize specific names for engenharia.xlsx
            if self.is_engenharia_file:
                try:
                    parent_id_col = headers.index("parent_part_number")
                    component_id_col = headers.index("part_number")
                except ValueError:
                    # Fallback if specific headers not found in engenharia.xlsx, but still engenharia file
                    QMessageBox.warning(self, "Aviso de Cabeçalho", "Cabeçalhos 'part_number' ou 'parent_part_number' não encontrados em engenharia.xlsx. Tentando mapeamento genérico.")

            # Generic fallback for other files or if specific names missing in engenharia.xlsx
            if parent_id_col == -1 and "ParentID" in headers:
                parent_id_col = headers.index("ParentID")
            if component_id_col == -1 and "ComponentID" in headers:
                component_id_col = headers.index("ComponentID")

            if parent_id_col == -1 or component_id_col == -1:
                QMessageBox.critical(self, "Erro de Cabeçalho", "Não foi possível identificar as colunas de 'part_number/ComponentID' e 'parent_part_number/ParentID'. Verifique os cabeçalhos da planilha.")
                self.structure_tree.addTopLevelItem(QTreeWidgetItem(["Erro", "Cabeçalhos de estrutura não encontrados."]))
                return

            children_map = {}
            all_component_ids = set()
            all_parent_ids_in_data = set()

            for row_idx in range(2, sheet.max_row + 1):
                row_values = [str(cell.value) if cell.value is not None else "" for cell in sheet[row_idx]]
                
                if len(row_values) > max(parent_id_col, component_id_col):
                    parent_id = row_values[parent_id_col].strip() if parent_id_col != -1 else ""
                    component_id = row_values[component_id_col].strip() if component_id_col != -1 else ""

                    if component_id: # Only process if component_id is not empty
                        all_component_ids.add(component_id)
                        if parent_id:
                            all_parent_ids_in_data.add(parent_id)
                            if parent_id not in children_map:
                                children_map[parent_id] = []
                            children_map[parent_id].append({"ID": component_id, "Data": row_values})
            
            root_items_ids = list(all_component_ids - all_parent_ids_in_data)

            if not root_items_ids:
                # If no clear root (e.g., all items are children, or circular reference, or single item)
                # Try to find an item that exists as a component but is not a parent of anything in the data
                # Or if there's only one component and it has no parent.
                if len(all_component_ids) == 1 and list(all_component_ids)[0] not in all_parent_ids_in_data:
                    root_items_ids = list(all_component_ids) # Single component as root
                elif all_component_ids:
                    # Fallback: if multiple items and no clear root, pick the first one with an empty parent_id
                    for row_idx in range(2, sheet.max_row + 1):
                        row_values = [str(cell.value) if cell.value is not None else "" for cell in sheet[row_idx]]
                        if len(row_values) > parent_id_col and row_values[parent_id_col].strip() == "":
                            root_items_ids.append(row_values[component_id_col].strip())
                            break
                    if not root_items_ids and all_component_ids: # Last resort, pick first component found
                        root_items_ids = [list(all_component_ids)[0]]

            if not root_items_ids:
                QMessageBox.information(self, "Nenhuma Estrutura Encontrada", f"Nenhum dado de estrutura hierárquica válido encontrado na planilha '{self.sheet_name}' do arquivo '{os.path.basename(self.file_path)}'. Verifique se 'parent_part_number/ParentID' e 'part_number/ComponentID' estão presentes e corretos.")
                self.structure_tree.addTopLevelItem(QTreeWidgetItem(["N/A", "Nenhum dado de estrutura."]))
                return

            # Add root items to the tree
            for root_id in root_items_ids:
                root_data = None
                for row_idx in range(2, sheet.max_row + 1):
                    row_values = [str(cell.value) if cell.value is not None else "" for cell in sheet[row_idx]]
                    if len(row_values) > component_id_col and row_values[component_id_col].strip() == root_id:
                        root_data = row_values
                        break
                
                if root_data:
                    root_q_item = QTreeWidgetItem(root_data)
                    self.structure_tree.addTopLevelItem(root_q_item)
                    self._add_items_to_tree(root_q_item, children_map, root_id, headers, parent_id_col, component_id_col)
                else: # This can happen if a root_id is identified but its row data isn't directly found (e.g., it's a parent but not a child itself)
                    # Create a dummy item for the root and add its children
                    dummy_root_item = QTreeWidgetItem([root_id] + [""] * (len(headers) - 1)) # Populate with root_id and empty values
                    self.structure_tree.addTopLevelItem(dummy_root_item)
                    self._add_items_to_tree(dummy_root_item, children_map, root_id, headers, parent_id_col, component_id_col)
                    QMessageBox.warning(self, "Aviso de Estrutura", f"O item raiz '{root_id}' foi identificado, mas sua linha de dados completa não foi encontrada. Exibindo apenas a subestrutura.")


            self.structure_tree.expandAll()

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Carregar Estrutura", f"Erro ao carregar dados da estrutura de '{os.path.basename(self.file_path)}' ({self.sheet_name}): {e}")
            self.structure_tree.addTopLevelItem(QTreeWidgetItem(["Erro", "Erro ao carregar dados."]))

    def _add_items_to_tree(self, parent_qtree_item, children_map, current_item_id, headers, parent_id_col, component_id_col):
        """Adiciona recursivamente itens ao QTreeWidget com base nas relações pai-filho."""
        if current_item_id in children_map:
            for child_entry in children_map[current_item_id]:
                child_id = child_entry["ID"]
                child_row_values = child_entry["Data"]
                
                # Create a list for the QTreeWidgetItem that matches the header order
                item_values = [""] * len(headers)
                for h_idx, header_name in enumerate(headers):
                    if h_idx == parent_id_col: # Do not display parent_id in child's row
                        item_values[h_idx] = ""
                    elif h_idx < len(child_row_values):
                        item_values[h_idx] = child_row_values[h_idx]

                q_item = QTreeWidgetItem([str(v) if v is not None else "" for v in item_values])
                parent_qtree_item.addChild(q_item)
                self._add_items_to_tree(q_item, children_map, child_id, headers, parent_id_col, component_id_col)

# Exemplo de uso (para testar este módulo individualmente)
if __name__ == "__main__":
    app = QApplication(sys.argv)
    project_root_test = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
    test_file_dir = os.path.join(project_root_test, 'user_sheets')
    os.makedirs(test_file_dir, exist_ok=True)
    
    # Example for bom_data.xlsx (original structure view test)
    test_bom_file_path = os.path.join(test_file_dir, "bom_data.xlsx") 
    if not os.path.exists(test_bom_file_path):
        wb_test = openpyxl.Workbook()
        ws_test = wb_test.active
        ws_test.title = "BOM"
        ws_test.append(["ID do BOM", "ID do Componente", "Nome do Componente", "Quantidade", "Unidade", "Ref Designator"])
        
        ws_structure_test = wb_test.create_sheet("structure")
        ws_structure_test.append(["ParentID", "ComponentID", "ComponentName", "Quantity", "Unit", "Type", "Notes"]) 
        ws_structure_test.append(["", "PROD-001", "Produto Completo", 1, "EA", "Assembly", "Notas do produto principal"])
        ws_structure_test.append(["PROD-001", "ASSY-001", "Sub-Montagem X", 1, "EA", "Assembly", "Montagem de teste"])
        ws_structure_test.append(["PROD-001", "PART-002", "Parafuso M5", 10, "PCS", "Part", "Parafuso padrão"])
        ws_structure_test.append(["ASSY-001", "PART-003", "Placa Controladora", 1, "EA", "Part", "Placa versão 2.0"])
        ws_structure_test.append(["ASSY-001", "PART-004", "Sensor de Temperatura", 2, "PCS", "Part", "Sensor para ambiente X"])
        ws_structure_test.append(["PART-003", "RES-001", "Resistor 10k", 5, "PCS", "Component", "Componente eletrônico"])
        wb_test.save(test_bom_file_path)
        print(f"Arquivo de teste '{test_bom_file_path}' criado com dados de exemplo.")

    # Example for engenharia.xlsx (new test case)
    test_engenharia_file_path = os.path.join(test_file_dir, "engenharia.xlsx")
    if not os.path.exists(test_engenharia_file_path):
        wb_engenharia = openpyxl.Workbook()
        ws_engenharia = wb_engenharia.active
        ws_engenharia.title = "Estrutura"
        ws_engenharia.append(["part_number", "parent_part_number", "quantidade", "materia_prima", "descrição_extra"])
        ws_engenharia.append(["PROD-001", "", 1, "Não", "Produto final principal"])
        ws_engenharia.append(["ASSY-A", "PROD-001", 1, "Não", "Sub-montagem do motor"])
        ws_engenharia.append(["COMP-001", "ASSY-A", 5, "Não", "Motor elétrico"])
        ws_engenharia.append(["RAW-MAT-001", "ASSY-A", 10, "Sim", "Fio de cobre 1mm"])
        ws_engenharia.append(["SHEET-B", "COMP-001", 1, "Sim", "Chapa de aço"])
        wb_engenharia.save(test_engenharia_file_path)
        print(f"Arquivo de teste '{test_engenharia_file_path}' criado com dados de exemplo.")


    # Test with bom_data.xlsx
    # window_bom = StructureViewTool(file_path=test_bom_file_path, sheet_name="structure")
    # window_bom.show()

    # Test with engenharia.xlsx
    window_engenharia = StructureViewTool(file_path=test_engenharia_file_path, sheet_name="Estrutura")
    window_engenharia.show()

    sys.exit(app.exec_())
