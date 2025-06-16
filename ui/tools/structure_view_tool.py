import os
import openpyxl
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTreeWidget, QTreeWidgetItem, QLabel, QMessageBox, QHeaderView
from PyQt5.QtCore import Qt

# Define o caminho padrão para o arquivo Excel para dados de estrutura
# Agora pode aceitar um file_path no __init__
# STRUCTURE_EXCEL_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'user_sheets', 'workspace_data.xlsx')
# STRUCTURE_SHEET_NAME = "structure" # Planilha para relações pai-filho

class StructureViewTool(QWidget):
    """
    GUI para visualizar a estrutura hierárquica (e.g., BOM ou estrutura de arquivo)
    de um item selecionado ou de uma planilha específica em um arquivo Excel.
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
        self.structure_tree.setHeaderLabels(["ID do Componente", "Nome do Componente", "Tipo", "Quantidade", "Unidade"])
        self.structure_tree.header().setSectionResizeMode(QHeaderView.ResizeToContents)
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
            headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
            
            header_map = {header: idx for idx, header in enumerate(headers)}
            
            children_map = {}
            for row_idx in range(2, sheet.max_row + 1):
                row_values = [cell.value for cell in sheet[row_idx]]
                
                # Certifica-se de que os índices existem antes de tentar acessá-los
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
                        "Unit": unit
                    })
            
            # Tenta encontrar a estrutura para um item "raiz" conceitual ou o item inicial
            # Para BOMs, muitas vezes há um item "pai" que contém todos os outros
            root_item_id = "ROOT" # Um ID de item pai genérico se não houver um pai claro no BOM

            # Tenta encontrar o item raiz mais provável (o primeiro parentID que não é filho de ninguém)
            all_parent_ids = set(children_map.keys())
            all_child_ids = set(c_data["ID"] for children_list in children_map.values() for c_data in children_list)
            
            possible_roots = list(all_parent_ids - all_child_ids)
            
            if possible_roots:
                # Se houver raízes claras, use a primeira ou peça ao usuário para escolher
                root_item_id = possible_roots[0]
            elif children_map:
                # Se não houver raízes claras, mas há dados, pegue o primeiro pai como raiz
                root_item_id = list(children_map.keys())[0]
            else:
                 QMessageBox.information(self, "Nenhuma Estrutura Encontrada", f"Nenhum dado de estrutura encontrado na planilha '{self.sheet_name}' do arquivo '{os.path.basename(self.file_path)}'.")
                 self.structure_tree.addTopLevelItem(QTreeWidgetItem(["N/A", "Nenhum dado de estrutura.", "", "", ""]))
                 return
            
            self._add_items_to_tree(self.structure_tree.invisibleRootItem(), children_map, root_item_id)
            self.structure_tree.expandAll()

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Carregar Estrutura", f"Erro ao carregar dados da estrutura de '{os.path.basename(self.file_path)}' ({self.sheet_name}): {e}")
            self.structure_tree.addTopLevelItem(QTreeWidgetItem(["Erro", "Erro ao carregar dados.", "", "", ""]))

    def _add_items_to_tree(self, parent_qtree_item, children_map, current_item_id):
        """Adiciona recursivamente itens ao QTreeWidget com base nas relações pai-filho."""
        if current_item_id in children_map:
            for child_data in children_map[current_item_id]:
                q_item = QTreeWidgetItem([
                    str(child_data.get('ID', '')),
                    str(child_data.get('Name', '')),
                    str(child_data.get('Type', '')),
                    str(child_data.get('Quantity', '')),
                    str(child_data.get('Unit', ''))
                ])
                parent_qtree_item.addChild(q_item)
                self._add_items_to_tree(q_item, children_map, child_data['ID'])


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
        ws_test.title = "BOM" # Aba padrão para BOM Manager
        ws_test.append(["ID do BOM", "ID do Componente", "Nome do Componente", "Quantidade", "Unidade", "Ref Designator"])
        ws_test.append(["BOM-001", "PART-001", "Motor Elétrico", 1, "EA", "M1"])
        
        ws_structure_test = wb_test.create_sheet("structure")
        ws_structure_test.append(["ParentID", "ComponentID", "ComponentName", "Quantity", "Unit", "Type"])
        ws_structure_test.append(["PROD-001", "ASSY-001", "Produto Completo", 1, "EA", "Assembly"])
        ws_structure_test.append(["ASSY-001", "SUB-ASSY-001", "Sub-Montagem X", 1, "EA", "Assembly"])
        ws_structure_test.append(["ASSY-001", "PART-002", "Parafuso M5", 10, "PCS", "Part"])
        ws_structure_test.append(["SUB-ASSY-001", "PART-003", "Placa Controladora", 1, "EA", "Part"])
        ws_structure_test.append(["SUB-ASSY-001", "PART-004", "Sensor de Temperatura", 2, "PCS", "Part"])
        wb_test.save(test_file_path)
        print(f"Arquivo de teste '{test_file_path}' criado com dados de exemplo.")

    window = StructureViewTool(file_path=test_file_path, sheet_name="structure")
    window.show()
    sys.exit(app.exec_())
