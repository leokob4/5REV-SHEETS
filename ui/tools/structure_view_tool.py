import os
import openpyxl
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QTreeWidget, QTreeWidgetItem, QLabel, QMessageBox, QHeaderView
from PyQt5.QtCore import Qt

# Define the path to the Excel file for structure data
STRUCTURE_EXCEL_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 'user_sheets', 'workspace_data.xlsx')
STRUCTURE_SHEET_NAME = "structure" # Sheet for parent-child relationships

class StructureViewTool(QWidget):
    """
    GUI for viewing the hierarchical structure (e.g., BOM or file structure)
    of a selected item from the workspace.
    """
    def __init__(self, item_id, item_name):
        super().__init__()
        self.item_id = item_id
        self.item_name = item_name
        self.setWindowTitle(f"Estrutura do Item: {self.item_name} ({self.item_id})")
        self.layout = QVBoxLayout(self)

        self.layout.addWidget(QLabel(f"<h2>Estrutura para: {self.item_name} ({self.item_id})</h2>"))
        self.layout.addWidget(QLabel("Exibindo subcomponentes e documentos relacionados."))

        self.structure_tree = QTreeWidget()
        self.structure_tree.setHeaderLabels(["ID", "Nome", "Tipo", "Quantidade", "Unidade"])
        self.structure_tree.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.layout.addWidget(self.structure_tree)

        self._load_structure_data()

    def _load_structure_data(self):
        """Loads and displays the hierarchical structure for the given item_id."""
        self.structure_tree.clear()
        
        try:
            if not os.path.exists(STRUCTURE_EXCEL_PATH):
                QMessageBox.warning(self, "Arquivo de Estrutura Não Encontrado", f"O arquivo de dados de estrutura não foi encontrado: {STRUCTURE_EXCEL_PATH}.")
                # Optionally create a dummy structure if the file is missing
                self._add_dummy_structure_if_missing()
                return

            wb = openpyxl.load_workbook(STRUCTURE_EXCEL_PATH)
            if STRUCTURE_SHEET_NAME not in wb.sheetnames:
                QMessageBox.warning(self, "Planilha de Estrutura Não Encontrada", f"A planilha '{STRUCTURE_SHEET_NAME}' não foi encontrada em '{STRUCTURE_EXCEL_PATH}'.")
                # Optionally create a dummy structure if the sheet is missing
                self._add_dummy_structure_if_missing()
                return

            sheet = wb[STRUCTURE_SHEET_NAME]
            headers = [cell.value for cell in sheet[1]]
            
            # Create a map for quick column access
            header_map = {header: idx for idx, header in enumerate(headers)}
            
            # Group children by parent ID
            children_map = {}
            for row_idx in range(2, sheet.max_row + 1):
                row_values = [cell.value for cell in sheet[row_idx]]
                
                # Use .get() with a default of None to handle missing columns gracefully
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
            
            # Find and display children of the current item
            if self.item_id in children_map:
                self._add_items_to_tree(self.structure_tree.invisibleRootItem(), children_map, self.item_id)
                self.structure_tree.expandAll() # Expand all nodes in structure view
            else:
                QMessageBox.information(self, "Nenhuma Estrutura Encontrada", f"Nenhum subcomponente ou documento encontrado para o item '{self.item_name}' (ID: {self.item_id}).")
                self.structure_tree.addTopLevelItem(QTreeWidgetItem(["N/A", "Nenhum subcomponente.", "", "", ""]))

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Carregar Estrutura", f"Erro ao carregar dados da estrutura para '{self.item_name}': {e}")
            self.structure_tree.addTopLevelItem(QTreeWidgetItem(["Erro", "Erro ao carregar dados.", "", "", ""]))

    def _add_items_to_tree(self, parent_qtree_item, children_map, current_item_id):
        """Recursively adds items to the QTreeWidget based on parent-child relationships."""
        if current_item_id in children_map:
            for child_data in children_map[current_item_id]:
                q_item = QTreeWidgetItem([
                    str(child_data.get('ID', '')), # Ensure string conversion
                    str(child_data.get('Name', '')),
                    str(child_data.get('Type', '')),
                    str(child_data.get('Quantity', '')),
                    str(child_data.get('Unit', ''))
                ])
                parent_qtree_item.addChild(q_item)
                # Recursively add children of this child
                self._add_items_to_tree(q_item, children_map, child_data['ID'])

    def _add_dummy_structure_if_missing(self):
        """
        Creates a dummy structure in the tree if the Excel file/sheet is missing,
        to provide a visual representation of how structure would look.
        """
        QMessageBox.information(self, "Estrutura de Exemplo", "Exibindo uma estrutura de exemplo. Por favor, crie o arquivo 'workspace_data.xlsx' com a planilha 'structure' para dados reais.")
        
        # Clear existing
        self.structure_tree.clear()

        # Add dummy data for the selected item's children
        if self.item_id == "PROJ-001":
            q_item_assy = QTreeWidgetItem(["ASSY-001", "Assembly-001 (Exemplo)", "Assembly", "1", "EA"])
            self.structure_tree.addTopLevelItem(q_item_assy)
            q_item_assy.addChild(QTreeWidgetItem(["PART-001", "Part-001 (Exemplo)", "Part", "2", "PCS"]))
            q_item_assy.addChild(QTreeWidgetItem(["COMP-001", "Component-XYZ (Exemplo)", "Component", "1", "PCS"]))
            self.structure_tree.addTopLevelItem(QTreeWidgetItem(["DRAW-001", "Drawing-CAD-001 (Exemplo)", "Document", "1", "EA"]))
        elif self.item_id == "ASSY-001":
            self.structure_tree.addTopLevelItem(QTreeWidgetItem(["PART-001", "Part-001 (Exemplo)", "Part", "2", "PCS"]))
            self.structure_tree.addTopLevelItem(QTreeWidgetItem(["COMP-001", "Component-XYZ (Exemplo)", "Component", "1", "PCS"]))
        else:
            self.structure_tree.addTopLevelItem(QTreeWidgetItem(["N/A", "Nenhum subcomponente (Exemplo)", "", "", ""]))

        self.structure_tree.expandAll()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    # For standalone testing, you can pass dummy IDs
    window = StructureViewTool("PROJ-001", "Demo Project - Rev A")
    window.show()
    sys.exit(app.exec_())
