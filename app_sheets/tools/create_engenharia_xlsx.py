import os
import openpyxl

# Define o caminho do diretório user_sheets
# O script assume que estará em '5REV-SHEETS/app_sheets/tools/create_engenharia_xlsx.py'
# project_root deve apontar para '5REV-SHEETS'
project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
user_sheets_dir = os.path.join(project_root, 'user_sheets') # CORRIGIDO: Agora aponta para 'user_sheets'
os.makedirs(user_sheets_dir, exist_ok=True) # Garante que o diretório user_sheets exista

file_path = os.path.join(user_sheets_dir, "engenharia.xlsx")
sheet_name = "Estrutura" # Nome da planilha para a estrutura de engenharia

# Define os cabeçalhos para o arquivo engenharia.xlsx
# Estes cabeçalhos serão a base para a criação inicial da planilha.
ENGENHARIA_HEADERS = [
    "part_number", "parent_part_number", "quantidade", "materia_prima"
]

# Dados de exemplo para a estrutura de engenharia
# Representa uma árvore de componentes e matérias-primas
sample_data = [
    # Produto Final (parent_part_number vazio ou N/A)
    ["PROD-001", "", 1, "Não"], # Produto Principal
    
    # Submontagens de PROD-001
    ["ASSY-A", "PROD-001", 1, "Não"], # Submontagem A
    ["ASSY-B", "PROD-001", 2, "Não"], # Submontagem B
    
    # Componentes de ASSY-A
    ["COMP-001", "ASSY-A", 5, "Não"], # Componente pré-fabricado
    ["RAW-MAT-001", "ASSY-A", 10, "Sim"], # Matéria-prima 1
    
    # Componentes de ASSY-B
    ["COMP-002", "ASSY-B", 3, "Não"], # Componente 2
    ["RAW-MAT-002", "ASSY-B", 1, "Sim"], # Matéria-prima 2
    ["RAW-MAT-003", "ASSY-B", 0.5, "Sim"], # Matéria-prima 3 (por kg/metro)
    
    # Componentes de COMP-001 (se for uma sub-estrutura interna, por exemplo)
    ["SUB-COMP-01", "COMP-001", 1, "Não"],
    ["RAW-MAT-004", "COMP-001", 20, "Sim"]
]

def create_engenharia_xlsx():
    """Cria ou atualiza o arquivo engenharia.xlsx com os cabeçalhos e dados de exemplo."""
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet_name

        # Adiciona os cabeçalhos
        ws.append(ENGENHARIA_HEADERS)

        # Adiciona os dados de exemplo
        for row_data in sample_data:
            ws.append(row_data)

        wb.save(file_path)
        print(f"Arquivo '{os.path.basename(file_path)}' criado/atualizado com a planilha '{sheet_name}'.")
    except Exception as e:
        print(f"Erro ao criar/atualizar {os.path.basename(file_path)}: {e}")

if __name__ == "__main__":
    create_engenharia_xlsx()
