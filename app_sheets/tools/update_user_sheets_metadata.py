import os
import openpyxl
from openpyxl.utils import get_column_letter

# Define os caminhos dos diretórios relativos à raiz do projeto
# O script assume que estará em '5REV-SHEETS/app_sheets/tools/update_user_sheets_metadata.py'
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
USER_SHEETS_DIR = os.path.join(PROJECT_ROOT, "user_sheets")
APP_SHEETS_DIR = os.path.join(PROJECT_ROOT, "app_sheets")
DB_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "db.xlsx")

def get_db_updated_schema():
    """
    Carrega o esquema de cabeçalhos desejado da planilha 'db_updated' em db.xlsx.
    Esta é a fonte de verdade para os cabeçalhos que devem ser aplicados.
    Retorna um dicionário aninhado:
    {
        'caminho/do/arquivo.xlsx': {
            'NomeDaPlanilha': ['Header1', 'Header2', ...] # Lista ORDENADA de nomes de cabeçalho
        }
    }
    """
    desired_schema = {}
    try:
        if not os.path.exists(DB_EXCEL_PATH):
            print(f"Erro: O arquivo db.xlsx não foi encontrado em: {DB_EXCEL_PATH}")
            return desired_schema

        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        if "db_updated" not in wb.sheetnames: # Modificado para usar 'db_updated'
            print(f"Erro: A planilha 'db_updated' não foi encontrada em {DB_EXCEL_PATH}")
            return desired_schema

        sheet = wb["db_updated"]
        headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
        
        header_map = {
            "Arquivo (Caminho)": -1,
            "Nome da Coluna (Cabeçalho)": -1,
            "pagina_arquivo": -1
            # 'descr_variavel' não é necessário para esta função, pois só precisamos dos nomes das colunas
        }
        for idx, h in enumerate(headers):
            if h in header_map:
                header_map[h] = idx

        required_headers = ["Arquivo (Caminho)", "Nome da Coluna (Cabeçalho)", "pagina_arquivo"]
        if any(header_map[key] == -1 for key in required_headers):
            print(f"Aviso: A planilha 'db_updated' em {DB_EXCEL_PATH} não possui todos os cabeçalhos essenciais: {required_headers}. "
                  "O carregamento do esquema desejado pode estar incompleto.")
            
        # Dicionário temporário para construir a lista ordenada de cabeçalhos por arquivo/planilha
        temp_ordered_headers = {}

        for row_idx in range(2, sheet.max_row + 1):
            row_values = [cell.value for cell in sheet[row_idx]]
            
            file_path_raw = row_values[header_map["Arquivo (Caminho)"]] if header_map["Arquivo (Caminho)"] != -1 and header_map["Arquivo (Caminho)"] < len(row_values) else None
            column_name = row_values[header_map["Nome da Coluna (Cabeçalho)"]] if header_map["Nome da Coluna (Cabeçalho)"] != -1 and header_map["Nome da Coluna (Cabeçalho)"] < len(row_values) else None
            sheet_name_from_db = row_values[header_map["pagina_arquivo"]] if header_map["pagina_arquivo"] != -1 and header_map["pagina_arquivo"] < len(row_values) else None

            if file_path_raw and column_name and sheet_name_from_db:
                normalized_file_path = file_path_raw.replace('\\', '/')
                
                if normalized_file_path not in temp_ordered_headers:
                    temp_ordered_headers[normalized_file_path] = {}
                
                if sheet_name_from_db not in temp_ordered_headers[normalized_file_path]:
                    temp_ordered_headers[normalized_file_path][sheet_name_from_db] = []

                temp_ordered_headers[normalized_file_path][sheet_name_from_db].append(str(column_name))
            else:
                print(f"Aviso: Ignorando linha incompleta/malformada em 'db_updated' (linha {row_idx}): {row_values}")

        # Converte as listas para a estrutura final de 'desired_schema'
        for file_path, sheets_data in temp_ordered_headers.items():
            desired_schema[file_path] = {}
            for sheet_name, headers_list in sheets_data.items():
                desired_schema[file_path][sheet_name] = headers_list

    except Exception as e:
        print(f"Erro ao carregar schema de db_updated: {e}")
    return desired_schema

def update_excel_headers(file_path, sheet_name, new_headers):
    """
    Atualiza os cabeçalhos de uma planilha Excel, preservando os dados existentes.
    Se a planilha não existir, ela será criada.
    Se os cabeçalhos mudarem, os dados serão reordenados ou preenchidos/removidos.
    """
    try:
        wb = None
        if os.path.exists(file_path):
            wb = openpyxl.load_workbook(file_path)
        else:
            wb = openpyxl.Workbook()
            # Remove a sheet padrão 'Sheet' se for um workbook novo
            if 'Sheet' in wb.sheetnames:
                del wb['Sheet']

        if sheet_name not in wb.sheetnames:
            ws = wb.create_sheet(sheet_name)
            print(f"Criada nova planilha '{sheet_name}' em '{os.path.basename(file_path)}'.")
        else:
            ws = wb[sheet_name]

        # Coleta os cabeçalhos atuais (se existirem) e os dados existentes
        current_headers = [cell.value for cell in ws[1] if cell.value is not None] if ws.max_row > 0 else []
        # Garante que os dados sejam lidos antes de qualquer modificação de coluna
        existing_data = []
        if ws.max_row > 1:
            for row_idx in range(2, ws.max_row + 1):
                row_data = [ws.cell(row=row_idx, column=col_idx).value for col_idx in range(1, ws.max_column + 1)]
                existing_data.append(row_data)

        # Mapeamento de cabeçalhos antigos para novos índices
        old_header_to_index = {header: i for i, header in enumerate(current_headers)}
        
        # Cria uma nova lista de dados reordenados
        reordered_data = []
        for row_original in existing_data:
            new_row = [None] * len(new_headers) # Inicializa a nova linha com None
            for new_col_idx, new_header in enumerate(new_headers):
                if new_header in old_header_to_index:
                    old_col_idx = old_header_to_index[new_header]
                    if old_col_idx < len(row_original):
                        new_row[new_col_idx] = row_original[old_col_idx]
            reordered_data.append(new_row)

        # Limpa o conteúdo da planilha (cabeçalhos e dados)
        # Atenção: ws.delete_rows(1, ws.max_row) pode ser perigoso se ws.max_row for 0 ou 1
        # É mais seguro iterar e limpar ou recriar a planilha
        if ws.max_row > 0:
            for row in list(ws.rows): # Percorre uma cópia das linhas para evitar problemas ao deletar
                ws.delete_rows(row[0].row, 1) # Deleta uma linha por vez, a partir da primeira célula da linha

        # Adiciona os novos cabeçalhos
        ws.append(new_headers)

        # Adiciona os dados reordenados
        for row in reordered_data:
            ws.append(row)

        wb.save(file_path)
        print(f"✔ Planilha '{sheet_name}' em '{os.path.basename(file_path)}' atualizada com os novos cabeçalhos.")

    except Exception as e:
        print(f"Erro ao atualizar cabeçalhos da planilha '{sheet_name}' em '{os.path.basename(file_path)}': {e}")
        # É crucial não propagar o erro para evitar que uma falha em um arquivo pare todo o processo
        return False
    return True


def run_header_update():
    """
    Função principal para executar a atualização dos cabeçalhos em massa.
    """
    print("\n--- Iniciando atualização de cabeçalhos com base em 'db_updated' ---")
    desired_schema = get_db_updated_schema()

    if not desired_schema:
        print("Nenhum esquema desejado carregado de 'db_updated'. Nenhuma atualização será realizada.")
        return

    updated_count = 0
    failed_count = 0

    # Itera sobre os arquivos e planilhas definidos no esquema desejado
    for file_rel_path, sheets_data in desired_schema.items():
        # Converte o caminho relativo para absoluto
        full_file_path = os.path.join(PROJECT_ROOT, file_rel_path)
        
        # Garante que o diretório existe antes de tentar salvar
        os.makedirs(os.path.dirname(full_file_path), exist_ok=True)

        for sheet_name, new_headers in sheets_data.items():
            print(f"Processando: Arquivo '{os.path.basename(full_file_path)}', Planilha '{sheet_name}'...")
            if update_excel_headers(full_file_path, sheet_name, new_headers):
                updated_count += 1
            else:
                failed_count += 1

    print(f"\n--- Atualização de cabeçalhos concluída ---")
    print(f"Total de planilhas atualizadas com sucesso: {updated_count}")
    print(f"Total de planilhas com falha na atualização: {failed_count}")
    if failed_count > 0:
        print("Verifique os logs acima para detalhes das falhas.")

if __name__ == "__main__":
    run_header_update()

