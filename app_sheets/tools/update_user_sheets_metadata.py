import os
import openpyxl
import sys

# Define os caminhos de forma dinâmica a partir da localização do script
current_dir = os.path.dirname(os.path.abspath(__file__))
# Navega de app_sheets/tools para a raiz do projeto
project_root = os.path.dirname(os.path.dirname(current_dir)) 

USER_SHEETS_DIR = os.path.join(project_root, "user_sheets")
APP_SHEETS_DIR = os.path.join(project_root, "app_sheets")
# O db.xlsx está agora em user_sheets, conforme corrigimos anteriormente
DB_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "db.xlsx")

# Garante que os diretórios existam
os.makedirs(USER_SHEETS_DIR, exist_ok=True)
os.makedirs(APP_SHEETS_DIR, exist_ok=True)
# Garante que a pasta 'tools' dentro de 'app_sheets' exista, onde este script reside
os.makedirs(os.path.dirname(os.path.abspath(__file__)), exist_ok=True)


def _load_existing_db_db_data():
    """
    Carrega os dados existentes da planilha 'db_db' para um dicionário de lookup.
    Retorna: um dicionário onde a chave é (caminho_relativo_arquivo, nome_coluna)
             e o valor é {'pagina_arquivo': ..., 'descr_variavel': ...}.
    """
    existing_data = {}
    try:
        if not os.path.exists(DB_EXCEL_PATH):
            print(f"Aviso: O arquivo db.xlsx não foi encontrado em {DB_EXCEL_PATH}. Criando um novo ao salvar.")
            return existing_data # Retorna vazio se o arquivo não existe

        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        if "db_db" not in wb.sheetnames:
            print("Aviso: A planilha 'db_db' não foi encontrada em db.xlsx.")
            return existing_data # Retorna vazio se a planilha não existe

        sheet = wb["db_db"]
        headers = [cell.value for cell in sheet[1]] if sheet.max_row > 0 else []
        header_map = {h: idx for idx, h in enumerate(headers)}

        # Garante que as colunas essenciais para o lookup existem
        required_headers = ["Arquivo (Caminho)", "Nome da Coluna (Cabeçalho)", "pagina_arquivo", "descr_variavel"]
        if not all(h in header_map for h in required_headers):
            print(f"Aviso: A planilha 'db_db' não possui todos os cabeçalhos esperados para metadados: {', '.join(required_headers)}")
            return existing_data # Não podemos carregar corretamente sem os cabeçalhos

        for row_idx in range(2, sheet.max_row + 1):
            row_values = [cell.value for cell in sheet[row_idx]]
            
            # Acessa os valores de forma segura usando o mapa de cabeçalhos
            file_path_raw = row_values[header_map["Arquivo (Caminho)"]] if "Arquivo (Caminho)" in header_map and header_map["Arquivo (Caminho)"] < len(row_values) else None
            column_name = row_values[header_map["Nome da Coluna (Cabeçalho)"]] if "Nome da Coluna (Cabeçalho)" in header_map and header_map["Nome da Coluna (Cabeçalho)"] < len(row_values) else None
            pagina_arquivo = row_values[header_map["pagina_arquivo"]] if "pagina_arquivo" in header_map and header_map["pagina_arquivo"] < len(row_values) else None
            descr_variavel = row_values[header_map["descr_variavel"]] if "descr_variavel" in header_map and header_map["descr_variavel"] < len(row_values) else None

            # Use o caminho relativo normalizado como chave para consistência
            # o.path.relpath calcula o caminho relativo de 'file_path_raw' em relação a 'project_root'
            # e .replace('\\', '/') normaliza as barras para que funcionem em diferentes OS.
            if file_path_raw and column_name:
                normalized_path = os.path.relpath(file_path_raw, project_root).replace('\\', '/')
                existing_data[(normalized_path, str(column_name))] = {
                    'pagina_arquivo': pagina_arquivo if pagina_arquivo is not None else "",
                    'descr_variavel': descr_variavel if descr_variavel is not None else ""
                }
    except Exception as e:
        print(f"Erro ao carregar dados existentes de db_db: {e}")
    return existing_data

def _update_db_db_schema():
    """
    Coleta cabeçalhos de todos os arquivos Excel nas pastas user_sheets e app_sheets
    e os salva na planilha 'db_db' em db.xlsx, preservando descrições existentes.
    """
    print("\nIniciando atualização dos metadados das planilhas...")
    collected_headers_data = []
    
    # Carrega os dados existentes de 'db_db' para preservar as descrições
    existing_db_db_data = _load_existing_db_db_data()

    # Define os diretórios a serem escaneados
    directories_to_scan = [USER_SHEETS_DIR, APP_SHEETS_DIR]

    for base_dir in directories_to_scan:
        for root, _, files in os.walk(base_dir):
            for file_name in files:
                # Processa apenas arquivos .xlsx que não são temporários e não é o próprio db.xlsx
                # Também ignora os scripts Python dentro da pasta 'tools' se estivermos em app_sheets
                if file_name.endswith(".xlsx") and not file_name.startswith('~$') and file_name.lower() != "db.xlsx":
                    file_path = os.path.join(root, file_name)
                    
                    # Ignora pastas de scripts dentro de app_sheets para evitar tentar ler .py como .xlsx
                    if file_path.startswith(os.path.join(APP_SHEETS_DIR, "tools")):
                        continue

                    try:
                        wb = openpyxl.load_workbook(file_path, read_only=True)
                        for sheet_name in wb.sheetnames:
                            sheet = wb[sheet_name]
                            
                            # Defensive check for potentially empty or malformed sheets before accessing sheet[1]
                            # If sheet.max_row is 0, it means the sheet is empty, no headers to read.
                            if sheet.max_row == 0:
                                print(f"Aviso: Planilha '{sheet_name}' em '{os.path.basename(file_path)}' está vazia. Nenhuns cabeçalhos para coletar.")
                                continue # Skip to the next sheet

                            # Explicitly read headers from the first row, handling None values safely
                            headers = []
                            # Iterate up to max_column, but also check if max_column is a valid integer.
                            # Some malformed files might have max_column as None or 0.
                            valid_max_column = sheet.max_column if isinstance(sheet.max_column, int) and sheet.max_column > 0 else 0
                            
                            if valid_max_column == 0:
                                print(f"Aviso: Planilha '{sheet_name}' em '{os.path.basename(file_path)}' parece não ter colunas válidas. Nenhuns cabeçalhos para coletar.")
                                continue # Skip if no valid columns

                            for col_idx in range(1, valid_max_column + 1):
                                cell_value = sheet.cell(row=1, column=col_idx).value
                                if cell_value is not None:
                                    headers.append(str(cell_value))
                                else:
                                    # If a header cell is None, we can choose to skip it or add an empty string.
                                    # Adding an empty string ensures column count consistency, but might not be desired.
                                    # For now, let's append an empty string to keep alignment.
                                    headers.append("") # Append empty string for None headers

                            if not headers or all(h == "" for h in headers): # If all collected headers are empty strings
                                print(f"Aviso: Planilha '{sheet_name}' em '{os.path.basename(file_path)}' não possui cabeçalhos válidos na primeira linha.")
                                continue # Skip if no meaningful headers found

                            # Obtém o caminho relativo do arquivo em relação à raiz do projeto
                            relative_file_path = os.path.relpath(file_path, project_root).replace('\\', '/')

                            for header_name in headers:
                                # Usa a tupla (relative_file_path, header_name) para procurar dados existentes
                                lookup_key = (relative_file_path, str(header_name))
                                
                                existing_entry_dict = existing_db_db_data.get(lookup_key) # This gives the dictionary or None
                                
                                # Preserva a descrição existente se encontrada, caso contrário, deixa em branco
                                descr_variavel = existing_entry_dict['descr_variavel'] if existing_entry_dict else ""
                                
                                collected_headers_data.append([
                                    relative_file_path,
                                    str(header_name),
                                    sheet_name, # 'pagina_arquivo' é o nome real da planilha que está sendo lida
                                    descr_variavel
                                ])
                        print(f"Coletado cabeçalhos de: {os.path.basename(file_path)}")
                    except Exception as e:
                        print(f"Erro ao processar arquivo {file_name}: {e}")

    try:
        # Abre o db.xlsx. Se não existir, cria um novo.
        if os.path.exists(DB_EXCEL_PATH):
            wb_db = openpyxl.load_workbook(DB_EXCEL_PATH)
            # Remove a planilha 'db_db' se ela já existir para recriá-la com os dados atualizados
            if "db_db" in wb_db.sheetnames:
                del wb_db["db_db"]
        else:
            wb_db = openpyxl.Workbook()
            # Se é um novo workbook, garante que a primeira sheet seja ativa ou renomeia para evitar 'Sheet'
            if "Sheet" in wb_db.sheetnames and wb_db.sheetnames.index("Sheet") == 0:
                ws = wb_db.active
                ws.title = "users" # Cria 'users' como default se for novo (apenas um placeholder)
                # ws.append(["id", "username", "password_hash", "role", "full_name", "email", "phone", "department"]) # Comentar se não for criar a estrutura aqui

        ws_db_db = wb_db.create_sheet("db_db")
        # Define os cabeçalhos para a planilha 'db_db'
        db_db_headers = ["Arquivo (Caminho)", "Nome da Coluna (Cabeçalho)", "pagina_arquivo", "descr_variavel"]
        ws_db_db.append(db_db_headers)
        
        # Adiciona todos os dados de cabeçalho coletados
        for row_data in collected_headers_data:
            ws_db_db.append(row_data)
        
        wb_db.save(DB_EXCEL_PATH)
        print(f"\nMetadados atualizados com sucesso em '{DB_EXCEL_PATH}', planilha 'db_db'.")
    except Exception as e:
        print(f"Erro ao salvar metadados em db.xlsx: {e}")

if __name__ == "__main__":
    # Este script pode ser chamado com argumentos para diferentes ações
    if len(sys.argv) > 1:
        action = sys.argv[1]
        if action == "update_db_schema":
            _update_db_db_schema()
        else:
            print(f"Ação desconhecida: {action}")
    else:
        print("Uso: python update_user_sheets_metadata.py <ação>")
        print("Ações disponíveis: update_db_schema")

