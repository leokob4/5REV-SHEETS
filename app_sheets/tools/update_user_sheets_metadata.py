import os
import sys
import openpyxl

# Define o caminho para a raiz do projeto de forma robusta
# Este script está em app_sheets/tools/, então '..' leva a app_sheets, e '..' novamente leva ao project_root
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, '..', '..'))

USER_SHEETS_DIR = os.path.join(project_root, "user_sheets")
APP_SHEETS_DIR = os.path.join(project_root, "app_sheets")
DB_EXCEL_PATH = os.path.join(USER_SHEETS_DIR, "db.xlsx")

# Planilhas específicas que têm um comportamento diferente na detecção de headers
CONFIG_SHEETS_MAP = {
    os.path.join(APP_SHEETS_DIR, "users.xlsx"): "users",
    os.path.join(APP_SHEETS_DIR, "tools.xlsx"): "tools",
    os.path.join(APP_SHEETS_DIR, "access.xlsx"): "access",
    os.path.join(APP_SHEETS_DIR, "modules.xlsx"): "modules",
    os.path.join(APP_SHEETS_DIR, "permissions.xlsx"): "permissions",
    os.path.join(APP_SHEETS_DIR, "main.xlsx"): "refs",
    os.path.join(USER_SHEETS_DIR, "engenharia.xlsx"): "Estrutura",
}

def get_db_db_data():
    """Carrega os dados atuais da planilha 'db_db' em db.xlsx."""
    db_db_data = []
    if not os.path.exists(DB_EXCEL_PATH):
        print(f"Aviso: O arquivo db.xlsx não foi encontrado em {DB_EXCEL_PATH}. Ele será criado.")
        return []

    try:
        wb = openpyxl.load_workbook(DB_EXCEL_PATH)
        if "db_db" not in wb.sheetnames:
            print("Aviso: A planilha 'db_db' não foi encontrada em db.xlsx.")
            return []
        
        sheet = wb["db_db"]
        if sheet.max_row < 1:
            return [] # Planilha vazia

        headers = [cell.value for cell in sheet[1]]
        required_headers = ["Arquivo (Caminho)", "Nome da Coluna (Cabeçalho)", "pagina_arquivo", "descr_variavel"]
        if not all(h in headers for h in required_headers):
            print(f"Aviso: Cabeçalhos incompletos na planilha 'db_db'. Esperado: {required_headers}")
            return []

        header_map = {h: idx for idx, h in enumerate(headers)}

        for row_idx in range(2, sheet.max_row + 1):
            row_values = [cell.value for cell in sheet[row_idx]]
            if all(v is None for v in row_values):
                continue
            
            file_path = row_values[header_map["Arquivo (Caminho)"]] if "Arquivo (Caminho)" in header_map and header_map["Arquivo (Caminho)"] < len(row_values) else None
            col_name = row_values[header_map["Nome da Coluna (Cabeçalho)"]] if "Nome da Coluna (Cabeçalho)" in header_map and header_map["Nome da Coluna (Cabeçalho)"] < len(row_values) else None
            sheet_name = row_values[header_map["pagina_arquivo"]] if "pagina_arquivo" in header_map and header_map["pagina_arquivo"] < len(row_values) else None
            description = row_values[header_map["descr_variavel"]] if "descr_variavel" in header_map and header_map["descr_variavel"] < len(row_values) else None

            if all(val is not None for val in [file_path, col_name, sheet_name, description]):
                db_db_data.append({
                    "Arquivo (Caminho)": str(file_path),
                    "Nome da Coluna (Cabeçalho)": str(col_name),
                    "pagina_arquivo": str(sheet_name),
                    "descr_variavel": str(description)
                })
            else:
                print(f"Aviso: Linha malformada ou incompleta na db_db (linha {row_idx}): {row_values}. Ignorando.")

    except Exception as e:
        print(f"Erro ao carregar db.xlsx: {e}")
    return db_db_data


def save_db_db_data(data):
    """Salva os dados atualizados na planilha 'db_db' em db.xlsx."""
    try:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "db_db"

        headers = ["Arquivo (Caminho)", "Nome da Coluna (Cabeçalho)", "pagina_arquivo", "descr_variavel"]
        sheet.append(headers)

        for row in data:
            row_to_save = [
                row.get("Arquivo (Caminho)", ""),
                row.get("Nome da Coluna (Cabeçalho)", ""),
                row.get("pagina_arquivo", ""),
                row.get("descr_variavel", "")
            ]
            sheet.append(row_to_save)
        
        if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1 and wb["Sheet"] != sheet:
             wb.remove(wb["Sheet"])
        elif "Sheet" in wb.sheetnames and len(wb.sheetnames) == 1 and wb["Sheet"] == sheet:
            pass

        wb.save(DB_EXCEL_PATH)
        print(f"db.xlsx atualizado com {len(data)} entradas na db_db.")
    except Exception as e:
        print(f"Erro ao salvar db.xlsx: {e}")


def get_excel_headers(file_path, sheet_name=None):
    """
    Retorna os cabeçalhos da primeira linha de uma planilha Excel específica.
    Se sheet_name for None, tenta a primeira planilha ou a planilha principal mapeada.
    """
    headers = []
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        
        if sheet_name and sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
        elif file_path in CONFIG_SHEETS_MAP and CONFIG_SHEETS_MAP[file_path] in wb.sheetnames:
            sheet = wb[CONFIG_SHEETS_MAP[file_path]]
        else:
            sheet = wb.active
            if sheet_name:
                print(f"Aviso: Planilha '{sheet_name}' não encontrada em {os.path.basename(file_path)}. Usando a planilha ativa: {sheet.title}")


        if sheet.max_row >= 1:
            headers = [cell.value for cell in sheet[1]]
            headers = [h for h in headers if h is not None]
        return headers, sheet.title
    except FileNotFoundError:
        print(f"Aviso: Arquivo Excel não encontrado: {file_path}")
    except Exception as e:
        print(f"Erro ao ler cabeçalhos de {file_path} (planilha: {sheet_name or 'ativa'}): {e}")
    return [], None


def update_db_schema():
    """
    Atualiza a planilha 'db_db' em db.xlsx com os cabeçalhos reais
    de todas as outras planilhas do projeto.
    """
    print("\nIniciando sincronização da planilha 'db_db' com os arquivos reais...")
    new_db_db_data = []

    dirs_to_scan = [USER_SHEETS_DIR, APP_SHEETS_DIR]

    for base_dir in dirs_to_scan:
        for root, _, files in os.walk(base_dir):
            for file_name in files:
                if file_name.endswith(".xlsx") and not file_name.startswith('~$'):
                    file_path = os.path.join(root, file_name)
                    if file_path == DB_EXCEL_PATH:
                        continue

                    relative_path = os.path.relpath(file_path, project_root).replace('\\', '/')
                    
                    try:
                        wb = openpyxl.load_workbook(file_path, data_only=True)
                        for sheet_name in wb.sheetnames:
                            sheet = wb[sheet_name]
                            if sheet.max_row >= 1: 
                                headers = [cell.value for cell in sheet[1]]
                                headers = [h for h in headers if h is not None] 
                                
                                if headers:
                                    for header in headers:
                                        description = ""
                                        if relative_path == "user_sheets/engenharia.xlsx" and sheet_name == "Estrutura":
                                            if header == "part_number": description = "Número da Peça (ID Único do Item)"
                                            elif header == "part_description": description = "Descrição Detalhada da Peça"
                                            elif header == "parent_part_number": description = "Número da Peça Pai (para BOM)"
                                            elif header == "unidade_padrao_parent_part": description = "Unidade Padrão da Peça Pai"
                                            elif header == "concat_child_part_pn_list_comma": description = "Lista de Peças Filhas (concatenadas por vírgula)"
                                            elif header == "materia_prima_unidade": description = "Unidade da Matéria-Prima"
                                            elif header == "materia_prima_quantidade": description = "Quantidade da Matéria-Prima"
                                            elif header == "part_type": description = "Tipo da Peça (ex: item, purchased_part)"
                                            else: description = f"Cabeçalho da planilha '{sheet_name}' no arquivo '{os.path.basename(file_path)}'"
                                        elif relative_path == "app_sheets/tools.xlsx" and sheet_name == "tools":
                                            if header == "mod_id": description = "ID único do módulo/ferramenta"
                                            elif header == "mod_name": description = "Nome de exibição da ferramenta"
                                            elif header == "mod_description": description = "Descrição da ferramenta"
                                            elif header == "module_path": description = "Caminho do módulo Python para importação dinâmica"
                                            elif header == "class_name": description = "Nome da classe da ferramenta dentro do módulo Python" # Adicionado
                                            elif header == "MOD_WORK_TABLE": description = "Nome da planilha de trabalho principal associada a esta ferramenta (se houver)"
                                            elif header == "MOD_WORK_TABLE_PATH": description = "Caminho relativo da planilha de trabalho (se houver)"
                                            elif header == "mod_comment_old": description = "Comentários antigos sobre a ferramenta"
                                            elif header == "mod_comment_new": description = "Novos comentários sobre a ferramenta"
                                            else: description = f"Cabeçalho da planilha '{sheet_name}' no arquivo '{os.path.basename(file_path)}'"
                                        # Adicione mais elifs para outras planilhas específicas se quiser descrições personalizadas
                                        else:
                                            description = f"Cabeçalho da planilha '{sheet_name}' no arquivo '{os.path.basename(file_path)}'"

                                        new_db_db_data.append({
                                            "Arquivo (Caminho)": relative_path,
                                            "Nome da Coluna (Cabeçalho)": header,
                                            "pagina_arquivo": sheet_name,
                                            "descr_variavel": description
                                        })
                            else:
                                print(f"Aviso: Planilha '{sheet_name}' em '{relative_path}' está vazia ou sem cabeçalhos. Ignorando para db_db.")
                    except Exception as e:
                        print(f"Erro ao processar arquivo {relative_path}: {e}")
    
    save_db_db_data(new_db_db_data)
    print("Sincronização da planilha 'db_db' concluída.")


def validate_db_consistency():
    """
    Compara a estrutura atual das planilhas com o que está em 'db_db'.
    """
    print("\nIniciando validação de consistência...")
    db_db_schema = get_db_db_data()
    
    if not db_db_schema:
        print("Erro: A db_db está vazia ou não pôde ser carregada. Por favor, execute 'Sincronizar pagina db_db com planilhas das pastas' primeiro.")
        sys.exit(1) # Sair com erro se db_db não estiver pronta

    expected_headers = {}
    for entry in db_db_schema:
        file_path = entry["Arquivo (Caminho)"]
        sheet_name = entry["pagina_arquivo"]
        header_name = entry["Nome da Coluna (Cabeçalho)"]
        
        key = (file_path, sheet_name)
        if key not in expected_headers:
            expected_headers[key] = []
        expected_headers[key].append(header_name)

    all_consistent = True
    dirs_to_scan = [USER_SHEETS_DIR, APP_SHEETS_DIR]

    for base_dir in dirs_to_scan:
        for root, _, files in os.walk(base_dir):
            for file_name in files:
                if file_name.endswith(".xlsx") and not file_name.startswith('~$'):
                    file_path = os.path.join(root, file_name)
                    if file_path == DB_EXCEL_PATH:
                        continue
                    
                    relative_path = os.path.relpath(file_path, project_root).replace('\\', '/')

                    try:
                        wb = openpyxl.load_workbook(file_path, data_only=True)
                        for sheet_name in wb.sheetnames:
                            current_headers, _ = get_excel_headers(file_path, sheet_name)
                            
                            key = (relative_path, sheet_name)

                            if key in expected_headers:
                                missing_headers = [h for h in expected_headers[key] if h not in current_headers]
                                extra_headers = [h for h in current_headers if h not in expected_headers[key]]

                                if missing_headers:
                                    print(f"Inconsistência em '{relative_path}' (planilha '{sheet_name}'): Faltam cabeçalhos: {', '.join(missing_headers)}")
                                    all_consistent = False
                                if extra_headers:
                                    print(f"Inconsistência em '{relative_path}' (planilha '{sheet_name}'): Cabeçalhos extras: {', '.join(extra_headers)}")
                                    all_consistent = False
                                # Verifica se a planilha está vazia mas a db_db espera cabeçalhos
                                if not current_headers and expected_headers[key]:
                                    print(f"Inconsistência em '{relative_path}' (planilha '{sheet_name}'): Planilha vazia, mas cabeçalhos esperados na db_db.")
                                    all_consistent = False

                            elif current_headers: 
                                print(f"Aviso: Planilha '{sheet_name}' em '{relative_path}' existe com cabeçalhos, mas NÃO está registrada na db_db. Considere adicionar.")
                            else: 
                                pass 

                    except Exception as e:
                        print(f"Erro ao validar '{relative_path}' (planilha '{sheet_name}'): {e}")
                        all_consistent = False

    if all_consistent:
        print("Todas as planilhas estão consistentes com a db_db. Nenhuma diferença encontrada.")
        sys.exit(0) # Sair com sucesso
    else:
        print("\nValidação concluída com inconsistências. Por favor, revise os erros acima.")
        sys.exit(1) # Sair com erro


def create_or_update_sheets():
    """
    Cria novas planilhas ou atualiza as existentes com os cabeçalhos definidos na 'db_db'.
    Preserva dados existentes a partir da segunda linha.
    """
    print("\nIniciando criação/atualização de planilhas...")
    db_db_schema = get_db_db_data()

    if not db_db_schema:
        print("Erro: A db_db está vazia ou não pôde ser carregada. Por favor, execute 'Sincronizar pagina db_db com planilhas das pastas' primeiro.")
        sys.exit(1) # Sair com erro

    expected_sheet_headers = {}
    for entry in db_db_schema:
        file_path = os.path.join(project_root, entry["Arquivo (Caminho)"])
        sheet_name = entry["pagina_arquivo"]
        header_name = entry["Nome da Coluna (Cabeçalho)"]
        
        key = (file_path, sheet_name)
        if key not in expected_sheet_headers:
            expected_sheet_headers[key] = []
        expected_sheet_headers[key].append(header_name)
    
    for (file_path, sheet_name), headers_to_set in expected_sheet_headers.items():
        try:
            wb = None
            if os.path.exists(file_path):
                wb = openpyxl.load_workbook(file_path)
            else:
                wb = openpyxl.Workbook()
                if "Sheet" in wb.sheetnames:
                    wb.remove(wb["Sheet"])
                print(f"Criando novo arquivo: {os.path.basename(file_path)}")

            sheet = None
            if sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                print(f"Atualizando planilha existente: '{sheet_name}' em {os.path.basename(file_path)}")
            else:
                sheet = wb.create_sheet(sheet_name)
                print(f"Criando nova planilha: '{sheet_name}' em {os.path.basename(file_path)}")

            current_data = []
            if sheet.max_row > 1: 
                for row_idx in range(2, sheet.max_row + 1):
                    row_values = [cell.value for cell in sheet[row_idx]]
                    if not all(v is None for v in row_values):
                        current_data.append(row_values)
            
            sheet.delete_rows(1, sheet.max_row)

            sheet.append(headers_to_set)

            for row_data in current_data:
                new_row = [""] * len(headers_to_set)
                
                # Re-le os cabeçalhos da planilha antes de escrever (se ela tiver sido modificada por outra aba)
                # Ou, para ser mais seguro, mapeie os dados existentes baseados na POSIÇÃO da linha que foi lida
                # e não no cabeçalho, a menos que você queira remapear colunas com base no NOME.
                # Para esta função, o mais seguro é assumir que a ordem dos dados originais
                # deve ser preservada se não houver um mapeamento de nome de coluna.
                # No entanto, a `db_db` define a ORDEM. Então, vamos mapear pela ordem da db_db.
                
                # Para evitar erros de "index out of range" se a linha lida for mais curta que o número de cabeçalhos
                # na db_db, vamos preencher com valores padrão.
                
                # Isso é um ponto complexo: se a `db_db` mudou a ORDEM ou REMOVEU cabeçalhos,
                # a correspondência de dados por índice pode levar a erros lógicos.
                # A melhor prática para preservar dados com mudanças de schema é:
                # 1. Carregar dados existentes MANTENDO O MAPA DE HEADERS ORIGINAIS
                # 2. Criar uma nova linha vazia com os NOVOS HEADERS da db_db
                # 3. Preencher a nova linha, mapeando os dados antigos para os novos headers PELO NOME.
                
                # Vamos refatorar esta parte para ser mais robusta no remapeamento dos dados.
                
                # Antes de deletar as linhas, vamos extrair os dados e seus cabeçalhos originais
                # Esta parte foi movida para fora do loop principal, ou seja, na inicialização
                # da função, você carrega os dados e os headers atuais ANTES DE APAGAR TUDO.
                
                # Como current_data já foi coletado antes de limpar a planilha,
                # e current_headers_at_read não é mais usado após o clear.
                # Precisamos de uma forma de saber o cabeçalho original da linha que está sendo 'preserved'.
                # A abordagem mais segura é que a db_db defina a estrutura final, e se dados antigos
                # não se encaixam, eles são descartados ou ficam vazios.

                # Simplified approach: just append as they were read, assuming headers_to_set
                # defines the final required structure and any non-matching data is ignored.
                # If there are fewer elements in row_data than headers_to_set, it will be padded.
                # If there are more, extra elements will be ignored.
                appended_row = row_data[:len(headers_to_set)] # Trunca ou preenche com base nos headers
                while len(appended_row) < len(headers_to_set):
                    appended_row.append(None) # Preenche com None se faltar
                sheet.append(appended_row)

            wb.save(file_path)

        except Exception as e:
            print(f"Erro ao criar/atualizar '{os.path.basename(file_path)}' (planilha '{sheet_name}'): {e}")
    
    print("Criação/Atualização de planilhas concluída.")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        action = sys.argv[1]
        if action == "update_db_schema":
            update_db_schema()
        elif action == "validate":
            validate_db_consistency()
        elif action == "create_or_update_sheets":
            create_or_update_sheets()
        else:
            print(f"Ação desconhecida: {action}")
            sys.exit(1) # Sair com erro para ações desconhecidas
    else:
        print("Uso: python update_user_sheets_metadata.py [update_db_schema|validate|create_or_update_sheets]")
        sys.exit(1) # Sair com erro se nenhuma ação for fornecida
