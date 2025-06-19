import sys
import os
import subprocess

# Define o caminho para a raiz do projeto
# Este script está em app_sheets/tools/, então '..' leva a app_sheets, e '..' novamente leva ao project_root
current_dir = os.path.dirname(os.path.abspath(__file__))
project_root = os.path.abspath(os.path.join(current_dir, '..', '..'))

# O script principal de metadados está no mesmo diretório 'tools'
update_metadata_script_path = os.path.join(current_dir, 'update_user_sheets_metadata.py')

def run_validation():
    """
    Executa a ação 'validate' do script update_user_sheets_metadata.py.
    """
    print(f"Iniciando validação através de: {update_metadata_script_path}")
    
    if not os.path.exists(update_metadata_script_path):
        print(f"Erro: Script de atualização de metadados não encontrado em {update_metadata_script_path}")
        sys.exit(1)

    try:
        # Chama o script update_user_sheets_metadata.py com a ação 'validate'
        result = subprocess.run(
            [sys.executable, update_metadata_script_path, "validate"],
            capture_output=True,
            text=True,
            check=True
        )
        print(result.stdout.strip())
        if result.stderr:
            print(f"Erro/Aviso do script de metadados:\n{result.stderr.strip()}")
        sys.exit(0) # Saída de sucesso
    except subprocess.CalledProcessError as e:
        print(f"Erro na execução do validador (Código: {e.returncode}):\n{e.stdout}\n{e.stderr}")
        sys.exit(1) # Saída de erro
    except Exception as e:
        print(f"Erro inesperado ao chamar o validador: {e}")
        sys.exit(1)

if __name__ == "__main__":
    run_validation()

