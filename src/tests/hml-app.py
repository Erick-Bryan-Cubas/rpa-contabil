import os
import json
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv('envs/.env')

# Autenticação
username = os.getenv('usuario')
password = os.getenv('senha')
caminho_geral = os.getenv('caminho_geral')

# URL base do SharePoint
site_url = "https://planningassessoriaetributos-my.sharepoint.com/personal/arquivo_planning_com_br"

# Autenticação do contexto
ctx_auth = AuthenticationContext(site_url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(site_url, ctx_auth)
else:
    print(ctx_auth.get_last_error())
    exit(1)

# Caminho do arquivo JSON de configuração
config_file_path = 'configs/folders_test.json'

# Função para listar subpastas que contêm a pasta "Fiscal"
def list_fiscal_subfolders(ctx, folder_url):
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    subfolders = folder.folders
    ctx.load(subfolders)
    ctx.execute_query()

    fiscal_subfolders = []
    for subfolder in subfolders:
        subfolder_url = subfolder.serverRelativeUrl
        subfolder_folders = ctx.web.get_folder_by_server_relative_url(subfolder_url).folders
        ctx.load(subfolder_folders)
        ctx.execute_query()
        for sf in subfolder_folders:
            if sf.properties['Name'] == "Fiscal":
                fiscal_subfolders.append(subfolder.properties['Name'])
                break

    return fiscal_subfolders

# Função principal
def main():
    # Carregar configurações do arquivo JSON
    with open(config_file_path, 'r') as config_file:
        config = json.load(config_file)
    
    folders_access = {folder['Folder name'] for folder in config['FOLDERS_ACCESS']}
    folders_ignored = {folder['Folder name'] for folder in config['FOLDERS_IGNORED']}

    # Caminho relativo do SharePoint para a pasta geral
    folder_relative_url = '/personal/arquivo_planning_com_br/Documents/Arquivos/Carteiras 2023/Carteira Eduardo'

    # Listar pastas na pasta geral
    folders = ctx.web.get_folder_by_server_relative_url(folder_relative_url).folders
    ctx.load(folders)
    ctx.execute_query()

    for folder in folders:
        folder_name = folder.properties['Name']
        if folder_name in folders_access and folder_name not in folders_ignored:
            fiscal_subfolders = list_fiscal_subfolders(ctx, folder.serverRelativeUrl)
            if fiscal_subfolders:
                print(f"\nFolder: {folder_name}")
                for subfolder in fiscal_subfolders:
                    print(f"  Contains Fiscal subfolder: {subfolder}")

if __name__ == "__main__":
    main()
