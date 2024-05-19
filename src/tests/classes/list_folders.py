import os
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.runtime.auth.user_credential import UserCredential
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

# Função para listar pastas
def list_folders(ctx, folder_url):
    folders = ctx.web.get_folder_by_server_relative_url(folder_url).folders
    ctx.load(folders)
    ctx.execute_query()

    for folder in folders:
        print(f"Folder name: {folder.properties['Name']}")

# Caminho relativo do SharePoint para a pasta geral
folder_relative_url = '/personal/arquivo_planning_com_br/Documents/Arquivos/Carteiras 2023/Carteira Eduardo'

# Listar pastas
list_folders(ctx, folder_relative_url)
