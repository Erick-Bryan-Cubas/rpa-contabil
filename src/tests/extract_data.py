import os
import json
import logging
from datetime import datetime
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File
from dotenv import load_dotenv

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Carregar variáveis de ambiente
load_dotenv('envs/.env')

# Autenticação
username = os.getenv('usuario')
password = os.getenv('senha')
caminho_geral = os.getenv('caminho_geral')
caminho_landing_zone = os.getenv('caminho_landing_zone')

# URL base do SharePoint
site_url = "https://planningassessoriaetributos-my.sharepoint.com/personal/arquivo_planning_com_br"
logging.info("Autenticando no SharePoint...")

# Autenticação do contexto
ctx_auth = AuthenticationContext(site_url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(site_url, ctx_auth)
    logging.info("Autenticação no SharePoint bem-sucedida.")
else:
    logging.error("Erro na autenticação do SharePoint: %s", ctx_auth.get_last_error())
    exit(1)

# Autenticação para a landing zone
landing_zone_url = "https://planningassessoriaetributos-my.sharepoint.com/personal/erick_bryan_planning_com_br"
logging.info("Autenticando na landing zone...")

landing_zone_auth = AuthenticationContext(landing_zone_url)
if landing_zone_auth.acquire_token_for_user(username, password):
    ctx_landing_zone = ClientContext(landing_zone_url, landing_zone_auth)
    logging.info("Autenticação na landing zone bem-sucedida.")
else:
    logging.error("Erro na autenticação da landing zone: %s", landing_zone_auth.get_last_error())
    exit(1)

# Caminho do arquivo JSON de configuração
config_file_path = 'configs/folders_test.json'

# Função para procurar e copiar arquivos
def search_and_copy_files(ctx, folder_url, target_ctx, target_folder_sped, target_folder_relatorios, target_folder_guias, month_year):
    logging.info("Procurando arquivos em: %s", folder_url)
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    subfolders = folder.folders
    files = folder.files
    ctx.load(subfolders)
    ctx.load(files)
    ctx.execute_query()

    # Procurar arquivos txt na pasta Sped Contribuições
    for file in files:
        file_name = file.properties['Name']
        if file_name.startswith("SPED_PISCOFINS") and file_name.endswith(".txt"):
            new_file_name = f"{month_year}_{file_name}"
            logging.info("Encontrado arquivo SPED_PISCOFINS: %s", file_name)
            copy_file(ctx, file, target_ctx, target_folder_sped, new_file_name)

    # Procurar arquivos pdf na subpasta Composição
    for subfolder in subfolders:
        if subfolder.properties['Name'] == "Composição":
            composition_folder_url = subfolder.serverRelativeUrl
            composition_folder = ctx.web.get_folder_by_server_relative_url(composition_folder_url)
            composition_files = composition_folder.files
            ctx.load(composition_files)
            ctx.execute_query()
            for comp_file in composition_files:
                if comp_file.properties['Name'].endswith(".pdf"):
                    logging.info("Encontrado arquivo PDF em Composição: %s", comp_file.properties['Name'])
                    copy_file(ctx, comp_file, target_ctx, target_folder_relatorios, comp_file.properties['Name'])

def search_and_copy_guias(ctx, fiscal_folder_url, target_ctx, target_folder_guias):
    logging.info("Procurando arquivos na pasta Guias Impostos: %s", fiscal_folder_url)
    fiscal_folder = ctx.web.get_folder_by_server_relative_url(fiscal_folder_url)
    subfolders = fiscal_folder.folders
    ctx.load(subfolders)
    ctx.execute_query()

    for subfolder in subfolders:
        if subfolder.properties['Name'] == "Guias Impostos":
            guias_folder_url = subfolder.serverRelativeUrl
            guias_folder = ctx.web.get_folder_by_server_relative_url(guias_folder_url)
            guias_subfolders = guias_folder.folders
            ctx.load(guias_subfolders)
            ctx.execute_query()
            for guias_subfolder in guias_subfolders:
                if guias_subfolder.properties['Name'] == "Federal":
                    federal_folder_url = guias_subfolder.serverRelativeUrl
                    federal_folder = ctx.web.get_folder_by_server_relative_url(federal_folder_url)
                    federal_files = federal_folder.files
                    ctx.load(federal_files)
                    ctx.execute_query()
                    for federal_file in federal_files:
                        if federal_file.properties['Name'].endswith(".pdf"):
                            logging.info("Encontrado arquivo PDF em Guias Impostos: %s", federal_file.properties['Name'])
                            copy_file(ctx, federal_file, target_ctx, target_folder_guias, federal_file.properties['Name'])

# Função para copiar arquivos
def copy_file(ctx, source_file, target_ctx, target_folder_url, new_file_name):
    source_file_url = source_file.serverRelativeUrl
    logging.info("Copiando arquivo de %s para %s", source_file_url, os.path.join(target_folder_url, new_file_name))
    file_content = File.open_binary(ctx, source_file_url).content
    target_folder = target_ctx.web.get_folder_by_server_relative_url(target_folder_url)
    target_folder.upload_file(new_file_name, file_content).execute_query()
    logging.info("Arquivo %s copiado com sucesso para %s", new_file_name, target_folder_url)

# Função para listar subpastas que contêm a pasta "Fiscal"
def list_fiscal_subfolders(ctx, folder_url):
    logging.info("Listando subpastas em: %s", folder_url)
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
                logging.info("Encontrada subpasta Fiscal em: %s", subfolder_url)
                break

    return fiscal_subfolders

# Função principal
def main():
    # Carregar configurações do arquivo JSON
    with open(config_file_path, 'r') as config_file:
        config = json.load(config_file)
    logging.info("Configurações carregadas com sucesso.")

    folders_access = {folder['Folder name'] for folder in config['FOLDERS_ACCESS']}
    folders_ignored = {folder['Folder name'] for folder in config['FOLDERS_IGNORED']}

    # Caminho relativo do SharePoint para a pasta geral
    folder_relative_url = '/personal/arquivo_planning_com_br/Documents/Arquivos/Carteiras 2023/Carteira Eduardo'

    # Listar pastas na pasta geral
    logging.info("Listando pastas na pasta geral: %s", folder_relative_url)
    folders = ctx.web.get_folder_by_server_relative_url(folder_relative_url).folders
    ctx.load(folders)
    ctx.execute_query()

    for folder in folders:
        folder_name = folder.properties['Name']
        if folder_name in folders_access and folder_name not in folders_ignored:
            fiscal_subfolders = list_fiscal_subfolders(ctx, folder.serverRelativeUrl)
            for subfolder in fiscal_subfolders:
                for year in range(2024, 2025):  # Modifique este range conforme necessário
                    for month in range(1, 13):
                        month_folder = f"{month:02d}-{year}"
                        fiscal_folder_url = os.path.join(folder.serverRelativeUrl, subfolder, "Fiscal", str(year), month_folder)
                        sped_folder_url = os.path.join(fiscal_folder_url, "Sped Contribuições")
                        logging.info("Procurando arquivos na pasta: %s", sped_folder_url)
                        search_and_copy_files(ctx, sped_folder_url, ctx_landing_zone, "/personal/erick_bryan_planning_com_br/Documents/landing_zone/SpedContribuicoes", "/personal/erick_bryan_planning_com_br/Documents/landing_zone/RelatoriosContasRecebidas", "/personal/erick_bryan_planning_com_br/Documents/landing_zone/GuiasImpostos", month_folder)
                        
                        guias_folder_url = os.path.join(fiscal_folder_url, "Guias Impostos")
                        logging.info("Procurando arquivos na pasta: %s", guias_folder_url)
                        search_and_copy_guias(ctx, fiscal_folder_url, ctx_landing_zone, "/personal/erick_bryan_planning_com_br/Documents/landing_zone/GuiasImpostos")

if __name__ == "__main__":
    main()
