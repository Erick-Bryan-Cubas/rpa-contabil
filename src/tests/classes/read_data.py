import os
import json
import logging
from PyPDF2 import PdfReader
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
caminho_landing_zone = os.getenv('caminho_landing_zone')

# URL base da landing zone do SharePoint
landing_zone_url = "https://planningassessoriaetributos-my.sharepoint.com/personal/erick_bryan_planning_com_br"
logging.info("Autenticando na landing zone...")

# Autenticação do contexto
ctx_auth = AuthenticationContext(landing_zone_url)
if ctx_auth.acquire_token_for_user(username, password):
    ctx = ClientContext(landing_zone_url, ctx_auth)
    logging.info("Autenticação na landing zone bem-sucedida.")
else:
    logging.error("Erro na autenticação da landing zone: %s", ctx_auth.get_last_error())
    exit(1)

# Função para listar arquivos PDF na pasta GuiasImpostos
def list_pdfs(ctx, folder_url):
    logging.info("Listando arquivos PDF na pasta: %s", folder_url)
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    files = folder.files
    ctx.load(files)
    ctx.execute_query()

    pdf_files = [file.properties['Name'] for file in files if file.properties['Name'].endswith(".pdf")]
    return pdf_files

# Função para ler o conteúdo de um PDF
def read_pdf_content(ctx, folder_url, pdf_name):
    pdf_url = os.path.join(folder_url, pdf_name)
    file_content = File.open_binary(ctx, pdf_url).content

    with open(pdf_name, 'wb') as pdf_file:
        pdf_file.write(file_content)

    with open(pdf_name, 'rb') as pdf_file:
        pdf_reader = PdfReader(pdf_file)
        pdf_text = ''
        for page_num in range(len(pdf_reader.pages)):
            pdf_page = pdf_reader.pages[page_num]
            pdf_text += pdf_page.extract_text()
    
    os.remove(pdf_name)  # Remover o arquivo PDF baixado após a leitura

    return pdf_text

# Função para corrigir caracteres, extrair o CNPJ, nome da empresa, valor total, data de vencimento, data de apuração, número do documento, código e descrição do imposto
def extract_data(pdf_content):
    corrected_content = pdf_content.replace('�', 'ã').replace('�', 'ç')
    
    # Extrair CNPJ e nome da empresa
    start_index = corrected_content.find("Documento de Arrecadação\nde Receitas Federais\n \n") + len("Documento de Arrecadação\nde Receitas Federais\n \n")
    cnpj = corrected_content[start_index:start_index + 18]
    end_index = corrected_content.find("\nPeríodo de Apuração", start_index)
    company_name = corrected_content[start_index + 19:end_index].strip()
    
    # Extrair valor total do documento
    value_start_index = corrected_content.find("Valor Total do Documento\n") + len("Valor Total do Documento\n")
    value_end_index = corrected_content.find("CNPJ", value_start_index)
    total_value = corrected_content[value_start_index:value_end_index].strip()
    
    # Extrair data de vencimento
    due_date_start_index = corrected_content.find("Pagar este documento até\n") + len("Pagar este documento até\n")
    due_date_end_index = corrected_content.find("Observações", due_date_start_index)
    due_date = corrected_content[due_date_start_index:due_date_end_index].strip()
    
    # Extrair data de apuração
    apuration_date_start_index = corrected_content.find("Razão Social\n") + len("Razão Social\n")
    apuration_date_end_index = corrected_content.find(" ", apuration_date_start_index)
    apuration_date = corrected_content[apuration_date_start_index:apuration_date_end_index].strip()

    # Extrair número do documento
    doc_number_start_index = corrected_content.find("Número do Documento\n") + len("Número do Documento\n")
    doc_number_end_index = corrected_content.find("Pagar este", doc_number_start_index)
    doc_number = corrected_content[doc_number_start_index:doc_number_end_index].strip()

    # Extrair código e descrição do imposto
    tax_code_start_index = corrected_content.find("Total Multa JurosComposição do Documento de Arrecadação\n") + len("Total Multa JurosComposição do Documento de Arrecadação\n")
    tax_code = corrected_content[tax_code_start_index:tax_code_start_index + 4].strip()
    tax_description = corrected_content[tax_code_start_index + 4:tax_code_start_index + 40].strip()
    
    return cnpj, company_name, total_value, due_date, apuration_date, doc_number, tax_code, tax_description

# Função para salvar os dados extraídos em um arquivo JSON
def save_data_to_json(cnpj, company_name, total_value, due_date, apuration_date, doc_number, tax_code, tax_description, output_filename):
    with open(output_filename, 'w') as json_file:
        json.dump({
            "CNPJ": cnpj,
            "Company Name": company_name,
            "Total Value": total_value,
            "Due Date": due_date,
            "Apuration Date": apuration_date,
            "Document Number": doc_number,
            "Tax Code": tax_code,
            "Tax Description": tax_description
        }, json_file, ensure_ascii=False, indent=4)
    logging.info("Dados salvos em %s", output_filename)

# Função principal
def main():
    # Caminho relativo do SharePoint para a pasta GuiasImpostos na landing zone
    guias_folder_url = '/personal/erick_bryan_planning_com_br/Documents/landing_zone/GuiasImpostos'

    # Listar arquivos PDF na pasta GuiasImpostos
    pdf_files = list_pdfs(ctx, guias_folder_url)
    logging.info("Arquivos PDF encontrados: %s", pdf_files)

    # Ler o conteúdo de um PDF específico como teste
    pdf_name = 'Infinity - Darf de COFINS 01.2024.pdf'
    if pdf_name in pdf_files:
        logging.info("Lendo o PDF: %s", pdf_name)
        pdf_content = read_pdf_content(ctx, guias_folder_url, pdf_name)

        # Extrair o CNPJ, o nome da empresa, o valor total, a data de vencimento, a data de apuração, o número do documento, o código e a descrição do imposto
        cnpj, company_name, total_value, due_date, apuration_date, doc_number, tax_code, tax_description = extract_data(pdf_content)

        # Salvar os dados em um arquivo JSON
        output_filename = pdf_name.replace('.pdf', '_data.json')
        save_data_to_json(cnpj, company_name, total_value, due_date, apuration_date, doc_number, tax_code, tax_description, output_filename)
    else:
        logging.error("O PDF %s não foi encontrado na pasta %s", pdf_name, guias_folder_url)

if __name__ == "__main__":
    main()
