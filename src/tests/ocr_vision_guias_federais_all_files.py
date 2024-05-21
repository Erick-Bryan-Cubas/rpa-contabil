import os
import logging
import fitz  # PyMuPDF
import base64
import json
import requests
from datetime import datetime
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File
import difflib
import re

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Carregar variáveis de ambiente do arquivo .env
load_dotenv('envs/.env')

# Variáveis de ambiente
google_credentials_path = os.getenv('GOOGLE_CLOUD_CREDENTIALS')
caminho_landing_zone = os.getenv('caminho_landing_zone')
folder_path = "GuiasImpostos"
usuario = os.getenv('usuario')
senha = os.getenv('senha')
api_key = os.getenv('API_KEY')

# Verificar se a API_KEY foi carregada corretamente
if not api_key:
    logging.error("API_KEY não encontrada. Certifique-se de que a variável está definida no arquivo .env")
    exit(1)

# Diretórios de armazenamento
pdf_dir = 'data/files'
images_dir = 'data/images'
output_dir = 'data/output'

# Criação dos diretórios se não existirem
os.makedirs(pdf_dir, exist_ok=True)
os.makedirs(images_dir, exist_ok=True)
os.makedirs(output_dir, exist_ok=True)

# Verificar se o arquivo de credenciais existe
if not os.path.exists(google_credentials_path):
    logging.error(f"Arquivo de credenciais não encontrado: {google_credentials_path}")
    raise FileNotFoundError(f"Arquivo de credenciais não encontrado: {google_credentials_path}")

# Autenticação
logging.info("Autenticando na landing zone...")
landing_zone_url = "https://planningassessoriaetributos-my.sharepoint.com/personal/erick_bryan_planning_com_br"
ctx_auth = AuthenticationContext(landing_zone_url)
if ctx_auth.acquire_token_for_user(usuario, senha):
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

# Função para baixar o PDF
def download_pdf(context, server_relative_url, file_path):
    response = File.open_binary(context, server_relative_url)

    with open(file_path, 'wb') as file:
        file.write(response.content)
    logging.info(f"Arquivo {file_path} baixado com sucesso.")

    # Verificar se o PDF pode ser aberto com PyMuPDF
    try:
        fitz.open(file_path)
    except fitz.FileDataError as e:
        logging.error(f"Erro ao abrir o arquivo PDF: {e}")
        # Log do conteúdo da resposta se não for um PDF válido
        logging.error(f"Conteúdo da resposta: {response.content.decode('utf-8')}")
        raise ValueError(f"Arquivo {file_path} não é um PDF válido.")

# Função para verificar se o PDF existe
def check_pdf_exists(file_path):
    if not os.path.exists(file_path):
        logging.error(f"Arquivo PDF não encontrado: {file_path}")
        raise FileNotFoundError(f"Arquivo PDF não encontrado: {file_path}")
    else:
        logging.info(f"Arquivo PDF encontrado: {file_path}")

# Função para converter PDF em imagens
def convert_pdf_to_images(pdf_path, images_dir):
    check_pdf_exists(pdf_path)
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap()
        image_path = os.path.join(images_dir, f"page_{page_num}.png")
        pix.save(image_path)
        images.append(image_path)
    logging.info(f"PDF convertido em {len(images)} imagens.")
    return images

# Função para extrair texto das imagens usando Google Cloud Vision API
def extract_text_from_images(images):
    texts = []
    url = f"https://vision.googleapis.com/v1/images:annotate?key={api_key}"
    headers = {'Content-Type': 'application/json'}

    for image_path in images:
        with open(image_path, "rb") as image_file:
            my_base64 = base64.b64encode(image_file.read()).decode('utf-8')
        
        data = {
            'requests': [
                {
                    'image': {
                        'content': my_base64
                    },
                    'features': [
                        {
                            'type': 'TEXT_DETECTION'
                        }
                    ]
                }
            ]
        }

        response = requests.post(url, headers=headers, data=json.dumps(data))
        response.raise_for_status()  # Lança um erro se a requisição falhar
        r = response.json()

        if 'error' in r:
            logging.error(f"Erro na resposta do Vision API: {r['error']['message']}")
            continue

        if 'textAnnotations' in r['responses'][0]:
            text = r['responses'][0]['textAnnotations'][0]['description']
            texts.append(text)
        else:
            texts.append("")

    return texts

# Função para encontrar termos semelhantes
def find_similar_term(term, lines):
    match = difflib.get_close_matches(term, lines, n=1, cutoff=0.8)
    return match[0] if match else None

# Função para validar datas no formato dd/mm/yyyy
def is_valid_date(date_str):
    if len(date_str) != 10:
        return False
    try:
        day, month, year = map(int, date_str.split('/'))
        return 1 <= day <= 31 and 1 <= month <= 12 and len(str(year)) == 4
    except ValueError:
        return False

# Função para validar formato do Número do Documento
def is_valid_numero_documento(numero):
    if len(numero) != 21:
        return False
    pattern = re.compile(r'\d{2}.\d{2}.\d{5}.\d{7}-\d{1}')
    return bool(pattern.match(numero))

def get_codigo_denominacao_info(line):
    codigos = {
        "8189": "PIS FATURAMENTO 02 PIS FATURAMENTO PJ EM GERAL",
        "2889": "IRPJ LUCRO PRESUMIDO Principal",
        "2372": "CSLL - DEMAIS Principal",
        "2172": "COFINS CONTRIB P/ FIN. SEG. SOCIAL",
        "2009": "IRPJ LUCRO PRESUMIDO",
        "8109": "PIS - FATURAMENTO Principal"
    }

    for codigo, descricao in codigos.items():
        if line.startswith(codigo):
            return codigo, descricao
    return None, None

# Função para processar o texto extraído e gerar o JSON formatado
def process_text_and_generate_json(text):
    if not text.startswith("Receita Federal\nDocumento de Arrecadação\nde Receitas Federais\n"):
        logging.warning("PDF fora do formato de Guia Federal.")
        return None

    try:
        lines = text.split('\n')
        data = {}

        # Extrair o nome do arquivo
        data["Nome do Arquivo"] = pdf_file_name
        
        # Extrair o CNPJ
        cnpj_index = find_similar_term("Razão Social", lines)
        if cnpj_index:
            cnpj_index = lines.index(cnpj_index) + 1
            cnpj = lines[cnpj_index].split()[0]
            data["CNPJ"] = cnpj
        else:
            logging.error("Erro ao encontrar 'Razão Social' para extrair o CNPJ.")
            return None

        # Extrair a Razão Social
        try:
            razao_social_index = cnpj_index
            razao_social = " ".join(lines[razao_social_index].split()[1:])
            data["Razão Social"] = razao_social
        except IndexError:
            logging.error("Erro ao extrair 'Razão Social'.")
            return None

        # Extrair o Período de Apuração
        try:
            periodo_apuracao = next((line for line in lines if is_valid_date(line)), None)
            if periodo_apuracao:
                data["Periodo de Apuração"] = periodo_apuracao
            else:
                logging.error(f"Período de Apuração não encontrado.")
                return None
        except ValueError:
            logging.error("Erro ao encontrar 'Periodo de Apuração'.")
            return None

        # Extrair a Data de Vencimento
        try:
            data_vencimento = next((line for line in lines if is_valid_date(line) and line != data["Periodo de Apuração"]), None)
            if data_vencimento:
                data["Data de Vencimento"] = data_vencimento
            else:
                logging.error(f"Data de Vencimento não encontrada.")
                return None
        except ValueError:
            logging.error("Erro ao encontrar 'Data de Vencimento'.")
            return None

        # Verificar se Período de Apuração é menor que Data de Vencimento
        try:
            periodo_apuracao_date = datetime.strptime(data["Periodo de Apuração"], "%d/%m/%Y")
            data_vencimento_date = datetime.strptime(data["Data de Vencimento"], "%d/%m/%Y")
            if periodo_apuracao_date >= data_vencimento_date:
                logging.error(f"Período de Apuração {data['Periodo de Apuração']} não pode ser maior ou igual à Data de Vencimento {data['Data de Vencimento']}.")
                return None
        except ValueError as e:
            logging.error(f"Erro ao comparar datas: {str(e)}")
            return None

        # Extrair Observações
        try:
            observacoes_index = next((line for line in lines if "Darf emitido pelo Sicalc Web" in line or line.startswith("Nº Recibo Declaração")), None)
            if observacoes_index:
                observacoes_index = lines.index(observacoes_index)
                observacoes = lines[observacoes_index]
                data["Observações"] = observacoes
            else:
                logging.error("Erro ao encontrar 'Observações'.")
                return None
        except ValueError:
            logging.error("Erro ao encontrar 'Observações'.")
            return None

        # Extrair Número do Documento
        try:
            numero_documento = next((line for line in lines if is_valid_numero_documento(line)), None)
            if numero_documento:
                data["Número do Documento"] = numero_documento
            else:
                logging.error(f"Número do Documento inválido: {numero_documento}")
                return None
        except ValueError:
            logging.error("Erro ao encontrar 'Número do Documento'.")
            return None

        # Extrair Valor Total do Documento
        valor_total_documento_term = find_similar_term("Valor Total do Documento", lines) or find_similar_term("Valor Total de Documento", lines)
        if valor_total_documento_term:
            valor_total_documento_index = lines.index(valor_total_documento_term) + 1
            valor_total_documento = lines[valor_total_documento_index]
            data["Valor Total do Documento"] = valor_total_documento
        else:
            logging.error("Erro ao encontrar 'Valor Total do Documento'.")
            return None

        # Extrair Código Denominação e Descrição Cod Denominação
        try:
            codigo_denom_line = next((line for line in lines if any(line.startswith(codigo) for codigo in ["8189", "2889", "2372", "2172", "2009", "8109"])), None)
            if codigo_denom_line:
                codigo_denom, descricao_denom = get_codigo_denominacao_info(codigo_denom_line)
                if codigo_denom:
                    data["Código Denominação"] = codigo_denom
                    data["Descrição Cod Denominação"] = descricao_denom
                else:
                    logging.error("Erro ao extrair 'Código Denominação'.")
                    return None
            else:
                logging.error("Erro ao encontrar 'Código Denominação'.")
                return None
        except ValueError:
            logging.error("Erro ao encontrar 'Código Denominação'.")
            return None

        return data

    except Exception as e:
        logging.error(f"Erro ao processar texto: {str(e)}")
        return None

# Função para salvar textos extraídos em um arquivo JSON
def save_texts_to_json(texts, output_path):
    with open(output_path, 'w', encoding='utf-8') as json_file:
        json.dump(texts, json_file, ensure_ascii=False, indent=4)
    logging.info(f"Textos extraídos salvos em {output_path}")

# Função para apagar arquivos temporários (PDF e imagens)
def delete_temp_files(pdf_path, images):
    if os.path.exists(pdf_path):
        os.remove(pdf_path)
        logging.info(f"Arquivo PDF {pdf_path} apagado.")
    for image in images:
        if os.path.exists(image):
            os.remove(image)
            logging.info(f"Imagem {image} apagada.")

# Listar arquivos na pasta especificada
folder_url = f"/personal/erick_bryan_planning_com_br/Documents/landing_zone/{folder_path}"
pdf_files = list_pdfs(ctx, folder_url)

# Verificar se foram encontrados arquivos PDF
if not pdf_files:
    logging.error(f"Nenhum arquivo PDF encontrado na pasta {folder_path}")
    raise FileNotFoundError(f"Nenhum arquivo PDF encontrado na pasta {folder_path}")

# Lista para armazenar todos os dados extraídos
all_data = []

# Iterar sobre todos os arquivos PDF
for pdf_file_name in pdf_files:
    # Caminho completo do PDF
    pdf_file_path = os.path.join(pdf_dir, pdf_file_name)

    # Construir a URL completa do PDF
    server_relative_url = f"{folder_url}/{pdf_file_name}"

    # Baixar o PDF
    try:
        download_pdf(ctx, server_relative_url, pdf_file_path)
    except ValueError as e:
        logging.error(e)
        continue

    # Converter PDF em imagens
    try:
        images = convert_pdf_to_images(pdf_file_path, images_dir)
    except fitz.FileDataError as e:
        logging.error(f"Erro ao abrir o arquivo PDF: {e}")
        continue

    # Extrair texto das imagens
    texts = extract_text_from_images(images)

    # Processar e adicionar os dados extraídos
    for text in texts:
        processed_data = process_text_and_generate_json(text)
        if processed_data:
            all_data.append(processed_data)

    # Apagar os arquivos temporários
    delete_temp_files(pdf_file_path, images)

# Caminho do arquivo de saída JSON consolidado
output_json_path = os.path.join(output_dir, "consolidated_data.json")

# Salvar os dados extraídos em um único arquivo JSON
save_texts_to_json(all_data, output_json_path)
