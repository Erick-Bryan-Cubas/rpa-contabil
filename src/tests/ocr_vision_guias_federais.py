import os
import logging
import fitz  # PyMuPDF
import base64
import json
import requests
from dotenv import load_dotenv
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Carregar variáveis de ambiente do arquivo .env
load_dotenv('envs/.env')

# Variáveis de ambiente
google_credentials_path = os.getenv('GOOGLE_CLOUD_CREDENTIALS')
caminho_landing_zone = os.getenv('caminho_landing_zone')
pdf_file_name = "Loteamento Alto do Cruzeiro - Darf de COFINS 04.2024 (18 lotes).pdf"
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

# Caminho completo do arquivo PDF
pdf_file_path = os.path.join(pdf_dir, pdf_file_name)

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

# Função para processar o texto extraído e gerar o JSON formatado
def process_text_and_generate_json(text):
    lines = text.split('\n')
    data = {}

    # Extrair o CNPJ
    cnpj_index = lines.index("Razão Social") + 1
    data["CNPJ"] = lines[cnpj_index].split()[0]

    # Extrair a Razão Social
    razao_social_index = cnpj_index
    data["Razão Social"] = " ".join(lines[razao_social_index].split()[1:])

    # Extrair o Período de Apuração
    inicio_periodo_apuracao = text.index("\nData de Vencimento\n") + len("\nData de Vencimento\n")
    fim_periodo_apuracao = inicio_periodo_apuracao + 10
    data["Periodo de Apuração"] = text[inicio_periodo_apuracao:fim_periodo_apuracao]

    # Extrair a Data de Vencimento
    inicio_data_vencimento = fim_periodo_apuracao + 1
    fim_data_vencimento = text.index("\nObservações\n")
    data["Data de Vencimento"] = text[inicio_data_vencimento:fim_data_vencimento]

    # Extrair Observações
    observacoes_index = lines.index("Observações") + 1
    data["Observações"] = lines[observacoes_index]

    # Extrair Número do Documento
    numero_documento_index = lines.index("Número do Documento") + 1
    data["Número do Documento"] = lines[numero_documento_index]

    # Extrair Valor Total do Documento
    valor_total_documento_index = lines.index("Valor Total do Documento") + 1
    data["Valor Total do Documento"] = lines[valor_total_documento_index]

    # Extrair Código Denominação
    codigo_denominacao_index = lines.index("Código Denominação") + 1
    data["Código Denominação"] = " ".join(lines[codigo_denominacao_index:codigo_denominacao_index + 3])

    return [data]

# Função para salvar textos extraídos em um arquivo JSON
def save_texts_to_json(texts, output_path):
    with open(output_path, 'w', encoding='utf-8') as json_file:
        json.dump(texts, json_file, ensure_ascii=False, indent=4)
    logging.info(f"Textos extraídos salvos em {output_path}")

# Função para salvar JSON formatado em um arquivo separado
def save_formatted_json(text, output_path):
    processed_text = process_text_and_generate_json(text)
    with open(output_path, 'w', encoding='utf-8') as json_file:
        json.dump(processed_text, json_file, ensure_ascii=False, indent=4)
    logging.info(f"JSON formatado salvo em {output_path}")

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

# Verificar se o arquivo desejado está na lista
logging.info(f"Arquivos encontrados: {pdf_files}")
if pdf_file_name not in pdf_files:
    logging.error(f"Arquivo {pdf_file_name} não encontrado na pasta {folder_path}")
    raise FileNotFoundError(f"Arquivo {pdf_file_name} não encontrado na pasta {folder_path}")

# Construir a URL completa do PDF
server_relative_url = f"{folder_url}/{pdf_file_name}"

# Baixar o PDF
try:
    download_pdf(ctx, server_relative_url, pdf_file_path)
except ValueError as e:
    logging.error(e)
    raise

# Converter PDF em imagens
try:
    images = convert_pdf_to_images(pdf_file_path, images_dir)
except fitz.FileDataError as e:
    logging.error(f"Erro ao abrir o arquivo PDF: {e}")
    raise Exception(f"Erro ao abrir o arquivo PDF: {e}")

# Extrair texto das imagens
texts = extract_text_from_images(images)

# Caminho do arquivo de saída JSON
output_json_path = os.path.join(output_dir, "extracted_text.json")
formatted_json_path = os.path.join(output_dir, "formatted_data.json")

# Salvar os textos extraídos em um arquivo JSON
save_texts_to_json(texts, output_json_path)

# Salvar o JSON formatado em um arquivo separado
save_formatted_json(texts[0], formatted_json_path)

# Apagar os arquivos temporários
delete_temp_files(pdf_file_path, images)
