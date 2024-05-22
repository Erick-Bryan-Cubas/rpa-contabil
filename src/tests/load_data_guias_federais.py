from google.cloud import bigquery
from google.oauth2 import service_account
import json
import re

# Configurações
project_id = "bi-planning-367317"
dataset_id = "BI_AUDITORIA_SPED"
table_id = "bi_guias_impostos_federais"
json_file_path = "data/output/consolidated_data.json"
credentials_path = "keys/bi-planning.json"

# Carregar credenciais
credentials = service_account.Credentials.from_service_account_file(credentials_path)

# Inicializar o cliente do BigQuery
client = bigquery.Client(credentials=credentials, project=project_id)

# Referência da tabela
table_ref = client.dataset(dataset_id).table(table_id)

# Função para limpar os nomes dos campos
def clean_field_name(field_name):
    # Remover caracteres não permitidos
    field_name = re.sub(r'[^\w]', '_', field_name)
    return field_name

# Carregar dados do arquivo JSON
with open(json_file_path, 'r', encoding='utf-8') as file:
    json_data = json.load(file)

# Limpar os nomes dos campos
cleaned_data = []
for record in json_data:
    cleaned_record = {clean_field_name(key): value for key, value in record.items()}
    cleaned_data.append(cleaned_record)

# Preparar a configuração do carregamento
job_config = bigquery.LoadJobConfig(
    source_format=bigquery.SourceFormat.NEWLINE_DELIMITED_JSON,
    autodetect=True,
    write_disposition=bigquery.WriteDisposition.WRITE_TRUNCATE,
)

# Carregar dados para o BigQuery
load_job = client.load_table_from_json(
    cleaned_data, table_ref, job_config=job_config
)

# Esperar até o job completar
load_job.result()

print(f"Dados carregados para {dataset_id}.{table_id}")
