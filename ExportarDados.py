import requests
import pandas as pd
from datetime import datetime
import os

# ===== CONFIG =====
tenant_id = "XXXXXXXXXXXXXXXXXX"
client_id = "XXXXXXXXXXXXXXXXXX"
client_secret = "XXXXXXXXXXXXXXXXXXXXXXXXXXX"
dataset_id = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"

caminho = r"C:\Users\raquel.costa\Documentos\Dimensionamento Mensal"

# cria pasta se não existir
os.makedirs(caminho, exist_ok=True)

# ===== DATA ATUAL =====
hoje = datetime.now()
ano = hoje.year
mes = hoje.month

# ===== AUTENTICAÇÃO =====
url_auth = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

data = {
    "grant_type": "client_credentials",
    "client_id": client_id,
    "client_secret": client_secret,
    "scope": "https://analysis.windows.net/powerbi/api/.default"
}

response = requests.post(url_auth, data=data)

if response.status_code != 200:
    raise Exception(f"Erro autenticação: {response.text}")

access_token = response.json()["access_token"]

headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

# ===== QUERY DAX =====
query = f"""
EVALUATE

SELECTCOLUMNS(

    FILTER(

        Geral,

        YEAR(Geral[Data]) = {ano} &&

        MONTH(Geral[Data]) = {mes}

    ),

    "Data", Geral[Data],

    "Intervalo", Geral[Intervalo],

    "Equipes", Geral[Equipes],

    "Volume", Geral[Volume],

    "TMA", Geral[TMA]

)
"""

body = {
    "queries": [{"query": query}]
}

url = f"https://api.powerbi.com/v1.0/myorg/datasets/{dataset_id}/executeQueries"

response = requests.post(url, headers=headers, json=body)

if response.status_code != 200:
    raise Exception(f"Erro na query: {response.text}")

result = response.json()

# ===== TRATAR =====
rows = result['results'][0]['tables'][0]['rows']
df = pd.DataFrame(rows)

if df.empty:
    print("Nenhum dado retornado!")
else:
    print("Quantidade de linhas:", len(df))
    print(df.head())
    # ===== SALVAR =====
    data_str = hoje.strftime("%Y-%m-%d")

    csv_path = os.path.join(caminho, f"Geral_{data_str}.csv")
    excel_path = os.path.join(caminho, f"Geral_{data_str}.xlsx")

    df.to_csv(csv_path, index=False, encoding='utf-8-sig')
    df.to_excel(excel_path, index=False)

    print("Arquivos salvos:")
    print(csv_path)
    print(excel_path)
