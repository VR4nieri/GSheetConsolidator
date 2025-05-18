import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import time
from gspread_dataframe import get_as_dataframe
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import re

# CONFIGURAÇÕES
FOLDER_ID = 'SUA_PASTA_ID_AQUI'
ARQUIVO_SAIDA = 'arquivo_consolidado.xlsx'
SCOPES = [
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]
CREDS = Credentials.from_service_account_file(
    'CAMINHO_PARA_SEU_ARQUIVO_JSON_CREDENCIAIS.json',
    scopes=SCOPES
)
gc = gspread.authorize(CREDS)

def listar_arquivos_na_pasta(folder_id):
    drive_service = build('drive', 'v3', credentials=CREDS)
    arquivos = []
    page_token = None
    while True:
        resposta = drive_service.files().list(
            q=f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet'",
            spaces='drive',
            fields='nextPageToken, files(id, name)',
            pageToken=page_token
        ).execute()
        arquivos.extend(resposta.get('files', []))
        page_token = resposta.get('nextPageToken', None)
        if page_token is None:
            break
    return arquivos

def upload_para_drive(caminho_arquivo, nome_arquivo, folder_id):
    drive_service = build('drive', 'v3', credentials=CREDS)
    arquivo_metadata = {
        'name': nome_arquivo,
        'parents': [folder_id],
        'mimeType': 'application/vnd.google-apps.spreadsheet'
    }
    media = MediaFileUpload(caminho_arquivo, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    arquivo = drive_service.files().create(
        body=arquivo_metadata,
        media_body=media,
        fields='id, webViewLink'
    ).execute()
    print(f"Arquivo enviado com sucesso! Link: {arquivo['webViewLink']}")

# PROCESSAMENTO
dados_consolidados = []
planilhas = listar_arquivos_na_pasta(FOLDER_ID)
print(f"{len(planilhas)} planilhas encontradas.")

for planilha in planilhas:
    try:
        arquivo = gc.open_by_key(planilha['id'])
        titulo = planilha['name']
        for aba in arquivo.worksheets():
            nome_aba = aba.title
            nome_normalizado = nome_aba.lower()
            if re.search(r'aula\s*', nome_normalizado) and 'exemplo' in nome_normalizado: //personalizar aqui de acordo com o que tem no titulo das abas
                print(f"Lendo aba: {nome_aba} da planilha: {titulo}")
                df = get_as_dataframe(aba, evaluate_formulas=True, dtype=str, header=None)
                df.columns = df.columns.astype(str)
                df.dropna(how='all', inplace=True)
                df.dropna(axis=1, how='all', inplace=True)
                if len(df) < 1:
                    print(f"Aba {nome_aba} em {titulo} tem menos de 2 linhas após limpeza. Pulando.")
                    continue
                df = df.iloc[1:].reset_index(drop=True)
                try:
                    datas = pd.to_datetime(df.iloc[:, 0], errors='coerce')
                    if datas.notna().any():
                        df = df[datas.dt.year == 2025]
                    else:
                        print(f"Nenhuma data válida em {titulo} - {nome_aba}. Ignorando filtro de ano.")
                except Exception as e:
                    print(f"Erro ao filtrar datas na planilha {titulo}, aba {nome_aba}: {e}")
                    continue
                if len(df) > 0:
                    df.insert(0, 'Nome da Aba', nome_aba)
                    df.insert(1, 'Título da Planilha', titulo)
                    dados_consolidados.append(df)
                time.sleep(2)
    except Exception as e:
        print(f"Erro ao acessar {planilha['name']}: {e}")

if dados_consolidados:
    df_final = pd.concat(dados_consolidados, ignore_index=True)
    print(f"\nExemplo dos dados consolidados:")
    print(df_final.head())
    print(f"Total de linhas: {len(df_final)}")
    try:
        df_final.to_excel(ARQUIVO_SAIDA, index=False)
        print(f"Arquivo salvo: {ARQUIVO_SAIDA}")
        upload_para_drive(ARQUIVO_SAIDA, ARQUIVO_SAIDA, FOLDER_ID)
    except Exception as e:
        print(f"Erro ao salvar ou enviar: {e}")
else:
    print("\nNenhuma aba válida foi encontrada para consolidação.")
