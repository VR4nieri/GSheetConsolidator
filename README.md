# GSheetConsolidator

É um script em Python para **consolidar automaticamente dados de várias planilhas Google**, filtrando por abas específicas (ex: "aula", "exemplo") e salvando os resultados em um único arquivo `.xlsx`, que é enviado automaticamente ao Google Drive.

## Funcionalidades
- Conecta-se ao Google Drive e Google Sheets via API
- Filtra e coleta dados apenas das abas relevantes
- Aplica limpeza de dados e filtragem por ano (ex: 2025)
- Concatena tudo em um único DataFrame
- Salva como Excel (.xlsx)
- Envia o arquivo final para o Drive automaticamente

## Tecnologias utilizadas
- Python
- gspread
- pandas
- Google API Client
- Google Drive API
- Google Sheets API

## Pré-requisitos
- Conta de serviço com permissões de acesso ao Drive e Planilhas
- APIs ativadas no [Google Cloud Console](https://console.cloud.google.com/):
  - Google Drive API
  - Google Sheets API

## Estrutura esperada
- Pasta no Google Drive contendo planilhas do tipo Google Sheets
- Cada planilha deve conter abas com nomes contendo "aula" e "mini"
- Primeira coluna com datas para filtragem por ano

## Como usar

1. Crie e baixe o JSON da conta de serviço
2. Compartilhe a pasta no Drive com o e-mail da conta de serviço
3. Atualize as variáveis:
   - `FOLDER_ID` → ID da pasta do Drive
   - `ARQUIVO_SAIDA` → Nome desejado para o arquivo final
   - Caminho para o JSON na variável `CREDS`

4. Instale os pacotes:
```bash
pip install gspread pandas google-api-python-client gspread-dataframe
