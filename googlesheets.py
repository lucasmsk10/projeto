import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# The ID of a sample document.
SAMPLE_SPREADSHEET_ID = '1LxGAbfb8iQ7FiihLVrUJd7x43iwuYP14MltDhctfKq8'
SHEET_NAME = 'CONTROLE DE ENTRADAS DE MATERIAIS EXTERNOS'
COLUMN_G_RANGE = f'{SHEET_NAME}!G1:G999'  # Coluna G, limite de 1000 linhas
COLUMN_K_RANGE = f'{SHEET_NAME}!K1:K999'  # Coluna K, limite de 1000 linhas

def main():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credenciais.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Faz a requisição para obter os dados das colunas G e K
        sheet = service.spreadsheets()
        result_g = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=COLUMN_G_RANGE).execute()
        result_k = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=COLUMN_K_RANGE).execute()

        valores_g = result_g.get('values', [])
        valores_k = result_k.get('values', [])

        # Itera sobre os valores e imprime os correspondentes onde K == "ENTREGUE"
        for valor_g, valor_k in zip(valores_g, valores_k):
            if valor_k and valor_k[0].strip().upper() == 'ENTREGUE':
                #print(f"Coluna G: {valor_g[0]} | Coluna K: {valor_k[0]}")
                print(f"{valor_g[0]} | {valor_k[0]}")

    except HttpError as err:
        print(f"An error occurred: {err}")

if __name__ == "__main__":
    main()