import os.path
import pandas as pd
import plotly.express as px
from dash import Dash, dcc, html, Input, Output
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
SAMPLE_SPREADSHEET_ID = '1LxGAbfb8iQ7FiihLVrUJd7x43iwuYP14MltDhctfKq8'
SHEET_NAME = 'CONTROLE DE ENTRADAS DE MATERIAIS EXTERNOS'
COLUMN_G_RANGE = f'{SHEET_NAME}!G1:G999'
COLUMN_K_RANGE = f'{SHEET_NAME}!K1:K999'

def fetch_data():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            from google_auth_oauthlib.flow import InstalledAppFlow
            flow = InstalledAppFlow.from_client_secrets_file('credenciais.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()
        result_g = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=COLUMN_G_RANGE).execute()
        result_k = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=COLUMN_K_RANGE).execute()

        valores_g = result_g.get('values', [])
        valores_k = result_k.get('values', [])

        data = {'Informacao': [], 'Status': []}
        for valor_g, valor_k in zip(valores_g, valores_k):
            if valor_g and valor_k:
                data['Informacao'].append(valor_g[0])
                data['Status'].append(valor_k[0].strip().upper())

        df = pd.DataFrame(data)
        return df

    except HttpError as err:
        print(f"An error occurred: {err}")
        return pd.DataFrame()

def create_dashboard(df):
    app = Dash(__name__, suppress_callback_exceptions=True)

    df_filtered = df[df['Status'] != 'SITUAÇÃO']

    app.layout = html.Div([
        dcc.Location(id='url', refresh=False),
        dcc.Store(id='selected-status', data=None),
        html.Div(id='page-content')
    ], style={'textAlign': 'center'})

    @app.callback(
        Output('page-content', 'children'),
        Input('url', 'pathname')
    )
    def display_page(pathname):
        if pathname == '/details':
            return details_page()
        else:
            return main_page()

    def main_page():
        status_counts = df_filtered['Status'].value_counts().reset_index()
        status_counts.columns = ['Status', 'Número de itens']

        fig = px.bar(
            status_counts,
            x='Status',
            y='Número de itens',
            color='Status',
            text='Número de itens',
            title='Comparação de itens entregues e aguardando retirada',
            labels={'Número de itens': 'Número de itens', 'Status': 'Status'}
        )

        fig.update_traces(texttemplate='%{text}', textposition='outside', cliponaxis=False)
        fig.update_layout(
            xaxis_title='Status',
            yaxis_title='Número de itens',
            title='Comparação de status de cada item',
            uniformtext_mode='hide',
            uniformtext_minsize=14
        )

        return html.Div([
            html.H1("Dashboard de controle de materiais", style={'textAlign': 'center'}),
            dcc.Graph(
                id='status-bar-chart',
                figure=fig
            ),
            dcc.Link('Ver detalhes dos itens', href='/details')
        ], style={'textAlign': 'center'})

    def details_page():
        return html.Div([
            html.H1("Detalhes dos itens por status"),
            dcc.Link('Voltar para o dashboard', href='/'),
            html.Div(id='item-details')
        ], style={'textAlign': 'center'})

    @app.callback(
        Output('selected-status', 'data'),
        Input('status-bar-chart', 'clickData')
    )
    def update_selected_status(clickData):
        if clickData:
            selected_status = clickData['points'][0]['x']
            return {'status': selected_status}
        return None

    @app.callback(
        Output('item-details', 'children'),
        Input('selected-status', 'data')
    )
    def display_item_details(selected_status):
        filtered_df = df_filtered

        if selected_status:
            status = selected_status['status']
            filtered_df = filtered_df[filtered_df['Status'] == status]

        grouped = filtered_df.groupby('Status')

        details = []
        for status, group in grouped:
            details.append(html.H2(f"Detalhes dos itens com status '{status}'"))
            for index, row in group.iterrows():
                details.append(html.Div(f"- {row['Informacao']}"))

        if not details:
            details.append(html.Div("Nenhum item encontrado."))

        return details

    app.run_server(debug=True)

def main():
    df = fetch_data()
    if not df.empty:
        create_dashboard(df)
    else:
        print("Nenhum dado disponível para criar o dashboard.")

if __name__ == "__main__":
    main()