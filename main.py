from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    sheet = service.spreadsheets()
    alunos = sheet.values().get(spreadsheetId='1Us3Inu3gQvVz6XFjownctGIdLW3bp1hBOcq1sI9-dgA',
                                range='engenharia_de_software!B4:B27').execute()
    faltas = sheet.values().get(spreadsheetId='1Us3Inu3gQvVz6XFjownctGIdLW3bp1hBOcq1sI9-dgA',
                                range='engenharia_de_software!C4:C27').execute()
    p1 = sheet.values().get(spreadsheetId='1Us3Inu3gQvVz6XFjownctGIdLW3bp1hBOcq1sI9-dgA',
                                range='engenharia_de_software!D4:D27').execute()
    p2 = sheet.values().get(spreadsheetId='1Us3Inu3gQvVz6XFjownctGIdLW3bp1hBOcq1sI9-dgA',
                                range='engenharia_de_software!E4:E27').execute()
    p3 = sheet.values().get(spreadsheetId='1Us3Inu3gQvVz6XFjownctGIdLW3bp1hBOcq1sI9-dgA',
                                range='engenharia_de_software!F4:F27').execute()
    total_de_aulas = sheet.values().get(spreadsheetId='1Us3Inu3gQvVz6XFjownctGIdLW3bp1hBOcq1sI9-dgA',
                            range='engenharia_de_software!A2:H2').execute()

    values_alunos = alunos.get('values', [])
    values_faltas = faltas.get('values', [])
    values_nota_p1 = p1.get('values', [])
    values_nota_p2 = p2.get('values', [])
    values_nota_p3 = p3.get('values', [])
    values_total_aulas = int(total_de_aulas.get('values', [])[0][0][27:])

    values_situacao = []
    values_nota_aprovacao_final = []

    for i in range(len(values_alunos)):
        nr_falta = int(values_faltas[i][0])
        lista_situacao = []
        lista_nota_aprovacao_final = []
        if nr_falta/values_total_aulas > 0.25:
            lista_situacao.append("Reprovado por falta")
            values_situacao.append(lista_situacao)
            lista_nota_aprovacao_final.append('0')
            values_nota_aprovacao_final.append(lista_nota_aprovacao_final)
        else:
            nota_p1 = int(values_nota_p1[i][0])
            nota_p2 = int(values_nota_p2[i][0])
            nota_p3 = int(values_nota_p3[i][0])
            media = float((nota_p1 + nota_p2 + nota_p3)/3)
            if media > 70:
                lista_situacao.append("Aprovado")
                values_situacao.append(lista_situacao)
                lista_nota_aprovacao_final.append('0')
                values_nota_aprovacao_final.append(lista_nota_aprovacao_final)
            elif media < 50:
                lista_situacao.append("Reprovado por Nota")
                values_situacao.append(lista_situacao)
                lista_nota_aprovacao_final.append('0')
                values_nota_aprovacao_final.append(lista_nota_aprovacao_final)
            else:
                lista_situacao.append("Exame Final")
                values_situacao.append(lista_situacao)
                nota_aprovacao_final = round(100 - media, 2)
                lista_nota_aprovacao_final.append(nota_aprovacao_final)
                values_nota_aprovacao_final.append(lista_nota_aprovacao_final)

    sheet.values().update(spreadsheetId='1Us3Inu3gQvVz6XFjownctGIdLW3bp1hBOcq1sI9-dgA',
                                range='engenharia_de_software!G4:G27', valueInputOption="RAW",
                                   body={"values": values_situacao}).execute()

    sheet.values().update(spreadsheetId='1Us3Inu3gQvVz6XFjownctGIdLW3bp1hBOcq1sI9-dgA',
                                   range='engenharia_de_software!H4:H27', valueInputOption="RAW",
                                   body={"values": values_nota_aprovacao_final}).execute()

if __name__ == '__main__':
    main()
