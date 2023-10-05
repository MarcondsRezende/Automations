# --------------------------------------------- Bibliotecas necessárias. ---------------------------------------------
import PySimpleGUI as sg
import os
import pandas as pd
from smtplib import SMTP_SSL, SMTP_SSL_PORT
from email.mime.multipart import MIMEMultipart, MIMEBase
from email.mime.text import MIMEText
from email import encoders

# --------------------------------------------- Criando visualização no PySimpleGui ---------------------------------------------
sg.theme('Dark2')

fonte_princial = ('Ubuntu', 12)
fonte_secundaria = ('Ubuntu', 10)
fonte_terciaria = ('Ubuntu', 9)

login_senha = [
    [sg.Text("Email "), sg.Input(key='email', size=(40, 2))],
    [sg.Text("Senha"), sg.InputText(
        key='senha', password_char='*', size=(40, 2))],
    [sg.Button("Login")]
]

planilha = [
    [sg.Text("Selecione a planilha:", font=fonte_princial)],
    [sg.InputText(key="file_path"),
     sg.FileBrowse("Procurar", font=fonte_secundaria, initial_folder=os.getcwd())],
    [sg.Button("Obter Nomes das Colunas",font=fonte_secundaria)],
    [sg.Multiline(size=(35, 5), key="output", disabled=True, font=fonte_terciaria)]
]

pergunta_anexo = [
    [sg.Text("Tem anexo?", font=fonte_princial)],
    [sg.Radio('Sim', 'group 1',
              key='anexo', enable_events=True)],
    [sg.Radio('Não', 'group 1',
              key='anexo', enable_events=True)],
]

titulo = [
    [sg.Text("Titulo do email", font=fonte_princial),
     sg.InputText(key="titulo")],
]


mensagem = [
    [sg.Text("Corpo do email", font=fonte_secundaria)],
    [sg.Multiline(key="body", size=(60, 10))],
    [sg.Button("Enviar"), sg.Button("Cancelar")]
]

# assinatura = [
#     [sg.Text("Assinatura", font=fonte_secundaria)],
#     [sg.Multiline(key="body", size=(60, 5))],
#     [sg.Button("Enviar"), sg.Button("Cancelar")]
# ]

layout = [
    [sg.Column(login_senha, vertical_alignment='center',element_justification='center')],
    [sg.HorizontalSeparator()],
    [sg.Column(planilha, element_justification='left'), 
     sg.VerticalSeparator(), 
     sg.Column(pergunta_anexo, vertical_alignment='left', element_justification='left')],
    [sg.HorizontalSeparator()],
    [sg.Column(titulo, vertical_alignment='center',element_justification='center')],
    [sg.HorizontalSeparator()],
    [sg.Column(mensagem, element_justification='left')],
]


window = sg.Window('Automação de Envios', layout)

# --------------------------------------------- Código ---------------------------------------------


# Função para identificar valores ausentes.
def isNaN(value):
    try:
        import math
        return math.isnan(float(value))
    except:
        return False
    

def Novos(df, email, senha, titulo, mensagem, variaveis_adicionais=None):
    # Criando e Puxando as variaveis para a ferramenta.
    for i, message in enumerate(df["email"]):
        nome = df.loc[i, "nome"]
        anexo = df.loc[i, "anexo"]

        # Verifica se variaveis_adicionais é fornecido e se a chave está presente no DataFrame
        if variaveis_adicionais is not None:
            for chave, valor in variaveis_adicionais.items():
                if chave in df.columns:
                    variavel = df.loc[i, chave]
                    # Substitui a variável no texto de saída
                    output_text = output_text.replace("{" + chave + "}", str(variavel))

        with open("LogEnviados.txt", "a") as arquivo:
            email_enviar = df.loc[i, "EMAIL"]
            if not isNaN(email_enviar):
                email_message = MIMEMultipart()
                email_message.add_header('To', email_enviar)
                email_message.add_header('From', email)
                email_message.add_header('Subject', titulo)
                email_message.add_header('X-Priority', '1')


# --------------------------------------------- PARTE FINAL DO PROGRAMA ---------------------------------------------
while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break

    if event == "Obter Nomes das Colunas":
        file_path = values["file_path"]
        try:
            df = pd.read_excel(file_path)
            colunas = df.columns
            output_text = "\n".join(colunas)
            window["output"].update(output_text)
        except Exception as e:
            window["output"].update(f"Erro ao abrir o arquivo: {str(e)}")


window.close()