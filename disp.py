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
    

def emails(df, email, senha, titulo, mensagem, variaveis_adicionais=None):
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
            if not pd.isna(email_enviar):
                email_message = MIMEMultipart()
                email_message.add_header('To', email_enviar)
                email_message.add_header('From', email)
                email_message.add_header('Subject', titulo)
                email_message.add_header('X-Priority', '1')

                #html_part = MIMEText(mensage, 'html') #Mudar

                #email_message.attach(html_part)

                smtp_server = SMTP_SSL(
                    'smtp-mail.outlook.com', port=587)   # Configuração do servidor SMTP do Outlook
                smtp_server.starttls()   # Use TLS para criptografar a conexão
                #smtp_server.set_debuglevel(1)         #linha para depuração
                smtp_server.login(email, senha)
                smtp_server.sendmail(email, email_enviar,
                                     email_message.as_string())   # Envie o e-mail
                smtp_server.quit()   # Feche a conexão SMTP


# --------------------------------------------- PARTE FINAL DO PROGRAMA ---------------------------------------------
while True:
    eventos, database = window.read()
    if eventos == sg.WINDOW_CLOSED:
        break
    if eventos == 'Enviar':
        if len(database["-FILE_PATH-"]) > 0:
            if len(database["titulo"]) > 0:
                #if database["anexo"]:
                if database["Novos"]:
                    email = database["email"]
                    senha = database["senha"]
                    titulo = database["titulo"]
                    df = pd.read_excel(database["-FILE_PATH-"])

                    emails(df, email, senha, titulo, mensagem, variaveis_adicionais=None)
                    sg.popup("Emails enviados com sucesso!",
                             "Dexei um arquivo chamado LogEnviados junto comigo!", "Da uma verificadinha la!")

            else:
                sg.popup("Coloque um titulo para enviar o email")
        else:
            sg.popup("Preciso saber qual planilha tenho que ler os dados!")