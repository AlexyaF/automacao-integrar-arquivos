import os
import re
import win32com.client as win32
from datetime import datetime
from dotenv import load_dotenv


load_dotenv()

def verifArquivos():
    path = os.getenv("PATHFOLDER")
    files = os.listdir(path)

    for file in files:
        arquivo = os.path.join(path, file)
        # Abrindo o arquivo para procurar tabulações
        with open(arquivo, 'r', encoding='utf-8') as arq:
            linhas_corrigidas = [linha.replace('\t', '') for linha in arq]

        # Salvando o arquivo corrigido
        with open(arquivo, 'w', encoding='utf-8') as arq:
            arq.writelines(linhas_corrigidas)




def enviar_email_com_anexo(anexo=None):
    data_atual = datetime.today()
    destinatario = os.getenv("DESTINATARIO")
    copia = os.getenv("COPIA")
    if copia:
        # Removendo os colchetes e convertendo para lista
        copia = copia.strip("[]").replace(" ", "").split(",")
    else:
        copia = []

    assunto = f"****TESTE**** Integração CPFL - {data_atual}"
    corpo  = "Segue em anexo casos integrados no dia de hoje."
    print(f"Destinatário: {destinatario}")
    print(f"Cópia: {copia}")
    try:
        outlook = win32.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)
        email.To = destinatario
        email.CC = ";".join(copia) if copia else ""
        email.Subject = assunto
        email.Body = corpo
        if anexo : email.Attachments.Add(anexo)
        email.Send()
        print(f"Email enviado para {destinatario} ")
    except Exception as e:
        print(f"Erro ao enviar o email: {e}")
  