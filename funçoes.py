import os
import re
import win32com.client as win32
from datetime import datetime

def verifArquivos():
    path = r"\\10.44.250.4\M-Energia\Colaboradores\Alexya Silva\CPFL"
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
    destinatario = "alexya.fortunato@crefaz.com.br"
    # copia = "nasly.carmo@crefaz.com.br"
    assunto = f"Integração CPFL - {data_atual}"
    corpo  = "Segue em anexo casos integrados no dia de hoje."
    try:
        outlook = win32.Dispatch('Outlook.Application')
        email = outlook.CreateItem(0)
        email.To = destinatario
        # email.CC = ";".join(copia)
        email.Subject = assunto
        email.Body = corpo
        if anexo : email.Attachments.Add(anexo)
        email.Send()
        print(f"Email enviado para {destinatario} ")
    except Exception as e:
        print(f"Erro ao enviar o email: {e}")
  