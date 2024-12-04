import os
import win32com.client as win32
from datetime import datetime
from ftplib import FTP
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from time import sleep
from dotenv import load_dotenv
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
    data = datetime.today()
    data_atual = data.strftime("%d/%m/%Y")
    destinatario = os.getenv("DESTINATARIO")
    copia = os.getenv("COPIA")
    if copia:
        # Removendo os colchetes e convertendo para lista
        copia = copia.strip("[]").replace(" ", "").split(",")
    else:
        copia = []

    assunto = f"****TESTE**** Integração CPFL - {data_atual}"
    corpo  = "Segue em anexo casos que passaram pelo script de integração."
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
  


def mover_arquivos_processado(folder, ftp):
    pathLocal = os.getenv('PATHFOLDER_INTEGRADOS')
    integrar = os.getenv('PATHFOLDER')
    integrar_files= os.listdir(integrar)
    integrados= os.listdir(pathLocal)

    try:
        # Navegando e baixando arquivos das pastas remotas
        print(f"Navegando para: {folder}")
        ftp.cwd(f'/cobrconta/GMP/{folder}/0713464419/RetornoGMP/Processado')
        files = ftp.nlst()

        if files:
            # Calculando arquivos que ainda não foram integrados
            diff_files = list(set(files) - set(integrados))
            #Dos arquivos não integrados quais já estão na pasta de integrar
            final_files = list(set(diff_files) - set(integrar_files))
            #Mover somente arquivos que nao estão na pasta integrar
            for files in final_files:
                # Construindo o caminho completo para salvar o arquivo localmente
                allpath = os.path.join(integrar, files)
                
                # Baixando o arquivo
                with open(allpath, "wb") as local_file:
                    ftp.retrbinary(f"RETR {files}", local_file.write)
                print(f"Arquivo '{files}' movido para '{allpath}'")
        else: 
            print("Sem arquivos")

    except Exception as e:
        print(f"Erro ao conectar ao FTP: {e}")



def mover_integrados(file):
    origem = os.path.join(os.getenv('PATHFOLDER'), file)
    destino = os.path.join(os.getenv('PATHFOLDER_INTEGRADOS'), file)

    shutil.move(origem, destino)



#Abrir navegador/ site
def abrir_driver(navegador):
    navegador.get(os.getenv('SITE'))
    #Buscar o elemento e escrever o usuário
    navegador.find_element('xpath', '//*[@id="formulario"]/div[1]/input').send_keys(os.getenv('USER'))
    navegador.find_element('xpath', '//*[@id="formulario"]/div[2]/input').send_keys(os.getenv('PASSWORD'))

    #Entrar
    navegador.find_element('xpath', '//*[@id="btentrar"]').click()

    #Buscar tela de integracao
    try:
        navegador.find_element('xpath', '//*[@id="field5"]').send_keys("Retorno CPFL")
        sleep(5)

        # Aguarda o elemento ficar visível
        wait = WebDriverWait(navegador, 10000)  # Substitua 'navegador' pelo nome do driver
        element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="field5-suggestions"]//span[text()="RETORNO CPFL"]')))
        element.click()
    except:
        print("ERRO!")



#Buscar arquivos na pasta de integrar
def integrar(navegador):
    pathFolder = os.getenv('PATHFOLDER')
    files = os.listdir(pathFolder)

    #Salvar retorno da tela
    returns = []

    for file in files:
        allpath = os.path.join(pathFolder, file)
        
        #Upload arquivo
        wait = WebDriverWait(navegador, 10000) # Aguarda até o elemento estar visível
        navegador.switch_to.frame("cont") # mudando para o iframe
        input_element = wait.until(EC.visibility_of_element_located((By.ID, 'Upload1'))) #aguardo até a visibilidade do elemento e identifico 
        input_element.send_keys(allpath) #envio o arquivo

        #Integrar click
        navegador.find_element('xpath','//*[@id="Btncef"]').click()
    

        try:
            retorno_element = WebDriverWait(navegador, 10000).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="Label9"]'))
            )
            texto = retorno_element.text
            mover_integrados(file) #passar arquivo para integrados

        except:
            print("Tempo limite excedido ao esperar pela integração.")
            texto = "Erro: tempo limite excedido."
            
        navegador.switch_to.default_content() #saindo do iframe 
        joinArquivoRetorno = [file, texto]
        returns.append(joinArquivoRetorno)
    return returns