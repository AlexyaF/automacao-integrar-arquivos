import os
import win32com.client as win32
from datetime import datetime
from ftplib import FTP
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
from time import sleep
from dotenv import load_dotenv
from dotenv import load_dotenv


load_dotenv()


def conexao_ftp():
    #Configurações FTP
    config ={
        'host':os.getenv('HOST_FTP'),
        'user':os.getenv('USER_FTP'),
        'password':os.getenv('PASSWORD_FTP')
    }
    # Conectando ao servidor FTP
    try:
        ftp = FTP(config['host'])
        ftp.login(user=config['user'], passwd=config['password'])
        print("===== Conexão estabelecida com FTP ===== ")
    except Exception as e:
        print(f"Erro ao conectar ao FTP: {e}")
    
    return ftp



def verifArquivos():
    print('===== VERIFICANDO ARQUIVO =====')
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
    print('===== ENVIANDO E-MAIL =====')
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
    corpo  = f"Segue em anexo casos que passaram pelo script de integração."
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
    print('===== CONFERINDO ARQUIVOS INTEGRADOS NO FTP  =====')
    pathLocal = os.getenv('PATHFOLDER_INTEGRADOS')
    integrar = os.getenv('PATHFOLDER')
    integrar_files= os.listdir(integrar)
    integrados= os.listdir(pathLocal)

    folderProcess = ['Processado', 'Processados']
    for folders in folderProcess:
        try:
            # Navegando e baixando arquivos das pastas remotas
            print(f"Navegando para: {folder} em '{folders}'")
            ftp.cwd(f'/cobrconta/GMP/{folder}/0713464419/RetornoGMP/{folders}')
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


def mover_integrados(file, allpath):
    print('===== MOVENDO INTEGRADOS =====')

    if 'D003' in file:
        folder = 'D003 - Santa Cruz'
    elif 'D004' in file:
        folder = 'D004 - Leste Paulista'
    elif 'D005' in file:
        folder = 'D005 - Sul Paulista'
    elif 'D006' in file:
        folder = 'D006 - Jaguari'
    elif 'D007' in file:
        folder = 'D007 - Mococa'
    elif 'D008' in file:
        folder = 'D008 - RGE'
    elif 'D009' in file:
        folder = 'D009 - RGE SUL'
    elif 'CPFL' in file:
        folder = 'CPFL'
    elif 'PIRA' in file:
        folder = 'PIRA'
    else:
        print("Sem pasta correspondente")


    origem = allpath
    destino = os.path.join(os.getenv('PATHSERVER'), folder,'Integrados')
    destino_base_comparacao = os.getenv('PATHFOLDER_INTEGRADOS')

    #Copiar arquivo para pasta 'Integrados', pasta base da comparação ao baixar arquivos do ftp
    shutil.copy(origem, destino_base_comparacao)
    #Mover arquivos para suas pastas pertencentes
    shutil.move(origem, destino)
    print(f'Arquivo: "{file}", movido de "{origem}", para "{destino}"')



#Abrir navegador/ site
def abrir_driver(navegador):
    print('===== ABRINDO SITE =====')
    navegador.get(os.getenv('SITE'))

    #Buscar o elemento e passar o usuário
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



def mover_arquivos_txt(folder, ftp):
    print("===== BUSCANDO ARQUIVOS NAO INTEGRADOS FTP (TXT'S) =====")
    pathLocal = os.getenv('PATHFOLDER_INTEGRADOS')
    integrar = os.getenv('PATHFOLDER')
    integrar_files= os.listdir(integrar)
    integrados= os.listdir(pathLocal)
    
    try:
        # Configurando a codificação correta
        ftp.encoding = 'latin-1'  

        # Navega para a pasta do FTP
        ftp.cwd(f"/cobrconta/GMP/{folder}/0713464419/RetornoGMP")
    except Exception as e:
        print(f"Erro ao conectar ao FTP: {e}")
        return  # Interrompe a execução em caso de erro

    try:
        # Lista os arquivos na pasta
        files = ftp.nlst()

        # Filtra apenas os arquivos .txt
        arquivos_txt = [file for file in files if file.endswith('.txt')]

        if arquivos_txt:
            diff_integrados = list(set(arquivos_txt) - set(integrados))
            diff_integrar = list(set(diff_integrados) - set(integrar_files))
            if diff_integrar:
                for file in diff_integrar:
                    allPath = os.path.join(integrar, file)

                    #Baixando o arquivo na pasta integrar
                    with open(allPath, "wb") as local_file:
                        ftp.retrbinary(f"RETR {file}", local_file.write)
                        print(f"Arquivo '{file}' movido para '{allPath}'")

                    #Mover dentro do FTP
                    ftp.cwd(f"/cobrconta/GMP/{folder}/0713464419/RetornoGMP/Processados")
                    # Fazer upload do arquivo para a pasta de destino
                    with open(allPath, "rb") as local_file:
                        ftp.storbinary(f"STOR {file}", local_file)
                    print(f"Arquivo '{file}' enviado para '{ftp.pwd()}'")

                    ftp.cwd(f"/cobrconta/GMP/{folder}/0713464419/RetornoGMP")  # Retorna ao diretório original
                    ftp.delete(f"/cobrconta/GMP/{folder}/0713464419/RetornoGMP/{file}") #Apaga
                    print(f"Arquivo '{file}' deletado de '{ftp.pwd()}'")

    except UnicodeDecodeError as e:
        print(f"Erro de decodificação: {e}")
    except Exception as e:
        print(f"Erro ao listar ou processar arquivos: {e}")



#Buscar arquivos na pasta de integrar
def integrar(navegador, returns):
    # Logando tempo de execução
    start_time = time.time()

    print('===== INICIANDO INTEGRAÇÃO =====')
    pathFolder = os.getenv('PATHFOLDER')
    files = os.listdir(pathFolder)

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
            file_start_time = time.time()
            retorno_element = WebDriverWait(navegador, 10000).until(
                EC.visibility_of_element_located((By.XPATH, '//*[@id="Label9"]'))
            )
            texto = retorno_element.text
            elapsed_time = time.time() - file_start_time  # Fim do tempo
            print(f"===== CONCLUÍDO PARA: {file}. Tempo: {elapsed_time:.2f} segundos =====")
            try:
                mover_integrados(file, allpath)
            except Exception as e:
                print(f"Erro ao mover arquivo {file}: {e}")

        except TimeoutException:
            # Captura exclusivamente erros de tempo limite
            print("Tempo limite excedido ao esperar pela integração.")
            texto = "Erro: tempo limite excedido."
        except Exception as e:
            # Captura erros gerais
            print(f"Erro inesperado ao processar o arquivo {file}: {e}")
            texto = f"Erro inesperado: {e}"
        finally:
            navegador.switch_to.default_content()  # Sai do iframe   

        joinArquivoRetorno = [file, texto]
        returns.append(joinArquivoRetorno)
    print(f"===== PROCESSAMENTO TOTAL FINALIZADO. Tempo total: {time.time() - start_time:.2f} segundos =====")
    return returns





