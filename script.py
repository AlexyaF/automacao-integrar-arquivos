from selenium import webdriver
import funcoes
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException
from ftplib import FTP
from selenium.webdriver.chrome.service import Service 
import pandas as pd
import os
import time
from dotenv import load_dotenv

#Carregando palavras chaves
load_dotenv()


servico = Service(ChromeDriverManager().install()) # Identifica a versão do navegador atual e vai baixar o Chrome Driver mais recente.
navegador = webdriver.Chrome()
navegador.set_page_load_timeout(10000) 

folders=os.getenv("FOLDERS")
folders_list = folders.strip("[]").split(",") # Removendo colchetes e transformando em lista
folders_list = [folder.strip() for folder in folders_list] # Removendo espaços em branco ao redor de cada item (se houver)

#abrir navegador
funcoes.abrir_driver(navegador)

returns = []

#Para Cada CIA que tenha arquivos quais precisam ser importados
for folder in folders_list:
    funcoes.marcacao_cod(f"INICIANDO {folder}", "titulo")
    ftp = funcoes.conexao_ftp()
    try:
        #Buscar arquivos no FTP
        funcoes.mover_arquivos_processado(folder, ftp)

        #if folder not in ['D003', 'D007', 'D006', 'D005', 'D004']:
        if folder != 'D003':
            #Buscar arquivos no FTP, txt fora das pastas camunda
            funcoes.mover_arquivos_txt(folder, ftp)

        #Verificação arquivos, retirando tabulações, preparando para integração
        funcoes.verifArquivos()
        
        try:
            retorno = funcoes.integrar(navegador, returns)

        except TimeoutException:
            funcoes.marcacao_cod('Timeout na tentativa. ', 'erro')
            navegador.switch_to.default_content()  # Sai do iframe
            funcoes.email_erro("Integração", e)


        except Exception as e:
            funcoes.marcacao_cod(f"Erro inesperado: {e}. ", 'erro')
            navegador.switch_to.default_content()  # Sai do iframe
            funcoes.email_erro("Integração", e)


        funcoes.marcacao_cod(f"Processo concluído com sucesso para a pasta: {folder}", 'titulo')

    except TimeoutException as e:
        funcoes.marcacao_cod(f"Erro de timeout durante o processamento da pasta {folder}: {e}", 'erro')
        funcoes.email_erro("Integração", e)
    
    except Exception as e:
        # Captura qualquer outro erro
        funcoes.marcacao_cod(f"Erro inesperado durante o processamento da pasta {folder}: {e}", 'erro')
        funcoes.email_erro("Integração", e)


funcoes.marcacao_cod("Encerrando conexão FTP", 'titulo')
ftp.close()


if retorno:
    #Criar Excel com resultado
    pathSalvar = os.getenv('PATHSAVE')
    df = pd.DataFrame(retorno, columns=['Arquivo', 'Retorno'])
    df.to_excel(pathSalvar, index=False)

    #enviar e-mail
    funcoes.enviar_email_com_anexo(pathSalvar)