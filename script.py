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

        #Buscar arquivos no FTP, txt fora das pastas camunda
        funcoes.mover_arquivos_txt(folder, ftp)

        #Verificação arquivos, retirando tabulações, preparando para integração
        funcoes.verifArquivos()
        
        try:
            retorno = funcoes.integrar(navegador, returns)

        except TimeoutException:
            funcoes.marcacao_cod('Timeout na tentativa. Aguardando 5 minutos antes de tentar novamente.', 'erro')
            navegador.switch_to.default_content()  # Sai do iframe
            time.sleep(300)  # Espera 5 minutos antes da nova tentativa com o mesmo arquivo
            retorno = funcoes.integrar(navegador, returns)


        except Exception as e:
            funcoes.marcacao_cod(f"Erro inesperado: {e}. Aguardando 5 minutos antes de tentar novamente.", 'erro')
            navegador.switch_to.default_content()  # Sai do iframe
            time.sleep(300)  # Espera 5 minutos antes da nova tentativa com o mesmo arquivo
            retorno = funcoes.integrar(navegador, returns)


        funcoes.marcacao_cod(f"Processo concluído com sucesso para a pasta: {folder}", 'titulo')

    except TimeoutException as e:
        funcoes.marcacao_cod(f"Erro de timeout durante o processamento da pasta {folder}: {e}", 'erro')
    
    except Exception as e:
        # Captura qualquer outro erro
        funcoes.marcacao_cod(f"Erro inesperado durante o processamento da pasta {folder}: {e}", 'erro')


funcoes.marcacao_cod("Encerrando conexão FTP", 'titulo')
ftp.close()


if retorno:
    #Criar Excel com resultado
    pathSalvar = os.getenv('PATHSAVE')
    df = pd.DataFrame(retorno, columns=['Arquivo', 'Retorno'])
    df.to_excel(pathSalvar, index=False)

    #enviar e-mail
    funcoes.enviar_email_com_anexo(pathSalvar)