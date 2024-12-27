from selenium import webdriver
import funcoes
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException
from ftplib import FTP
from selenium.webdriver.chrome.service import Service 
import pandas as pd
import os

from dotenv import load_dotenv
import time

#Carregando palavras chaves
load_dotenv()

returns = []

servico = Service(ChromeDriverManager().install()) # Identifica a versão do navegador atual e vai baixar o Chrome Driver mais recente.
navegador = webdriver.Chrome()
navegador.set_page_load_timeout(10000)

funcoes.verifArquivos()
funcoes.abrir_driver(navegador)

try:
    retorno = funcoes.integrar(navegador, returns)

except TimeoutException as e:
    funcoes.marcacao_cod('Timeout na tentativa.', 'erro')
    navegador.switch_to.default_content()  # Sai do iframe
    funcoes.email_erro("Integração", e)


except Exception as e:
    funcoes.marcacao_cod(f"Erro inesperado: {e}.", 'erro')
    navegador.switch_to.default_content()  # Sai do iframe
    funcoes.email_erro("Integração", e)


#Criar Excel com resultado
pathSalvar = os.getenv('PATHSAVE')
df = pd.DataFrame(retorno, columns=['Arquivo', 'Retorno'])
df.to_excel(pathSalvar, index=False)

#enviar e-mail
funcoes.enviar_email_com_anexo(pathSalvar)