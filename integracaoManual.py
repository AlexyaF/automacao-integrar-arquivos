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

funcoes.abrir_driver(navegador)

try:
    retorno = funcoes.integrar(navegador, returns)

except TimeoutException:
    funcoes.marcacao_cod('Timeout na tentativa. Aguardando 5 minutos antes de tentar novamente.', 'erro')
    navegador.switch_to.default_content()  # Sai do iframe
    time.sleep(300)  # Espera 5 minutos antes da nova tentativa
    retorno = funcoes.integrar(navegador, returns)


except Exception as e:
    funcoes.marcacao_cod(f"Erro inesperado: {e}. Aguardando 5 minutos antes de tentar novamente.", 'erro')
    navegador.switch_to.default_content()  # Sai do iframe
    time.sleep(300)  # Espera 5 minutos antes da nova tentativa
    retorno = funcoes.integrar(navegador, returns)


#Criar Excel com resultado
pathSalvar = os.getenv('PATHSAVE')
df = pd.DataFrame(retorno, columns=['Arquivo', 'Retorno'])
df.to_excel(pathSalvar, index=False)

#enviar e-mail
funcoes.enviar_email_com_anexo(pathSalvar)