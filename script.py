from selenium import webdriver
import funcoes
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from selenium.webdriver.common.action_chains import ActionChains
import os
from time import sleep
from dotenv import load_dotenv

#Carregando palavras chaves
load_dotenv()


#Verificação arquivos, retirando tabulações, preparando para integração
funcoes.verifArquivos()


#Abrir navegador/ site
servico = Service(ChromeDriverManager().install()) # Identifica a versão do navegador atual e vai baixar o Chrome Driver mais recente.
navegador = webdriver.Chrome()
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
    wait = WebDriverWait(navegador, 10)  # Substitua 'navegador' pelo nome do driver
    element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="field5-suggestions"]//span[text()="RETORNO CPFL"]')))
    element.click()
except:
    print("ERRO!")


#arquivos 
pathFolder = os.getenv('PATHFOLDER')
files = os.listdir(pathFolder)

#Salvar retorno da tela
returns = []

for file in files:
    allpath = os.path.join(pathFolder, file)
    
    #Upload arquivo
    wait = WebDriverWait(navegador, 10) # Aguarda até o elemento estar visível
    navegador.switch_to.frame("cont") # mudando para o iframe
    input_element = wait.until(EC.visibility_of_element_located((By.ID, 'Upload1'))) #aguardo até a visibilidade do elemento e identifico 
    input_element.send_keys(allpath) #envio o arquivo

    #Integrar click
    navegador.find_element('xpath','//*[@id="Btncef"]').click()
  

    #salvar retorno tela
    reotrno = navegador.find_element('xpath', '//*[@id="Label9"]')
    if reotrno:
        texto = navegador.find_element('xpath', '//*[@id="Label9"]').text
    
    navegador.switch_to.default_content() #saindo do iframe 
    joinArquivoRetorno = [file, texto]
    returns.append(joinArquivoRetorno)



#Criar Excel com resultado
pathSalvar = os.getenv('PATHSAVE')
df = pd.DataFrame(returns, columns=['Arquivo', 'Retorno'])
df.to_excel(pathSalvar, index=False)

#enviar e-mail
funcoes.enviar_email_com_anexo(pathSalvar)