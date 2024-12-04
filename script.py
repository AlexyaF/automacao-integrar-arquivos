from selenium import webdriver
import funcoes
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException
from ftplib import FTP
from selenium.webdriver.chrome.service import Service 
import pandas as pd
import os
from dotenv import load_dotenv

#Carregando palavras chaves
load_dotenv()

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
    print("Conexão estabelecida com FTP")
except Exception as e:
    print(f"Erro ao conectar ao FTP: {e}")


servico = Service(ChromeDriverManager().install()) # Identifica a versão do navegador atual e vai baixar o Chrome Driver mais recente.
navegador = webdriver.Chrome()
navegador.set_page_load_timeout(10000) 

folders=os.getenv("FOLDERS")
folders_list = folders.strip("[]").split(",") # Removendo colchetes e transformando em lista
folders_list = [folder.strip() for folder in folders_list] # Removendo espaços em branco ao redor de cada item (se houver)

#abrir navegador
funcoes.abrir_driver(navegador)

#Para Cada CIA que tenha arquivos quais precisam ser importados
for folder in folders_list:
    try:
        #Buscar arquivos no FTP
        funcoes.mover_arquivos_processado(folder, ftp)

        #Verificação arquivos, retirando tabulações, preparando para integração
        funcoes.verifArquivos()
        
        #Integrar
        retorno = funcoes.integrar(navegador)

        print(f"Processo concluído com sucesso para a pasta: {folder}")

    except TimeoutException as e:
        print(f"Erro de timeout durante o processamento da pasta {folder}: {e}")
    
    except Exception as e:
        # Captura qualquer outro erro
        print(f"Erro inesperado durante o processamento da pasta {folder}: {e}")


print("Encerrando conexão FTP")
ftp.quit()


#Criar Excel com resultado
pathSalvar = os.getenv('PATHSAVE')
df = pd.DataFrame(retorno, columns=['Arquivo', 'Retorno'])
df.to_excel(pathSalvar, index=False)

#enviar e-mail
funcoes.enviar_email_com_anexo(pathSalvar)