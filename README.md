# Sobre o Projeto  
**Este repositório tem como intuito a automatização do download de arquivos de um servidor FTP, sua integração em um sistema web via Selenium e a organização e armazenamento dos arquivos localmente.**

### **🚀 Principais funcionalidades:**  
- Conexão automática com o servidor FTP  
- Download apenas de arquivos ainda não processados  
- Integração dos arquivos em um sistema web via Selenium  
- Captura e armazenamento das respostas do sistema  
- Geração de um relatório em Excel  
- Envio automático do relatório para os líderes da equipe  

### **🛠️ Tecnologias Utilizadas**
- Python 
- Selenium para automação web
- Pandas para manipulação de dados
- FTP Library para conexão com o servidor
- OpenPyXL para geração do relatório em Excel

### **⚙️ Principais Funções**
- `mover_arquivos_processado` → Responsável por verificar e mover arquivos já integrados para as pastas corretas. Se um arquivo for enviado para a pasta de "processado" por engano, ele ainda será integrado, minimizando falhas na automação.
- `mover_arquivos_txt` → Identifica arquivos que ainda não foram processados (aqueles que não estão em nenhuma pasta específica). Após o processamento, a função move os arquivos para a pasta de "processado".
- `integrar` → Responsável pelo upload dos arquivos na interface web, incluindo tratamento de exceções, espera de resposta e armazenamento dos resultados.

### **📚 Aprendizados com o Projeto**
  **- Biblioteca `dotenv`**: Ajuda a manter credências e variáveis sensíveis fora do codígo fonte, permitindo que você carregue as váriaveis de um arquivo .env para o código principal.  
    **Como usar?**  
      1- Instalar biblioteca:  
    ``` 
    pip install python-dotenv
    ```  
      2- Criar arquivo .env:  
    <sub> O arquivo .env não deve ser comitado no Git. Para garantir isso, adicione o arquivo .env ao arquivo .gitignore </sub>  
    ```
    PASS=EX859
    ```  
      3- Importar biblioteca e utilizar função para carregar variáves:  
     ```
     from dotenv import load_dotenv #importando biblioteca para o arquivo

     load_dotenv() #chamando função responsável por carregar variáveis do arquivo .env

     password = os.getenv("PASS") # Acessando uma variável específica
     ```  

  **- Comparação entre listas sem a necessidade de um loop:**  
      Antes, eu realizava comparações entre listas utilizando um loop para verificar a presença de cada elemento em uma lista de referência.
      ```
      # Listas para armazenar os itens únicos
      itens_unicos_exemplo1 = []
      
      # Comparação utilizando loop
      for item in exemplo1:
          if item not in exemplo2:
              itens_unicos_exemplo1.append(item) 
      ```

      No entanto, é possível utilizar conjuntos (set) para tornar essa comparação mais eficiente e legível.

      **Exemplo usando conjuntos (set):**
    ```
      #Listas de exemplo
      exemplo1 = [1, 2, 3, 4, 5]
      exemplo2 = [4, 5, 6, 7]
      
      # Comparação usando set
      itens_unicos_exemplo1 = list(set(exemplo1) - set(exemplo2))
      itens_unicos_exemplo2 = list(set(exemplo2) - set(exemplo1))
    ```
        - O conjunto à esquerda do operador de subtração **sempre** serve como referência na comparação, resultando apenas em seus elementos exclusivos, ou seja, aqueles que não estão presentes no conjunto à direita.  
        -`set()`: Elimina duplicatas e permite a comparação direta entre elementos de duas listas.  
        -list(..):  Converte o resultado de volta para uma lista.  
  **- Usar console navegador (DevTools):**  
