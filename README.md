# 📜 Sobre o Projeto  
**Este repositório tem como intuito a automatização do download de arquivos de um servidor FTP, sua integração em um sistema web via Selenium e a organização e armazenamento dos arquivos localmente.**

**🚀 Principais funcionalidades:**  
- Conexão automática com o servidor FTP  
- Download apenas de arquivos ainda não processados  
- Integração dos arquivos em um sistema web via Selenium  
- Captura e armazenamento das respostas do sistema  
- Geração de um relatório em Excel  
- Envio automático do relatório para os líderes da equipe  
  
  
**🛠️ Tecnologias Utilizadas**
- Python 🐍
- Selenium para automação web
- Pandas para manipulação de dados
- FTP Library para conexão com o servidor
- OpenPyXL para geração do relatório em Excel

  **⚙️ Principais Funções**
- `mover_arquivos_processado` → Responsável por verificar e mover arquivos já integrados para as pastas corretas. Se um arquivo for enviado para a pasta de "processado" por engano, ele ainda será integrado, minimizando falhas na automação.
- `mover_arquivos_txt` → Identifica arquivos que ainda não foram processados (aqueles que não estão em nenhuma pasta específica). Após o processamento, a função move os arquivos para a pasta de "processado".
- `integrar` → Responsável pelo upload dos arquivos na interface web, incluindo tratamento de exceções, espera de resposta e armazenamento dos resultados.
