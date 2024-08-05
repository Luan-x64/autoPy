# Auto NF
Este projeto é um script simples em Python desenvolvido para automatizar a extração e atualização de informações de Notas Fiscais Eletrônicas (NF-e) recebidas por e-mail, integrando essas informações diretamente em uma planilha do Google Sheets.

# Funcionalidades
- Conexão automática ao e-mail e filtragem de mensagens de um remetente específico.
- Download de arquivos XML anexados aos e-mails recebidos.
- Extração de informações relevantes do XML, como nome do cliente, número do pedido, número da NF, data de faturamento e valor.
- Atualização automática de uma planilha do Google Sheets com as informações extraídas, facilitando o acompanhamento e controle de pedidos faturados.

# Tecnologias Utilizadas
- Python
- IMAP para conexão ao servidor de e-mail
- XML Parsing com xml.etree.ElementTree
- Integração com Google Sheets usando gspread
- Fuzzy Matching com fuzzywuzzy para garantir a precisão da busca de dados

# Como Utilizar
 - 1. Clone o repositório.
 - 2. Configure as credenciais do e-mail e do Google Sheets no script.
 - 3. Execute o script para iniciar a automação.
 - 4. Este projeto é ideal para quem deseja automatizar o processo de acompanhamento de pedidos e notas fiscais, integrando facilmente informações de e-mails diretamente em planilhas online.
