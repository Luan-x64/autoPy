import imaplib
import email
from email.header import decode_header
import os
import time
import xml.etree.ElementTree as ET
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from fuzzywuzzy import process
import re



# Configurações do servidor IMAP e credenciais
IMAP_SERVER = ''
IMAP_PORT = 993
EMAIL_USER = ''
EMAIL_PASS = ''

# Diretório onde o arquivo XML será salvo
SAVE_DIR = './xml_files/'

# Configuração da API do Google Sheets
scope = ['https://spreadsheets.google.com/feeds']
credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
gc = gspread.authorize(credentials)

# Abrir a planilha pelo ID e selecionar a primeira aba
spreadsheet_id = ''
worksheet = gc.open_by_key(spreadsheet_id).get_worksheet(0)


def update_google_sheet(cliente, ordem_compra, nf, data_fat, valor_nf):
    header = worksheet.row_values(4)  # Obtém a primeira linha (cabeçalho)

    col_cliente = header.index('CLIENTE') + 1
    col_ordem = header.index('ORDEM DE COMPRA') + 1
    col_nf = header.index('NF') + 1
    col_datafat = header.index('DATA_FAT') + 1
    col_valor = header.index('Valor') + 1

    # Limite máximo de linhas para buscar (80 linhas)
    max_rows = min(len(worksheet.col_values(col_cliente)), 100)

    cliente_col_values = worksheet.col_values(col_cliente)[:max_rows]
    ordem_col_values = worksheet.col_values(col_ordem)[:max_rows]
    valor_col_values = worksheet.col_values(col_valor)[:max_rows]

    # Converte os dados de comparação para maiúsculas para case-insensitive matching
    cliente_upper = cliente.upper()
    ordem_compra_upper = ordem_compra.strip()  # Não converter para maiúsculas, pois é um número
    valor_nf_str = str(valor_nf).upper()

    # Criação de listas para Fuzzy Matching
    table_client_names = [c.upper() for c in cliente_col_values]
    oc_table = [oc.strip() for oc in ordem_col_values]  # Mantém a forma original dos números
    valor_table = [valor.upper() for valor in valor_col_values]

    # Fuzzy matching para valor + nome do cliente
    valor_match, score_valor = process.extractOne(valor_nf_str, valor_table)
    if score_valor >= 65:
        matched_index = valor_table.index(valor_match)
        matched_row = matched_index + 1
        closest_match, score_cliente = process.extractOne(cliente_upper, [table_client_names[matched_index]])
        if score_cliente >= 65:
            # Verificar e atualizar NF e DATA_FAT
            nf_value = worksheet.cell(matched_row, col_nf).value
            datafat_value = worksheet.cell(matched_row, col_datafat).value
            if not nf_value and not datafat_value:
                worksheet.update_cell(matched_row, col_nf, nf)
                worksheet.update_cell(matched_row, col_datafat, data_fat)
                print(f"Atualizado na linha {matched_row}.")
            else:
                print(f"Linha {matched_row} já contém dados em NF e/ou DATA_FAT, ignorando a atualização.")
            return

    # Caso o valor + nome do cliente não sejam encontrados, verificar nome do cliente + ordem de compra
    closest_match, score_cliente = process.extractOne(cliente_upper, table_client_names)

    # Verificação direta para ordem de compra
    if ordem_compra_upper in oc_table:
        matched_index = oc_table.index(ordem_compra_upper)
        matched_row = matched_index + 1
        ordem_col_value = oc_table[matched_index].strip()
        
        # Confirmar se o nome do cliente também bate suficientemente
        if ordem_col_value == ordem_compra_upper and score_cliente >= 65:
            # Verificar e atualizar NF e DATA_FAT
            nf_value = worksheet.cell(matched_row, col_nf).value
            datafat_value = worksheet.cell(matched_row, col_datafat).value
            if not nf_value and not datafat_value:
                worksheet.update_cell(matched_row, col_nf, nf)
                worksheet.update_cell(matched_row, col_datafat, data_fat)
                print(f"Atualizado na linha {matched_row}.")
            else:
                print(f"Linha {matched_row} já contém dados em NF e/ou DATA_FAT, ignorando a atualização.")
        else:
            print("Google: Ordem de compra ou nome do cliente não correspondem suficientemente.")
    else:
        print("Google: Valor + Nome ou Nome + Ordem de Compra não encontrados ou não correspondem suficientemente.")
# Função para conectar e buscar e-mails
def fetch_emails(mail):
    try:
        # Seleciona a caixa de entrada
        mail.select('inbox')

        # Busca por e-mails não lidos do remetente específico
        search_criteria = '(UNSEEN FROM "wkjenw@hotmail.com")'
        status, messages = mail.search(None, search_criteria)

        if status != 'OK':
            print("Erro ao buscar e-mails.")
            return

        # Obtemos os IDs dos e-mails
        email_ids = messages[0].split()

        for email_id in email_ids:
            # Buscar o e-mail
            status, msg_data = mail.fetch(email_id, '(RFC822)')
            msg = email.message_from_bytes(msg_data[0][1])

            # Decodificar o assunto
            subject, encoding = decode_header(msg['Subject'])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else 'utf-8')

            # Iterar sobre as partes do e-mail
            for part in msg.walk():
                if part.get_content_maintype() == 'multipart':
                    continue

                if part.get('Content-Disposition') is None:
                    continue

                filename = part.get_filename()
                if filename:
                    if filename.endswith('.xml'):
                        # Salvar o arquivo XML
                        filepath = os.path.join(SAVE_DIR, filename)
                        if not os.path.isdir(SAVE_DIR):
                            os.makedirs(SAVE_DIR)
                        with open(filepath, 'wb') as f:
                            f.write(part.get_payload(decode=True))
                        print(f'Salvo: {filepath}')
                        process_xml(filepath)
    except imaplib.IMAP4.abort:
        print("Conexão interrompida. Tentando reconectar...")
        mail = conn_imap()
        fetch_emails(mail)

def process_xml(filepath):
    try:
        # Parse o arquivo XML
        tree = ET.parse(filepath)
        root = tree.getroot()

        # Definir o namespace
        ns = {'ns0': 'http://www.portalfiscal.inf.br/nfe'}

        # Extrair as informações desejadas
        # xNome do cliente (destinatário)
        xNome_cliente = root.find('.//ns0:dest/ns0:xNome', namespaces=ns).text if root.find('.//ns0:dest/ns0:xNome', namespaces=ns) is not None else 'Não encontrado'
        # xNome da empresa (emissor)
        xPed = root.find('.//ns0:xPed', namespaces=ns).text if root.find('.//ns0:xPed', namespaces=ns) is not None else 'Não encontrado'
        nNF = root.find('.//ns0:nNF', namespaces=ns).text if root.find('.//ns0:nNF', namespaces=ns) is not None else 'Não encontrado'
        dhRecbto = root.find('.//ns0:dhRecbto', namespaces=ns).text if root.find('.//ns0:dhRecbto', namespaces=ns) is not None else 'Não encontrado'
        valor_nf = root.find('.//ns0:vOrig', namespaces=ns).text if root.find('.//ns0:vOrig', namespaces=ns) is not None else 'Não encontrado'
         # Verificar se xPed é "Não encontrado" e procurar a informação em <infAdic>
        if xPed == 'Não encontrado':
            infCpl = root.find('.//ns0:infCpl', namespaces=ns).text if root.find('.//ns0:dhRecbto', namespaces=ns) is not None else 'Não encontrado'
            ordem_compra = re.search(r'ORDEM DE COMPRA NR (\d+)', infCpl)
            if ordem_compra:
                numero_ordem_compra = ordem_compra.group(1)
                xPed = numero_ordem_compra
            else:
                xPed = 'VERBAL'

            
        # Formatar dhRecbto para o formato desejado
        if dhRecbto != 'Não encontrado':
            try:
                # Converter a string para um objeto datetime
                dt = datetime.fromisoformat(dhRecbto.replace('Z', '+00:00'))
                # Formatar para dia/mês/ano - Hora
                dhRecbto_formatado = dt.strftime('%d/%m/%Y - %H:%M')
            except ValueError:
                dhRecbto_formatado = 'Formato inválido'
        else:
            dhRecbto_formatado = 'Não encontrado'

        # Exibir as informações extraídas
        print(f"xNome (cliente): {xNome_cliente}")
        print(f"xPed: {xPed}")
        print(f"nNF: {nNF}")
        print(f"dhRecbto: {dhRecbto_formatado}")
        print(f"Valor NF: {valor_nf}")
        update_google_sheet(xNome_cliente, xPed, nNF, dhRecbto_formatado, valor_nf)

    except Exception as e:
        print(f"Erro ao processar o XML: {e}")
    finally:
        # Remover o arquivo após o processamento
        os.remove(filepath)
        print(f"Arquivo removido: {filepath}")

def conn_imap():
    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(EMAIL_USER, EMAIL_PASS)
    return mail

def main():
    #update_google_sheet('tgm', '34635', '3213', '3312312')
    #update_google_sheet('menegotti', '185121', '321312313', '3312312')
    # Conectar ao servidor IMAP
    #filename = '41240773244626000117550010001271551170845586-procNFe.xml'
    #process_xml(os.path.join(SAVE_DIR, filename))
    mail = conn_imap()  # Conecta ao IMAP
    while True:
        fetch_emails(mail)
        # Aguardar 60 segundos antes de verificar novamente
        time.sleep(5)

if __name__ == '__main__':
    main()
