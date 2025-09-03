import imaplib
import email
import ssl
import re
import pandas as pd
from html import unescape
from dotenv import load_dotenv
import os

load_dotenv()
username = os.getenv("EMAIL_USUARIO")
password = os.getenv("EMAIL_SENHA")

def connect_and_read_email_eproc():
    server = 'imap.skymail.net.br'

    try:
        mail = imaplib.IMAP4_SSL(server)
        mail.login(username, password)

        # imp eproc
        folder_name = 'INBOX.Andamentos.eProc'
        folder_name_encoded = folder_name.replace(' ', '\040')  

        mail.select(folder_name_encoded)  

        print("Conexão bem-sucedida")

       
        status, messages = mail.search(None, 'UNSEEN')

        if status != "OK":
            print("Nenhum e-mail não lido encontrado.")
            return

        email_data = []

        for email_id in messages[0].split():
            status, msg_data = mail.fetch(email_id, '(RFC822)')

            if status != "OK":
                print("Erro ao recuperar o e-mail.")
                continue

            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])

                    body = ""
                    if msg.is_multipart():
                        for part in msg.walk():
                            content_type = part.get_content_type()
                            content_disposition = str(part.get("Content-Disposition"))

                            if "attachment" not in content_disposition:
                                if content_type == "text/plain":
                                    try:
                                        body = part.get_payload(decode=True).decode('utf-8')
                                    except UnicodeDecodeError:
                                        body = part.get_payload(decode=True).decode('ISO-8859-1')  
                                    break
                                elif content_type == "text/html":
                                    try:
                                        body = part.get_payload(decode=True).decode('utf-8')
                                    except UnicodeDecodeError:
                                        body = part.get_payload(decode=True).decode('ISO-8859-1')  
                                    break
                    else:
                        try:
                            body = msg.get_payload(decode=True).decode('utf-8')
                        except UnicodeDecodeError:
                            body = msg.get_payload(decode=True).decode('ISO-8859-1')  

                    
                    body = unescape(body)
                    body = body.replace("</td>", "\n").replace("<td>", "").replace("<tr>", "").replace("</tr>", "")
                    body = body.replace("<br>", "\n")
 
                    # Regex
                    processo_match = re.search(r'Num\.\s*Processo[:\-]?\s*(\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4})', body, re.IGNORECASE)
                    numero_processo = processo_match.group(1) if processo_match else "Não encontrado"
                    print(f"Numero do Processo: {numero_processo}") 

                    movimentacao_match = re.search(r'<td.*?>movimentação[:\-]?\s*(.*?)\s*<td.*?>evento número', body, re.IGNORECASE)

                    if movimentacao_match:
                        movimentacao = movimentacao_match.group(1).strip()
                    else:
                        movimentacao = "Não encontrado" 

                    print(f"Movimentação: {movimentacao}") 

                    email_data.append([numero_processo, movimentacao])

        df_eproc = pd.DataFrame(email_data, columns=['Número do Processo', 'Movimentação'])

        df_eproc.to_excel("andamentos_eproc.xlsx", index=False)
        print("Extração concluída. Os dados foram salvos.")

    except Exception as e:
        print(f"Erro ao conectar ou ler o e-mail: {e}")

connect_and_read_email_eproc()
