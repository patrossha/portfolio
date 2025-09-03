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

def connect_and_read_email_dcp():
    server = 'imap.skymail.net.br'
    try:
        mail = imaplib.IMAP4_SSL(server)
        mail.login(username, password)

        folder_name = 'INBOX.Andamentos.DCP' # imp dcp
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
                            if part.get_content_type() in ["text/plain", "text/html"]:
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

                    
                    body = unescape(body).replace("<br>", "\n")
                    body = re.sub(r'<[^>]+>', '', body)  #html solver

                    
                    blocos = re.split(r'(?=Processo:\s*\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4})', body)

                    for bloco in blocos:
                        # busca o número cnj
                        match_numero = re.search(r'Processo:\s*(\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4})', bloco)
                        numero_processo = match_numero.group(1) if match_numero else "Não encontrado"

                        # busca movimento
                        match_fase = re.search(r'ÚLTIMO MOVIMENTO:\s*(.*?)(?=\n|$)', bloco, re.IGNORECASE)
                        fase_atual = match_fase.group(1).strip() if match_fase else "Não encontrado"

                        if numero_processo != "Não encontrado":
                            print(f"Processo: {numero_processo} | Fase Atual: {fase_atual}")
                            email_data.append([numero_processo, fase_atual])

        
        df_dcp = pd.DataFrame(email_data, columns=['Processo', 'Fase Atual'])

        df_dcp.to_excel("andamentos_dcp.xlsx", index=False)
        print("Extração concluída. Os dados foram salvos.")

    except Exception as e:
        print(f"Erro ao conectar ou ler o e-mail: {e}")

connect_and_read_email_dcp()
