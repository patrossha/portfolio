import imaplib
import email
import re
import pandas as pd
from dotenv import load_dotenv
import os

load_dotenv()
username = os.getenv("EMAIL_USUARIO")
password = os.getenv("EMAIL_SENHA")

def get_email_body(msg):
    body = None
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition"))
            if content_type in ["text/plain", "text/html"] and "attachment" not in content_disposition:
                payload = part.get_payload(decode=True)
                if payload:
                    try:
                        body = payload.decode('utf-8')
                    except UnicodeDecodeError:
                        body = payload.decode('ISO-8859-1')
                break
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            try:
                body = payload.decode('utf-8')
            except UnicodeDecodeError:
                body = payload.decode('ISO-8859-1')
    return body or "Corpo do e-mail não encontrado"

def connect_and_read_email_esaj():
    server = 'imap.skymail.net.br'

    try:
        mail = imaplib.IMAP4_SSL(server)
        mail.login(username, password)
        folder_name = 'INBOX.Andamentos.eSAJ' # imp esaj
        mail.select(folder_name)

        print("Conexão bem-sucedida!")

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
                    body = get_email_body(msg)

                
                    # limpa o aviso final
                    body = re.split(r'AVISO|Aviso|aviso|_{5,}|-{5,}', body)[0]

                    # ajuste rotulo de recurso
                    processo_pattern = re.findall(
                        r'(?:Processo:|(?:Execução de Sentença:.*?|Recurso:.*?))\s*(\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4})',
                        body
                    )

                    movimentos_pattern = re.findall(
                        r'Novas Movimentações\s*([\s\S]+?)(?=(?:Processo:|Execução de Sentença:|Recurso:|$))',
                        body,
                        re.DOTALL
                    )

                    for i, processo in enumerate(processo_pattern):
                        if i < len(movimentos_pattern):
                            bloco = movimentos_pattern[i]

                            linhas = bloco.strip().splitlines()
                            movimentos_formatados = []

                            for linha in linhas:
                                linha = linha.strip()
                                if not linha:
                                    continue
                                movimento = re.sub(r'^\d{2}/\d{2}/\d{4}(?:\s+\d{2}:\d{2})?\s+', '', linha)
                                movimentos_formatados.append(movimento)

                            texto_final = "\n".join(movimentos_formatados).strip()
                            email_data.append([processo, texto_final])

        df_esaj = pd.DataFrame(email_data, columns=['Número do Processo', 'Movimentos'])
        
        df_esaj.to_excel("andamentos_esaj.xlsx", index=False)
        print("Extração concluída. Os dados foram salvos.")

    except Exception as e:
        print(f"Erro ao conectar ou ler o e-mail: {str(e)}")

connect_and_read_email_esaj()
