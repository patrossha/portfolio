from io import StringIO
import imaplib
import email
from bs4 import BeautifulSoup
import pandas as pd
import re
from dotenv import load_dotenv
import os

load_dotenv()
username = os.getenv("EMAIL_USUARIO")
password = os.getenv("EMAIL_SENHA")

def autenticar_email():
    try:
        mail = imaplib.IMAP4_SSL("imap.skymail.net.br")
        mail.login(username, password)
        print("Login bem-sucedido!")
        return mail
    except imaplib.IMAP4.error as e:
        print(f"Erro de login: {e}")
        return None

def limpar_descricao(texto):
    texto = re.sub(r'\s*\n\s*', ' ', texto)
    texto = re.sub(r'\s{2,}', ' ', texto)
    return texto.strip()

def extrair_publicacoes(html):
    soup = BeautifulSoup(html, 'html.parser')
    texto = soup.get_text(separator="\n")

    publicacoes = []
    numeros_processados = set()
    ids_processados = set() 
    bloco_atual = {}

    descricao_regex = r'Descrição:\s*(.*?)Ver no Legal One'
    id_regex = r'Processo\s*:\s*Proc\s*-\s*(\S+)'
    numero_regex = r'Número\s*:\s*([^\n]+)'  # Ajuste para capturar o número
    cliente_regex = r'Cliente principal:\s*([^\n]*)'
    contrario_regex = r'Contrário principal:\s*([^\n]*)'
    responsavel_regex = r'Responsável principal:\s*([^\n]*)'
    blocos = re.split(r'(?=^\s*Descrição:)', texto, flags=re.MULTILINE)

    for bloco in blocos:
        if not bloco.strip():
            continue

        descricao_match = re.search(descricao_regex, bloco, re.DOTALL)
        descricao_raw = descricao_match.group(1).strip() if descricao_match else ''
        descricao = limpar_descricao(descricao_raw)

        if not descricao:
            continue  # Ignorar blocos sem descrição

        # ID do processo
        id_match = re.search(id_regex, bloco)
        id_process = id_match.group(1) if id_match else ''

        if not id_process:
            continue 

        if id_process in ids_processados:
            continue  # Ignorar ID duplicada
        ids_processados.add(id_process)

        numero_match = re.search(numero_regex, bloco)
        numero = numero_match.group(1).strip().replace(": ", "") if numero_match else ''

        if not numero:
            continue 

        # normalizar o número excluindo pastas c agravo/recurso/2 inst
        numero_normalizado = re.sub(r'\/\d+', '', numero)

        if numero_normalizado in numeros_processados:
            continue  # Ignorar número duplicado
        numeros_processados.add(numero_normalizado)

        cliente_match = re.search(cliente_regex, bloco)
        contrario_match = re.search(contrario_regex, bloco)
        responsavel_match = re.search(responsavel_regex, bloco)

        cliente = cliente_match.group(1).strip().replace(":", "") if cliente_match else ''
        contrario = contrario_match.group(1).strip().replace(":", "") if contrario_match else ''
        responsavel = responsavel_match.group(1).strip().replace(":", "") if responsavel_match else ''

        # split
        cliente_nome = cliente.split()[0] if cliente else ''
        contrario_nome = contrario.split()[0] if contrario else ''
        responsavel_nome = responsavel.split()[0] if responsavel else ''

        partes = f"{cliente_nome} X {contrario_nome}" if cliente_nome and contrario_nome else ''

        publicacoes.append({
            'Número': numero,
            'Partes': partes,
            'Responsável principal': responsavel_nome,
            'Descrição': descricao,
            'ID': id_process
        })

    return publicacoes

def connect_and_read_email(mail):
    try:
        mail.select("INBOX.Publicacoes")
    except imaplib.IMAP4.error as e:
        print(f"Erro ao selecionar INBOX: {e}")
        return

    status, messages = mail.search(None, "UNSEEN")
    messages = messages[0].split()

    todos_resultados = []

    for num in messages:
        status, data = mail.fetch(num, "(RFC822)")
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    body = part.get_payload(decode=True).decode(errors="ignore")
                    break
        else:
            if msg.get_content_type() == "text/html":
                body = msg.get_payload(decode=True).decode(errors="ignore")

        if not body:
            continue

        publicacoes = extrair_publicacoes(body)
        todos_resultados.extend(publicacoes)

    df_legalone = pd.DataFrame(todos_resultados)
    df_legalone = df_legalone[['Número', 'Partes', 'Responsável principal', 'Descrição', 'ID']]  # Ordem das colunas
    print(df_legalone)

    df_legalone.to_excel("publicacoes_legalone.xlsx", index=False)
    print("Extração concluída. Os dados foram salvos.")

if __name__ == "__main__":
    mail = autenticar_email()
    if mail:
        connect_and_read_email(mail)
