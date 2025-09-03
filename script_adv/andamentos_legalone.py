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

try:
    mail = imaplib.IMAP4_SSL("imap.skymail.net.br")
    mail.login(username, password)
    print("Login bem-sucedido")
except imaplib.IMAP4.error as e:
    print(f"Erro de login: {e}")


def extrair_info_html(body_html):
    from bs4 import BeautifulSoup
    import re

    soup = BeautifulSoup(body_html, 'html.parser')
    texto = soup.get_text(separator='\n')

    processo_regex = re.compile(r'\d{7}-\d{2}\.\d{4}\.\d\.\d{2}\.\d{4}')
    
    # Lista para evitar duplicação
    numeros_processados = set()
    resultados = []

    for tag_processo in soup.find_all(string=processo_regex):
        numero_match = processo_regex.search(tag_processo)
        if not numero_match:
            continue

        numero = numero_match.group()

        if numero in numeros_processados:
            continue  # já processado, pula
        numeros_processados.add(numero)

        texto = soup.get_text(separator='\n')
        inicio = texto.find(numero)
        fim = len(texto)
        bloco_texto = texto[inicio:fim]

        cliente = re.search(r'Cliente\s+(?:principal)?:?\s*(.*?)\s{2,}', bloco_texto)
        contrario = re.search(r'Contrário\s+(?:principal)?:?\s*(.*?)\s{2,}', bloco_texto)

        cliente_nome = cliente.group(1).split(" ")[0] if cliente else ""
        contrario_nome = contrario.group(1).split(" ")[0] if contrario else ""
        partes = f"{cliente_nome} X {contrario_nome}" if cliente and contrario else ""

        responsavel_match = re.search(r'Responsável\s+(?:principal)?:?\s*([^\n\r]+)', bloco_texto)
        responsavel = responsavel_match.group(1).split(" ")[0].strip() if responsavel_match else ""

        descricao = ""
        for elem in tag_processo.parent.next_elements:
            if elem.name == 'table':
                descricoes = []
                for linha in elem.find_all('tr'):
                    celulas = linha.find_all('td')
                    for td in celulas:
                        if td.has_attr('colspan') and td['colspan'] == '2':
                            texto_desc = td.get_text(strip=True)
                            if texto_desc and texto_desc.strip().lower() != "descrição":
                                descricoes.append(texto_desc)
                descricao = " || ".join(descricoes)
                break

        resultados.append({
            "Número": numero,
            "Partes": partes,
            "Responsável": responsavel,
            "Descrição": descricao
        })

    return resultados


def connect_and_read_email_legalone():
    mail.select("INBOX.Andamentos")

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

        andamentos = extrair_info_html(body)
        todos_resultados.extend(andamentos)

    df_legalone = pd.DataFrame(todos_resultados)
    print (df_legalone)


    df_legalone.to_excel("andamentos_legalone.xlsx", index=False)
    print("Extração concluída. Os dados foram salvos.")


if __name__ == "__main__":
    connect_and_read_email_legalone()
