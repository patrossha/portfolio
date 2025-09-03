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

mail = imaplib.IMAP4_SSL("imap.skymail.net.br")

try:
    mail.login(username, password)
    print("Login bem-sucedido!")
except imaplib.IMAP4.error as e:
    print(f"Erro de login: {e}")
    exit()

# lista de advogados cadastrados busca oab
lista_advogados = [
    "Adv1",
    "Adv2",
    "Adv3"
]

def extrair_publicacoes_texto(texto):
    padrao_publicacao = re.findall(
        r'PROCESSO:\s*(\d{7}-\d{2}\.\d{4}\.\d{1}\.\d{2}\.\d{4})(.*?)(?=(?:Publicação|Publicao):|\Z)',
        texto, re.DOTALL
    )

    dados = []

    for numero_processo, descricao in padrao_publicacao:
        descricao = descricao.strip()
        advogado_mencionado = None
        for advogado in lista_advogados:
            if re.search(rf"\b{re.escape(advogado)}\b", descricao, re.IGNORECASE):
                advogado_mencionado = advogado
                break
        if advogado_mencionado:
            dados.append({
                "Número do Processo": numero_processo.strip(),
                "Advogado": advogado_mencionado,
                "Descrição": re.sub(r'\s+', ' ', descricao)
            })

    return dados

def connect_and_read_email():
    mail.select("INBOX.Publicacoes.OAB")
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
                if part.get_content_type() in ["text/plain", "text/html"]:
                    try:
                        body = part.get_payload(decode=True).decode(errors="ignore")
                        break
                    except:
                        continue
        else:
            body = msg.get_payload(decode=True).decode(errors="ignore")

        if not body:
            continue

        if "<html" in body.lower():
            soup = BeautifulSoup(body, "html.parser")
            texto = soup.get_text(separator="\n")
        else:
            texto = body

        resultados = extrair_publicacoes_texto(texto)
        todos_resultados.extend(resultados)

    try:
        df_processos = pd.read_excel('Processos20251.xlsx')
        print("Planilha de clientes/processos carregada com sucesso")
    except Exception as e:
        print(f"Erro ao carregar a planilha de processos/clientes: {e}")
        return

    if not todos_resultados:
        print("Nenhuma publicação relevante encontrada nos e-mails.")

        # Criar DataFrame vazio
        df_vazio = pd.DataFrame(columns=['Número do Processo', 'partes', 'Advogado', 'Descrição'])
        df_vazio.to_excel("recorte_oab.xlsx", index=False)
        print("Planilha vazia 'recorte_oab.xlsx' criada.")
        return


    df_oab = pd.DataFrame(todos_resultados)

    if df_processos.empty or df_oab.empty:
        print("Uma das planilhas está vazia.")
        return

    df_merged = df_oab.merge(
        df_processos[['Número do Processo', 'Advogado', 'Cliente principal', 'Contrário principal']],
        on='Número do Processo',
        how='left',
        suffixes=('_email', '_planilha')
    )

    df_merged['Advogado'] = df_merged['Advogado_planilha'].apply(lambda x: x.split()[0] if isinstance(x, str) else '')

    df_merged['cliente_nome'] = df_merged['Cliente principal'].apply(lambda x: x.split()[0] if isinstance(x, str) else '')
    df_merged['contrario_nome'] = df_merged['Contrário principal'].apply(lambda x: x.split()[0] if isinstance(x, str) else '')
    df_merged['partes'] = df_merged.apply(
        lambda row: f"{row['cliente_nome']} X {row['contrario_nome']}" if row['cliente_nome'] and row['contrario_nome'] else '',
        axis=1
    )

    df_final = df_merged[['Número do Processo', 'partes', 'Advogado', 'Descrição']]
    df_final = df_final.drop_duplicates(subset=["Número do Processo"])

    df_final.to_excel("recorte_oab.xlsx", index=False)
    print("Publicações extraídas e salvas com sucesso na planilha 'recorte_oab.xlsx'.")

if __name__ == "__main__":
    connect_and_read_email()

