"""
Extrator de e-mails de arquivos .msg (Outlook)
------------------------------------------------
Procura, dentro do corpo (ou assunto) de arquivos .msg, o trecho:
    "The following address(es) failed permanent fatal errors"
e extrai o e-mail que aparece entre os caracteres < >.

Gera:
  - emails_extraidos.txt  -> lista de e-mails encontrados
  - arquivos_com_erro.txt -> lista de arquivos onde não foi possível extrair

Requisitos:
    pip install extract-msg
"""

import os
import re
import sys
import tkinter as tk
from tkinter import filedialog

try:
    import extract_msg
except ImportError:
    print("Biblioteca 'extract-msg' não encontrada.")
    print("Instale com:  pip install extract-msg")
    sys.exit(1)


# Padrão que localiza o texto de erro (aceita pequenas variações de digitação/maiúsculas)
PADRAO_TRECHO = re.compile(
    r"following\s+address(?:es)?.{0,60}?failed.{0,60}?fatal\s+errors",
    re.IGNORECASE | re.DOTALL,
)

# Padrão para capturar o e-mail dentro de < >
PADRAO_EMAIL_ANGULO = re.compile(r"<\s*([^<>\s]+@[^<>\s]+)\s*>")

# Fallback: qualquer e-mail "solto" no texto, caso não haja <>
PADRAO_EMAIL_SOLTO = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")


def extrair_email_do_texto(texto):
    """Tenta encontrar o e-mail logo após o trecho de erro, preferindo o que está entre < >."""
    if not texto:
        return None

    match_trecho = PADRAO_TRECHO.search(texto)
    if match_trecho:
        # Olha um pedaço do texto logo após o trecho encontrado
        inicio = match_trecho.end()
        janela = texto[inicio: inicio + 500]

        m = PADRAO_EMAIL_ANGULO.search(janela)
        if m:
            return m.group(1)

        m = PADRAO_EMAIL_SOLTO.search(janela)
        if m:
            return m.group(0)

    # Se não achou o trecho de erro, tenta achar qualquer e-mail entre < > no texto todo
    m = PADRAO_EMAIL_ANGULO.search(texto)
    if m:
        return m.group(1)

    return None


def processar_arquivo(caminho):
    """Abre um .msg e tenta extrair o e-mail do corpo, depois do assunto."""
    try:
        msg = extract_msg.Message(caminho)
    except Exception as e:
        return None, f"Erro ao abrir arquivo: {e}"

    try:
        corpo = msg.body or ""
        assunto = msg.subject or ""
    except Exception as e:
        return None, f"Erro ao ler conteúdo: {e}"
    finally:
        try:
            msg.close()
        except Exception:
            pass

    # 1) Tenta no corpo do e-mail
    email = extrair_email_do_texto(corpo)
    if email:
        return email, None

    # 2) Tenta no assunto
    email = extrair_email_do_texto(assunto)
    if email:
        return email, None

    return None, "Não foi possível localizar o e-mail no corpo nem no assunto"


def selecionar_pasta():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    pasta = filedialog.askdirectory(title="Selecione a pasta com os arquivos .msg")
    root.destroy()
    return pasta


def main():
    pasta = selecionar_pasta()
    if not pasta:
        print("Nenhuma pasta selecionada. Encerrando.")
        return

    arquivos_msg = [
        f for f in os.listdir(pasta) if f.lower().endswith(".msg")
    ]

    if not arquivos_msg:
        print("Nenhum arquivo .msg encontrado na pasta selecionada.")
        return

    emails_encontrados = []
    arquivos_com_erro = []

    print(f"Encontrados {len(arquivos_msg)} arquivos .msg. Processando...\n")

    for nome_arquivo in arquivos_msg:
        caminho_completo = os.path.join(pasta, nome_arquivo)
        email, erro = processar_arquivo(caminho_completo)

        if email:
            emails_encontrados.append((nome_arquivo, email))
            print(f"[OK]   {nome_arquivo} -> {email}")
        else:
            arquivos_com_erro.append((nome_arquivo, erro))
            print(f"[FALHA] {nome_arquivo} -> {erro}")

    # Salva os e-mails extraídos
    caminho_emails = os.path.join(pasta, "emails_extraidos.txt")
    with open(caminho_emails, "w", encoding="utf-8") as f:
        f.write("E-mails extraídos dos arquivos .msg\n")
        f.write("=" * 40 + "\n\n")
        for nome_arquivo, email in emails_encontrados:
            f.write(f"{email}\t({nome_arquivo})\n")
        f.write(f"\nTotal: {len(emails_encontrados)} e-mail(s) extraído(s).\n")

    # Salva a lista de e-mails únicos também (útil para copiar/colar)
    caminho_unicos = os.path.join(pasta, "emails_extraidos_unicos.txt")
    emails_unicos = sorted(set(email for _, email in emails_encontrados))
    with open(caminho_unicos, "w", encoding="utf-8") as f:
        f.write("\n".join(emails_unicos))

    # Salva os arquivos com erro
    caminho_erros = os.path.join(pasta, "arquivos_com_erro.txt")
    with open(caminho_erros, "w", encoding="utf-8") as f:
        f.write("Arquivos onde NÃO foi possível extrair o e-mail\n")
        f.write("=" * 40 + "\n\n")
        for nome_arquivo, erro in arquivos_com_erro:
            f.write(f"{nome_arquivo}\t-> {erro}\n")
        f.write(f"\nTotal: {len(arquivos_com_erro)} arquivo(s) com falha.\n")

    print("\n" + "=" * 50)
    print(f"Concluído! {len(emails_encontrados)} e-mail(s) extraído(s), "
          f"{len(arquivos_com_erro)} arquivo(s) com falha.")
    print(f"Resultados salvos em:\n  {caminho_emails}\n  {caminho_unicos}\n  {caminho_erros}")


if __name__ == "__main__":
    main()
