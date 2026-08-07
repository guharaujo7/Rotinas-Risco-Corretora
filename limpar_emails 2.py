"""
Limpa um arquivo .txt (gerado pelo extrator) e deixa apenas os e-mails,
um por linha, prontos para usar em PROCV/VLOOKUP.

Uso:
    python limpar_emails.py
Uma janela abrirá para você selecionar o arquivo .txt de origem.

Gera na mesma pasta: emails_limpos.txt
"""

import re
import tkinter as tk
from tkinter import filedialog

PADRAO_EMAIL = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")


def selecionar_arquivo():
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    caminho = filedialog.askopenfilename(
        title="Selecione o arquivo .txt com os e-mails extraídos",
        filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")],
    )
    root.destroy()
    return caminho


def main():
    caminho = selecionar_arquivo()
    if not caminho:
        print("Nenhum arquivo selecionado. Encerrando.")
        return

    with open(caminho, "r", encoding="utf-8", errors="ignore") as f:
        conteudo = f.read()

    emails = PADRAO_EMAIL.findall(conteudo)

    caminho_saida = caminho.rsplit(".", 1)[0] + "_limpos.txt"
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write("\n".join(emails))

    print(f"{len(emails)} e-mail(s) salvos (com duplicados) em:\n{caminho_saida}")


if __name__ == "__main__":
    main()
