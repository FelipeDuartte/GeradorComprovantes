import pandas as pd
from docx import Document
import re
import sys
import os
import tkinter as tk
from tkinter import messagebox
import unicodedata
from tkinter import ttk

# =============================
# CAMINHO SEGURO (APENAS PARA ARQUIVOS EMBUTIDOS NO EXE)
# =============================
def caminho_absoluto(nome_arquivo):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, nome_arquivo)
    return os.path.join(os.path.abspath("."), nome_arquivo)

# =============================
# NORMALIZAR TEXTO
# =============================
def normalizar(texto):
    texto = str(texto).lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto

# =============================
# IDENTIFICAR COLUNAS AUTOMATICAMENTE
# =============================
def identificar_colunas(colunas):
    nome_col = None
    cpf_col = None
    valor_col = None
    candidatos_valor = []

    for col in colunas:
        col_norm = normalizar(col)

        if not nome_col and ("nome" in col_norm or "bolsista" in col_norm):
            nome_col = col

        if not cpf_col and "cpf" in col_norm:
            cpf_col = col

        if "valor" in col_norm or "total" in col_norm or "recebido" in col_norm:
            score = 0
            if "valor total" in col_norm:
                score += 5
            if "total recebido" in col_norm:
                score += 5
            if "recebido" in col_norm:
                score += 3
            if "total" in col_norm:
                score += 2
            if "valor" in col_norm:
                score += 1
            if "parcela" in col_norm or "mensal" in col_norm:
                score -= 3

            candidatos_valor.append((score, col))

    if candidatos_valor:
        candidatos_valor.sort(reverse=True)
        valor_col = candidatos_valor[0][1]

    if not all([nome_col, cpf_col, valor_col]):
        raise Exception(
            "Não foi possível identificar automaticamente as colunas.\n\n"
            "Verifique se o Excel possui colunas de Nome, CPF e Valor."
        )

    return {
        "nome": nome_col,
        "cpf": cpf_col,
        "valor": valor_col
    }

# =============================
# FORMATADORES
# =============================
def formatar_cpf(cpf):
    cpf = re.sub(r"\D", "", str(cpf))
    return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}" if len(cpf) == 11 else cpf

def formatar_valor(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# =============================
# SUBSTITUIR TEXTO NO WORD
# =============================
def substituir_texto(doc, mapa):
    for p in doc.paragraphs:
        for run in p.runs:
            for chave, valor in mapa.items():
                if chave in run.text:
                    run.text = run.text.replace(chave, valor)

    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for p in celula.paragraphs:
                    for run in p.runs:
                        for chave, valor in mapa.items():
                            if chave in run.text:
                                run.text = run.text.replace(chave, valor)

# =============================
# GERAR COMPROVANTES

def gerar_comprovantes(barra, status_label, botao):
    try:
        botao.config(state="disabled")
        status_label.config(text="Iniciando...")

        df = pd.read_excel("dadosteste.xlsx")
        df.columns = df.columns.str.strip()

        colunas = identificar_colunas(df.columns)

        total = len(df)
        barra["maximum"] = total
        barra["value"] = 0

        os.makedirs("output", exist_ok=True)

        for i, (_, linha) in enumerate(df.iterrows(), start=1):
            nome = str(linha[colunas["nome"]])
            cpf = formatar_cpf(linha[colunas["cpf"]])
            valor = formatar_valor(linha[colunas["valor"]])

            doc = Document(
                caminho_absoluto("MODELO COMPROVANTE RENDIMENTOS.docx")
            )

            substituir_texto(doc, {
                "{{Nome}}": nome,
                "{{CPF}}": cpf,
                "{{Valor}}": valor
            })

            doc.save(os.path.join("output", f"{nome}.docx"))

            barra["value"] = i
            status_label.config(
                text=f"Gerando {i} de {total} comprovantes..."
            )
            janela.update_idletasks()

        status_label.config(text="Concluído ✔")
        messagebox.showinfo(
            "Sucesso",
            "Comprovantes gerados com sucesso!\n\nConfira a pasta 'output'."
        )

    except Exception as e:
        messagebox.showerror("Erro", str(e))

    finally:
        botao.config(state="normal")

# =============================
# INTERFACE
# =============================
janela = tk.Tk()
janela.title("Gerador de Comprovantes")
janela.geometry("420x260")
janela.resizable(False, False)

tk.Label(
    janela,
    text="Gerador de Comprovantes",
    font=("Segoe UI", 15, "bold")
).pack(pady=20)

barra = ttk.Progressbar(
    janela,
    orient="horizontal",
    length=300,
    mode="determinate"
)
barra.pack(pady=15)

status_label = tk.Label(
    janela,
    text="Aguardando...",
    font=("Segoe UI", 10)
)
status_label.pack()

botao = tk.Button(
    janela,
    text="Gerar comprovantes",
    font=("Segoe UI", 11),
    width=28,
    height=2,
    command=lambda: gerar_comprovantes(barra, status_label, botao)
)
botao.pack(pady=20)

janela.mainloop()

