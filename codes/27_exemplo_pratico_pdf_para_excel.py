# ================================================================
# Exemplo Prático 27: Conversor de Tabelas PDF → Excel
#
# Este script extrai uma tabela de um arquivo PDF e salva os dados em um arquivo Excel.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install pdfplumber pandas openpyxl
# ================================================================
import pdfplumber
import pandas as pd

# Função para ler a tabela do PDF
def ler_tabela_pdf(caminho_arquivo):
    with pdfplumber.open(caminho_arquivo) as pdf:
        primeira_pagina = pdf.pages[0]
        tabela = primeira_pagina.extract_table()
    return tabela

# Função para salvar a tabela em um arquivo Excel
def salvar_tabela_excel(tabela, caminho_saida):
    df_tabela = pd.DataFrame(tabela[1:], columns=tabela[0])
    df_tabela.to_excel(caminho_saida, index=False)
    print(f"Tabela salva em Excel: {caminho_saida}")

# Exemplo de uso
caminho_pdf = r'C:\dev\python_escritorios\codes\tabela.pdf'
caminho_excel = r'C:\dev\python_escritorios\codes\tabela_convertida.xlsx'
tabela_extraida = ler_tabela_pdf(caminho_pdf)
salvar_tabela_excel(tabela_extraida, caminho_excel)
