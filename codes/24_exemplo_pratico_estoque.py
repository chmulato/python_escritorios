# ================================================================
# Exemplo Prático 24: Atualização Automática de Estoque
#
# Este script lê um arquivo CSV com dados de vendas e atualiza automaticamente
# as quantidades em estoque em uma planilha Excel.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install pandas openpyxl
# ================================================================
import pandas as pd

# Função para ler os dados de vendas
def ler_vendas(caminho_arquivo):
    return pd.read_csv(caminho_arquivo, sep=';', encoding='latin1')

# Função para atualizar o estoque
def atualizar_estoque(df_vendas, caminho_estoque):
    df_estoque = pd.read_excel(caminho_estoque)
    for _, row in df_vendas.iterrows():
        produto = row['Produto']
        quantidade_vendida = row['Quantidade']
        df_estoque.loc[df_estoque['Produto'] == produto, 'Estoque'] -= quantidade_vendida
    df_estoque.to_excel(caminho_estoque, index=False)
    print("Estoque atualizado com sucesso.")

# Exemplo de uso
caminho_vendas = r'C:\dev\python_escritorios\codes\vendas.csv'
caminho_estoque = r'C:\dev\python_escritorios\codes\estoque.xlsx'
df_vendas = ler_vendas(caminho_vendas)
print("Vendas lidas:")
print(df_vendas)
atualizar_estoque(df_vendas, caminho_estoque)
