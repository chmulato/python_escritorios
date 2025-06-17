# ================================================================
# Exemplo Prático 23: Leitura de Pedidos de Marketplaces
#
# Este script lê um arquivo CSV com pedidos de marketplaces,
# gera um relatório com o resumo dos pedidos e salva em CSV.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install pandas
# ================================================================
import pandas as pd

# Função para ler os pedidos recebidos
def ler_pedidos(caminho_arquivo):
    return pd.read_csv(caminho_arquivo, sep=';', encoding='latin1')

# Função para gerar o relatório com o resumo dos pedidos
def gerar_relatorio_pedidos(df_pedidos, caminho_saida):
    relatorio = df_pedidos.groupby('Produto').agg({'Quantidade': 'sum', 'Preço': 'mean'}).reset_index()
    relatorio['Total Vendas'] = relatorio['Quantidade'] * relatorio['Preço']
    relatorio.to_csv(caminho_saida, index=False, sep=';')
    print(f"Relatório de pedidos gerado: {caminho_saida}")

# Exemplo de uso
caminho_pedidos = r'C:\dev\python_escritorios\codes\pedidos_marketplace.csv'
caminho_relatorio = r'C:\dev\python_escritorios\codes\relatorio_pedidos.csv'

df_pedidos = ler_pedidos(caminho_pedidos)
print("Pedidos lidos:")
print(df_pedidos)

gerar_relatorio_pedidos(df_pedidos, caminho_relatorio)
