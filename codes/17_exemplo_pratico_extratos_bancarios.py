# ================================================================
# Exemplo Prático 17: Leitura e Consolidação de Extratos Bancários (CSV)
#
# Este script lê múltiplos arquivos CSV de extratos bancários,
# consolida as informações e gera um relatório financeiro resumido.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install pandas
# ================================================================
import pandas as pd
import glob
import os

# Função para ler o extrato bancário
def ler_extrato(caminho_arquivo):
    return pd.read_csv(caminho_arquivo, sep=';', encoding='latin1')

# Função para consolidar extratos de uma pasta
def consolidar_extratos(pasta_extratos):
    arquivos = glob.glob(os.path.join(pasta_extratos, '*.csv'))
    extratos = [ler_extrato(arquivo) for arquivo in arquivos]
    return pd.concat(extratos, ignore_index=True)

# Função para gerar o relatório financeiro
def gerar_relatorio(df_consolidado, caminho_saida):
    relatorio = df_consolidado.groupby('Tipo')['Valor'].sum().reset_index()
    relatorio.to_csv(caminho_saida, index=False, sep=';')
    print(f"Relatório financeiro gerado: {caminho_saida}")

# Exemplo de uso
pasta_extratos = r'C:\dev\python_escritorios\codes\extratos_bancarios'
caminho_relatorio = r'C:\dev\python_escritorios\codes\relatorio_financeiro.csv'

df_consolidado = consolidar_extratos(pasta_extratos)
print("Extratos consolidados:")
print(df_consolidado)

gerar_relatorio(df_consolidado, caminho_relatorio)
