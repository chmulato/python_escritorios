# Exemplo prático 3.6.1 - Análise de vendas de produtos online (CSV)
#
# Este script lê um arquivo 'vendas.csv' contendo registros de vendas de uma loja online,
# calcula o total vendido por produto e gera um novo arquivo 'relatorio_vendas.csv' com o resultado.
# Se já existir um relatório anterior, ele será removido antes de criar o novo.
#
# O arquivo 'vendas.csv' deve conter as colunas: data, produto, quantidade, preco_unitario.
# O relatório gerado terá as colunas: produto, total_vendido.
#
# Caminhos dos arquivos:
# - Entrada: C:\dev\python_escritorios\codes\vendas.csv
# - Saída:   C:\dev\python_escritorios\codes\relatorio_vendas.csv
# Exemplo prático 3.6.1 - Análise de vendas de produtos online (CSV)
#
# Este script lê um arquivo 'vendas.csv' contendo registros de vendas de uma loja online,
# calcula o total vendido por produto e gera um novo arquivo 'relatorio_vendas.csv' com o resultado.
# Se já existir um relatório anterior, ele será removido antes de criar o novo.
#
# O arquivo 'vendas.csv' deve conter as colunas: data, produto, quantidade, preco_unitario.
# O relatório gerado terá as colunas: produto, total_vendido.
#
# Caminhos dos arquivos:
# - Entrada: C:\dev\python_escritorios\codes\vendas.csv
# - Saída:   C:\dev\python_escritorios\codes\relatorio_vendas.csv
import os
import pandas as pd

# Caminhos dos arquivos
csv_path = r'C:\dev\python_escritorios\codes\vendas.csv'
relatorio_path = r'C:\dev\python_escritorios\codes\relatorio_vendas.csv'

# Apaga o relatório antigo, se existir
if os.path.exists(relatorio_path):
    os.remove(relatorio_path)

# Lê o arquivo de vendas
df = pd.read_csv(csv_path)

# Calcula o total vendido por produto
df['total'] = df['quantidade'] * df['preco_unitario']
df_relatorio = df.groupby('produto')['total'].sum().reset_index()
df_relatorio.rename(columns={'total': 'total_vendido'}, inplace=True)

# Salva o novo relatório
df_relatorio.to_csv(relatorio_path, index=False)

print("Relatório de vendas gerado com sucesso!")