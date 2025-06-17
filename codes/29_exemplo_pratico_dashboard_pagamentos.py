# ================================================================
# Exemplo Prático 29: Dashboard de Pagamentos
#
# Este script lê uma planilha Excel com dados de pagamentos e exibe
# um dashboard simples com gráficos usando pandas e matplotlib.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install pandas matplotlib openpyxl
# ================================================================
import pandas as pd
import matplotlib.pyplot as plt

# Caminho do arquivo Excel com os dados de pagamentos
caminho_excel = r'C:\dev\python_escritorios\codes\pagamentos.xlsx'

# Ler os dados da planilha
df = pd.read_excel(caminho_excel)

# Exibir resumo dos pagamentos
print("Resumo dos pagamentos:")
print(df.groupby('Status')['Valor'].sum())

# Gráfico de barras: Total pago x Em aberto
df_status = df.groupby('Status')['Valor'].sum()
df_status.plot(kind='bar', title='Total de Pagamentos por Status', ylabel='Valor (R$)')
plt.tight_layout()
plt.show()

# Gráfico de pizza: Distribuição dos pagamentos por cliente
df_cliente = df.groupby('Cliente')['Valor'].sum()
df_cliente.plot(kind='pie', autopct='%1.1f%%', title='Distribuição dos Pagamentos por Cliente')
plt.ylabel('')
plt.tight_layout()
plt.show()
