# ================================================================
# Relatório de Inadimplentes - Empresa de Aluguel de Imóveis
#
# Este script automatiza a geração de um relatório de inquilinos inadimplentes.
# Ele lê a planilha 'alugueis.xlsx', remove o relatório anterior se existir,
# gera um novo arquivo 'relatorio_inadimplentes.xlsx' apenas com os inadimplentes
# e simula o envio de e-mails de cobrança.
#
# Autor: Christian Mulato
# Data: Junho de 2025
# ================================================================

import pandas as pd
import os

relatorio_path = r'C:\dev\python_escritorios\codes\relatorio_inadimplentes.xlsx'
# 1. Importar bibliotecas
# Importar bibliotecas necessárias
# Apagar o relatório anterior, se existir
if os.path.exists(relatorio_path):
    os.remove(relatorio_path)

# 2. Ler a planilha de aluguéis
df = pd.read_excel(r'C:\dev\python_escritorios\codes\alugueis.xlsx')

# 3. Filtrar inadimplentes
inadimplentes = df[df['Pago (Sim/Não)'] == 'Não']

# 4. Salvar relatório dos inadimplentes
inadimplentes.to_excel(r'C:\dev\python_escritorios\codes\relatorio_inadimplentes.xlsx', index=False)

# 5. Simular envio de e-mails
for _, row in inadimplentes.iterrows():
    print(f"Enviando e-mail para {row['Inquilino']}:")
    print(f"Prezado(a) {row['Inquilino']}, seu aluguel do imóvel {row['Imóvel']} no valor de R$ {row['Valor Mensal']} está em aberto. Vencimento: {row['Data de Vencimento']}.")
    print("------------------------------\n")

# 6. Calcular total de aluguéis em aberto
total_inadimplentes = inadimplentes['Valor Mensal'].sum()
print(f"Total de aluguéis em aberto: R$ {total_inadimplentes}")

# 7. Calcular total de aluguéis pagos
total_pagos = df[df['Pago (Sim/Não)'] == 'Sim']['Valor Mensal'].sum()
print(f"Total de aluguéis pagos: R$ {total_pagos}")

# 8. Calcular total geral de aluguéis
total_geral = df['Valor Mensal'].sum()
print(f"Total geral de aluguéis: R$ {total_geral}")

# 9. Calcular porcentagem de inadimplência
porcentagem_inadimplencia = (total_inadimplentes / total_geral) * 100 if total_geral > 0 else 0
print(f"Porcentagem de inadimplência: {porcentagem_inadimplencia:.2f}%")

# 10. Exibir relatório final
print("\nRelatório Final:")
print(f"Total de aluguéis em aberto: R$ {total_inadimplentes}")
print(f"Total de aluguéis pagos: R$ {total_pagos}")
print(f"Total geral de aluguéis: R$ {total_geral}")

# 11. Salvar relatório final
relatorio_final = {
    'Total Aluguéis em Aberto': [total_inadimplentes],
    'Total Aluguéis Pagos': [total_pagos],
    'Total Geral de Aluguéis': [total_geral],
    'Porcentagem de Inadimplência': [porcentagem_inadimplencia]
}
relatorio_df = pd.DataFrame(relatorio_final)
relatorio_df.to_excel(r'C:\dev\python_escritorios\codes\relatorio_final.xlsx', index=False)

# 12. Exibir mensagem de conclusão
print("Relatório final salvo com sucesso em 'relatorio_final.xlsx'.")

# 13. Fim do script
print("Processamento concluído.")