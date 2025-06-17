# ================================================================
# Relatório de Inadimplentes - Envio Real de E-mails
#
# Este script automatiza a geração de um relatório de inquilinos inadimplentes,
# remove o relatório anterior se existir, gera um novo arquivo Excel,
# calcula estatísticas e envia e-mails reais de cobrança para cada inquilino inadimplente.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - Nunca compartilhe senhas em código público.
# - Teste o envio com poucos destinatários antes de usar em produção.
# - Certifique-se de que o envio de e-mails está de acordo com as políticas da sua empresa
#   e com a legislação vigente (ex: LGPD).
# - Para usar o envio real, adicione uma coluna 'E-mail' na planilha alugueis.xlsx
#   com o e-mail de cada inquilino.
# - Para Gmail, ative a autenticação em dois fatores e gere uma senha de aplicativo
#   para usar neste script.
# ================================================================
import pandas as pd
import os
import smtplib
from email.message import EmailMessage

relatorio_path = r'C:\dev\python_escritorios\codes\relatorio_inadimplentes.xlsx'
relatorio_final_path = r'C:\dev\python_escritorios\codes\relatorio_final.xlsx'
alugueis_path = r'C:\dev\python_escritorios\codes\alugueis.xlsx'

# Configurações do servidor de e-mail (exemplo para Gmail)
EMAIL_ADDRESS = 'seu_email@gmail.com'         # Substitua pelo seu e-mail
EMAIL_PASSWORD = 'sua_senha_de_aplicativo'    # Substitua pela sua senha de aplicativo

# Apagar o relatório anterior, se existir
if os.path.exists(relatorio_path):
    os.remove(relatorio_path)

# Ler a planilha de aluguéis
df = pd.read_excel(alugueis_path)

# Filtrar inadimplentes
inadimplentes = df[df['Pago (Sim/Não)'] == 'Não']

# Salvar relatório dos inadimplentes
inadimplentes.to_excel(relatorio_path, index=False)

# Função para enviar e-mail
def enviar_email(destinatario, nome, imovel, valor, vencimento):
    msg = EmailMessage()
    msg['Subject'] = 'Aviso de Inadimplência - Aluguel em Aberto'
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = destinatario
    corpo = (
        f"Prezado(a) {nome},\n\n"
        f"Informamos que o aluguel do imóvel {imovel} no valor de R$ {valor} "
        f"continua em aberto. Data de vencimento: {vencimento}.\n\n"
        "Por favor, regularize sua situação o quanto antes.\n\n"
        "Atenciosamente,\n"
        "Empresa de Aluguel de Imóveis"
    )
    msg.set_content(corpo)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        smtp.send_message(msg)
    print(f"E-mail enviado para {nome} ({destinatario})")

# Enviar e-mails reais para inadimplentes
# Certifique-se de que a planilha possui uma coluna 'E-mail' com os endereços dos inquilinos
for _, row in inadimplentes.iterrows():
    destinatario = row.get('E-mail', None)
    if destinatario and isinstance(destinatario, str) and '@' in destinatario:
        try:
            enviar_email(
                destinatario,
                row['Inquilino'],
                row['Imóvel'],
                row['Valor Mensal'],
                row['Data de Vencimento']
            )
        except Exception as e:
            print(f"Erro ao enviar e-mail para {row['Inquilino']}: {e}")
    else:
        print(f"E-mail não informado ou inválido para {row['Inquilino']}.")

# Estatísticas
total_inadimplentes = inadimplentes['Valor Mensal'].sum()
total_pagos = df[df['Pago (Sim/Não)'] == 'Sim']['Valor Mensal'].sum()
total_geral = df['Valor Mensal'].sum()
porcentagem_inadimplencia = (total_inadimplentes / total_geral) * 100 if total_geral > 0 else 0

print("\nRelatório Final:")
print(f"Total de aluguéis em aberto: R$ {total_inadimplentes}")
print(f"Total de aluguéis pagos: R$ {total_pagos}")
print(f"Total geral de aluguéis: R$ {total_geral}")
print(f"Porcentagem de inadimplência: {porcentagem_inadimplencia:.2f}%")

# Salvar relatório final
relatorio_final = {
    'Total Aluguéis em Aberto': [total_inadimplentes],
    'Total Aluguéis Pagos': [total_pagos],
    'Total Geral de Aluguéis': [total_geral],
    'Porcentagem de Inadimplência': [porcentagem_inadimplencia]
}
relatorio_df = pd.DataFrame(relatorio_final)
relatorio_df.to_excel(relatorio_final_path, index=False)

print("Relatório final salvo com sucesso em 'relatorio_final.xlsx'.")
print("Processamento concluído com sucesso!")
# Fim do script