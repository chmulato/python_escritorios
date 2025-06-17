# ================================================================
# Exemplo Prático 22: Acompanhamento de Entregas via Planilhas
#
# Este script lê uma planilha Excel com dados das entregas e envia
# notificações automáticas por e-mail para os clientes informando o status.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install pandas openpyxl
# ================================================================
import pandas as pd
import smtplib
from email.mime.text import MIMEText

# Função para ler a planilha de entregas
def ler_planilha_entregas(caminho_arquivo):
    return pd.read_excel(caminho_arquivo)

# Função para enviar notificações automáticas
def enviar_notificacoes_entregas(df_entregas, smtp_user, smtp_password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    for _, row in df_entregas.iterrows():
        destinatario = row['Email']
        assunto = f"Atualização sobre sua entrega - {row['Código da Entrega']}"
        corpo = (
            f"Prezado(a),\n\n"
            f"Informamos que o status da sua entrega {row['Código da Entrega']} foi atualizado para: {row['Status']}.\n\n"
            f"Atenciosamente,"
        )
        msg = MIMEText(corpo)
        msg['Subject'] = assunto
        msg['From'] = smtp_user
        msg['To'] = destinatario
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(msg)
        print(f"E-mail enviado para {destinatario} sobre a entrega {row['Código da Entrega']}.")

# Exemplo de uso
caminho_planilha = r'C:\dev\python_escritorios\codes\entregas.xlsx'
df_entregas = ler_planilha_entregas(caminho_planilha)
print("Entregas lidas da planilha:")
print(df_entregas)

# Enviar notificações (descomente a linha abaixo para enviar os e-mails)
# enviar_notificacoes_entregas(df_entregas, 'seu_email@gmail.com', 'sua_senha')

print("Processamento concluído.")
