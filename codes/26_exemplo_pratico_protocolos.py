# ================================================================
# Exemplo Prático 26: Controle Automatizado de Protocolos em Repartição Pública
#
# Este script lê uma planilha Excel com dados dos protocolos e envia
# notificações automáticas por e-mail para os responsáveis.
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

# Função para ler a planilha de protocolos
def ler_planilha_protocolos(caminho_arquivo):
    return pd.read_excel(caminho_arquivo)

# Função para enviar notificações automáticas
def enviar_notificacoes_protocolos(df_protocolos, smtp_user, smtp_password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    for _, row in df_protocolos.iterrows():
        destinatario = row['Email']
        assunto = f"Atualização sobre seu Protocolo - {row['Número do Protocolo']}"
        corpo = (
            f"Prezado(a),\n\n"
            f"Informamos que a situação do seu protocolo {row['Número do Protocolo']} foi atualizada para: {row['Situação']}.\n\n"
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
        print(f"E-mail enviado para {destinatario} sobre o protocolo {row['Número do Protocolo']}.")

# Exemplo de uso
caminho_planilha = r'C:\dev\python_escritorios\codes\protocolos_reparticao_publica.xlsx'
df_protocolos = ler_planilha_protocolos(caminho_planilha)
print("Protocolos lidos da planilha:")
print(df_protocolos)

# Enviar notificações (descomente a linha abaixo para enviar os e-mails)
# enviar_notificacoes_protocolos(df_protocolos, 'seu_email@gmail.com', 'sua_senha')

print("Processamento concluído.")
