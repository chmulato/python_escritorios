# ================================================================
# Exemplo Prático 15: Controle de Prazos Processuais com Alertas
#
# Este script lê uma planilha Excel com prazos processuais, identifica
# prazos próximos do vencimento e envia alertas por e-mail.
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
from datetime import datetime

# Função para ler a planilha de prazos
def ler_planilha(caminho_arquivo):
    return pd.read_excel(caminho_arquivo)

# Função para verificar prazos próximos do vencimento
def verificar_prazos(df, dias_aviso=3):
    hoje = datetime.now()
    # Converter coluna para datetime se necessário
    if not pd.api.types.is_datetime64_any_dtype(df['Data de Término']):
        df['Data de Término'] = pd.to_datetime(df['Data de Término'])
    proximos_prazos = df[(df['Data de Término'] - hoje).dt.days <= dias_aviso]
    return proximos_prazos

# Função para enviar e-mail de alerta
def enviar_alertas(df_prazos, smtp_user, smtp_password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    for _, row in df_prazos.iterrows():
        destinatario = row['Email']
        assunto = f"Aviso de Prazo Processual - {row['Número do Processo']}"
        corpo = (
            f"Prezado(a),\n\n"
            f"Este é um aviso de que o prazo processual referente ao processo {row['Número do Processo']} "
            f"está se aproximando.\nData de Término: {row['Data de Término'].strftime('%d/%m/%Y')}\n\nAtenciosamente,"
        )
        msg = MIMEText(corpo)
        msg['Subject'] = assunto
        msg['From'] = smtp_user
        msg['To'] = destinatario
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(msg)

# Exemplo de uso
caminho_planilha = r'C:\dev\python_escritorios\codes\prazos_processuais.xlsx'
df = ler_planilha(caminho_planilha)
df_prazos_proximos = verificar_prazos(df)

print("Prazos próximos do vencimento:")
print(df_prazos_proximos)

# Enviar alertas (descomente a linha abaixo para enviar os e-mails)
# enviar_alertas(df_prazos_proximos, 'seu_email@gmail.com', 'sua_senha')

print("Processamento concluído.")
