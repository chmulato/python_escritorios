# ================================================================
# Exemplo Prático 19: Envio Automático de Boletos por E-mail
#
# Este script gera um boleto bancário simples (CSV) e envia por e-mail
# como anexo para o cliente.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install pandas
# ================================================================
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Função para gerar o boleto bancário (CSV)
def gerar_boleto(dados_boleto, caminho_saida):
    df_boleto = pd.DataFrame([dados_boleto])
    df_boleto.to_csv(caminho_saida, index=False, sep=';')
    print(f"Boleto gerado: {caminho_saida}")

# Função para enviar o boleto por e-mail
def enviar_boleto_email(destinatario, assunto, corpo, caminho_boleto, smtp_user, smtp_password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    msg = MIMEMultipart()
    msg['Subject'] = assunto
    msg['From'] = smtp_user
    msg['To'] = destinatario
    msg.attach(MIMEText(corpo, 'plain'))
    with open(caminho_boleto, 'rb') as anexo:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(anexo.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f'attachment; filename={os.path.basename(caminho_boleto)}')
        msg.attach(parte)
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)
    print(f"E-mail enviado para {destinatario} com o boleto em anexo.")

# Exemplo de uso
dados_boleto = {
    'Cliente': 'João da Silva',
    'Valor': 150.00,
    'Vencimento': '2025-06-10'
}
caminho_boleto = r'C:\dev\python_escritorios\codes\boleto_joao_silva.csv'
gerar_boleto(dados_boleto, caminho_boleto)

# Enviar o boleto por e-mail (descomente a linha abaixo para enviar o e-mail)
# enviar_boleto_email('joao@email.com', 'Seu Boleto', 'Segue em anexo o boleto.', caminho_boleto, 'seu_email@gmail.com', 'sua_senha')
