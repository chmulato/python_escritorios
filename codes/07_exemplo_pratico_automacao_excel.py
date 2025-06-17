# ================================================================
# Envio Automático de Relatório Diário de Vendas por E-mail
#
# Este script lê um arquivo Excel (.xlsx) com o relatório de vendas,
# verifica sua existência e envia-o como anexo por e-mail.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - Certifique-se de instalar as bibliotecas necessárias: pip install pandas openpyxl
# - Verifique se o arquivo Excel está no caminho correto.
# ================================================================
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from datetime import datetime
import os
# Configurações do servidor SMTP
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_user = 'seu_email@gmail.com'
smtp_password = 'sua_senha'
# Função para enviar e-mail com anexo
def enviar_email_com_anexo(destinatario, assunto, corpo, caminho_anexo):
    # Criar o objeto de mensagem
    msg = MIMEMultipart()
    msg['Subject'] = assunto
    msg['From'] = smtp_user
    msg['To'] = destinatario
    # Adicionar o corpo do e-mail
    msg.attach(MIMEText(corpo, 'plain'))
    # Adicionar o anexo
    with open(caminho_anexo, 'rb') as anexo:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(anexo.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f'attachment; filename={caminho_anexo.split("/")[-1]}')
        msg.attach(parte)
    # Enviar o e-mail
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)
# Caminho do arquivo Excel
excel_path = r'C:\dev\python_escritorios\codes\relatorio_vendas.xlsx'
# Gerar dados fictícios de vendas
dados_vendas = {
    'Data': [datetime.now().strftime('%d/%m/%Y')],
    'Produto': ['Notebook'],
    'Quantidade': [5],
    'Valor Total': [15000.00]
}
df_vendas = pd.DataFrame(dados_vendas)
# Salvar o DataFrame em um arquivo Excel, sobrescrevendo se já existir
df_vendas.to_excel(excel_path, index=False)
print(f"Arquivo de vendas gerado em: {excel_path}")
# Exibir os dados de vendas no console para copiar e colar no Excel
print("\nCopie e cole os dados abaixo no Excel, se desejar:")
print(df_vendas.to_csv(sep='\t', index=False))
# Enviar o e-mail com o relatório
data_atual = datetime.now().strftime('%d/%m/%Y')
assunto = f'Relatório de Vendas - {data_atual}'
corpo = 'Prezada Ana, segue em anexo o relatório de vendas do dia anterior.'
enviar_email_com_anexo(
    'ana@techsolutions.com',
    assunto,
    corpo,
    excel_path
)
print("E-mail enviado com sucesso!")
# Nota: Certifique-se de que o acesso a aplicativos menos seguros esteja ativado na sua conta do Gmail.
# Além disso, considere usar OAuth2 para uma abordagem mais segura.