# ================================================================
# Exemplo Prático: Geração e Envio Automático de Relatório Diário de Vendas
#
# Este script gera dados fictícios de vendas, salva-os em um arquivo
# Excel (.xlsx), exibe os dados no console para facilitar a cópia
# para o Excel e envia o arquivo como anexo por e-mail.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - Certifique-se de instalar as bibliotecas necessárias: pip install pandas openpyxl faker
# - Verifique se o arquivo Excel está no caminho correto.
# ================================================================
import pandas as pd
from faker import Faker
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import os

# Configurações do servidor SMTP
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_user = 'seu_email@gmail.com'
smtp_password = 'sua_senha'

# Caminho do arquivo Excel
excel_path = r'C:\dev\python_escritorios\codes\relatorio_vendas.xlsx'

# 1. Geração de dados fictícios de vendas
fake = Faker('pt_BR')
dados_vendas = []
for _ in range(10):  # Gerar 10 registros fictícios
    dados_vendas.append({
        'data': fake.date_this_month(),
        'produto': fake.word(),
        'quantidade': fake.random_int(min=1, max=10),
        'preco_unitario': fake.random_int(min=50, max=500)
    })

# Criar DataFrame
df_vendas = pd.DataFrame(dados_vendas)

# 2. Salvando dados em um arquivo Excel (.xlsx)
df_vendas.to_excel(excel_path, index=False, sheet_name='Vendas')

# 3. Exibindo dados no console para cópia manual para o Excel
print("Dados de Vendas Gerados:")
print(df_vendas)

# 4. Envio do relatório como anexo por e-mail
data_atual = datetime.now().strftime('%d/%m/%Y')
assunto = f'Relatório de Vendas - {data_atual}'
corpo = 'Prezada Ana, segue em anexo o relatório de vendas do dia.'

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

# Enviar o e-mail com o relatório
enviar_email_com_anexo(
    'ana@techsolutions.com',
    assunto,
    corpo,
    excel_path
)

print("E-mail enviado com sucesso!")
# Limpeza do arquivo Excel após o envio
if os.path.exists(excel_path):
    os.remove(excel_path)
    print(f"Arquivo {excel_path} removido após o envio.")