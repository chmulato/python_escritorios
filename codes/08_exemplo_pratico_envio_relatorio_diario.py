# ================================================================
# Exemplo Prático: Geração e Envio Automático de Relatórios Diários de Vendas (Múltiplos Dias)
#
# Este script gera dados fictícios de vendas para vários dias, salva um arquivo
# Excel para cada dia, exibe os dados no console e envia todos os relatórios diários
# como anexos em um único e-mail.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install pandas openpyxl faker
# ================================================================
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from faker import Faker
from datetime import datetime, timedelta
import os

# Configurações do servidor SMTP
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_user = 'seu_email@gmail.com'
smtp_password = 'sua_senha'

# Parâmetros de geração de dados
num_dias = 3  # Quantidade de dias simulados
registros_por_dia = 5  # Vendas por dia
pasta_relatorios = r'C:\dev\python_escritorios\codes\relatorios_diarios'
os.makedirs(pasta_relatorios, exist_ok=True)

fake = Faker('pt_BR')
arquivos_gerados = []

# 1. Geração de dados fictícios e arquivos diários
for i in range(num_dias):
    data_venda = (datetime.now() - timedelta(days=i)).strftime('%d/%m/%Y')
    vendas = []
    for _ in range(registros_por_dia):
        vendas.append({
            'Data': data_venda,
            'Produto': fake.word(),
            'Quantidade': fake.random_int(min=1, max=10),
            'Valor Total': fake.random_int(min=100, max=2000)
        })
    df = pd.DataFrame(vendas)
    arquivo = os.path.join(pasta_relatorios, f'relatorio_vendas_{data_venda.replace("/", "-")}.xlsx')
    df.to_excel(arquivo, index=False)
    arquivos_gerados.append(arquivo)
    print(f"\nRelatório gerado para {data_venda}:")
    print(df.to_csv(sep='\t', index=False))

# 2. Envio de todos os relatórios diários em um único e-mail
assunto = f'Relatórios de Vendas - Últimos {num_dias} Dias'
corpo = f'Prezada Ana, seguem em anexo os relatórios de vendas dos últimos {num_dias} dias.'

def enviar_email_com_multiplos_anexos(destinatario, assunto, corpo, lista_anexos):
    msg = MIMEMultipart()
    msg['Subject'] = assunto
    msg['From'] = smtp_user
    msg['To'] = destinatario
    msg.attach(MIMEText(corpo, 'plain'))
    for caminho_anexo in lista_anexos:
        with open(caminho_anexo, 'rb') as anexo:
            parte = MIMEBase('application', 'octet-stream')
            parte.set_payload(anexo.read())
            encoders.encode_base64(parte)
            parte.add_header('Content-Disposition', f'attachment; filename={os.path.basename(caminho_anexo)}')
            msg.attach(parte)
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

enviar_email_com_multiplos_anexos(
    'ana@techsolutions.com',
    assunto,
    corpo,
    arquivos_gerados
)

print("\nE-mail com todos os relatórios diários enviado com sucesso!")