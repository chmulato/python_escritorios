# ================================================================
# Exemplo Prático 25: Envio de Notas Fiscais e Respostas Automáticas a Clientes
#
# Este script gera uma nota fiscal eletrônica (CSV) e envia por e-mail
# como anexo para o cliente, junto com uma resposta automática.
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

# Função para gerar a nota fiscal eletrônica
def gerar_nota_fiscal(dados_venda, caminho_saida):
    df_nota = pd.DataFrame([dados_venda])
    df_nota.to_csv(caminho_saida, index=False, sep=';')
    print(f"Nota fiscal gerada: {caminho_saida}")

# Função para enviar a nota fiscal por e-mail
def enviar_nota_fiscal_email(destinatario, assunto, corpo, caminho_nota, smtp_user, smtp_password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    msg = MIMEMultipart()
    msg['Subject'] = assunto
    msg['From'] = smtp_user
    msg['To'] = destinatario
    msg.attach(MIMEText(corpo, 'plain'))
    with open(caminho_nota, 'rb') as anexo:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(anexo.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f'attachment; filename={os.path.basename(caminho_nota)}')
        msg.attach(parte)
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)
    print(f"E-mail enviado para {destinatario} com a nota fiscal em anexo.")

# Exemplo de uso
dados_venda = {
    'Cliente': 'João da Silva',
    'Produto': 'Notebook',
    'Quantidade': 1,
    'Valor Unitário': 5000.00,
    'CNPJ': '12.345.678/0001-95'
}
caminho_nota = r'C:\dev\python_escritorios\codes\nota_fiscal_joao_silva.csv'
gerar_nota_fiscal(dados_venda, caminho_nota)

# Enviar a nota fiscal por e-mail (descomente a linha abaixo para enviar o e-mail)
# enviar_nota_fiscal_email(
#     'joao@email.com',
#     'Sua Nota Fiscal',
#     'Segue em anexo a nota fiscal referente à sua compra. Obrigado pela preferência!',
#     caminho_nota,
#     'seu_email@gmail.com',
#     'sua_senha'
# )
