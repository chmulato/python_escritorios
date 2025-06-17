# ================================================================
# Exemplo Prático 13: Interface Web Simples com PyWebIO
#
# Este script cria um formulário de contato web usando PyWebIO.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install pywebio
# ================================================================
from pywebio import start_server
from pywebio.input import input_group, input
from pywebio.output import put_text

def app():
    dados = input_group("Formulário de Contato", [
        input("Nome:"),
        input("Email:"),
        input("Mensagem:", type="textarea")
    ])
    put_text(f"Obrigado pelo contato, {dados[0]}!")
    put_text(f"Seu e-mail: {dados[1]}")
    put_text(f"Mensagem recebida: {dados[2]}")

if __name__ == '__main__':
    start_server(app, port=8080)
