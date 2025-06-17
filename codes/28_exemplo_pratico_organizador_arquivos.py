# ================================================================
# Exemplo Prático 28: Organizador de Arquivos em Pastas por Cliente
#
# Este script move arquivos de uma pasta de entrada para subpastas
# organizadas por cliente, com base no nome do cliente no arquivo.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install pandas (se for usar planilha para nomes de clientes)
# ================================================================
import os
import shutil

# Caminho da pasta de entrada com os arquivos a organizar
pasta_entrada = r'C:\dev\python_escritorios\codes\arquivos_entrada'
# Caminho da pasta de saída onde os arquivos serão organizados
pasta_saida = r'C:\dev\python_escritorios\codes\arquivos_organizados'

os.makedirs(pasta_saida, exist_ok=True)

# Função para extrair o nome do cliente do nome do arquivo
def extrair_cliente(nome_arquivo):
    # Exemplo: "cliente123_documento.pdf" -> "cliente123"
    return nome_arquivo.split('_')[0]

# Organizar arquivos
for nome_arquivo in os.listdir(pasta_entrada):
    caminho_arquivo = os.path.join(pasta_entrada, nome_arquivo)
    if os.path.isfile(caminho_arquivo):
        cliente = extrair_cliente(nome_arquivo)
        pasta_cliente = os.path.join(pasta_saida, cliente)
        os.makedirs(pasta_cliente, exist_ok=True)
        destino = os.path.join(pasta_cliente, nome_arquivo)
        shutil.move(caminho_arquivo, destino)
        print(f"Arquivo '{nome_arquivo}' movido para '{pasta_cliente}'.")

print("Organização concluída.")
