# ================================================================
# Exemplo Prático 10: Extração de Tabelas HTML para Excel
#
# Capítulo: 5. Web scraping e automação de sites
# Objetivo: Demonstrar como extrair tabelas de páginas HTML e salvar em Excel,
#           útil para escritórios que precisam coletar dados tabulares de sites,
#           como cotações, tabelas de processos, preços, etc.
#
# Área de aplicação: Escritórios de advocacia, contabilidade, logística, vendas.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install requests beautifulsoup4 pandas openpyxl
# ================================================================
import requests
from bs4 import BeautifulSoup
import pandas as pd

# URL do site com a tabela (substitua por um site real para testes)
url = 'https://example.com/tabela'

# Fazer uma requisição HTTP para obter o conteúdo da página
response = requests.get(url)
if response.status_code == 200:
    soup = BeautifulSoup(response.text, 'html.parser')
    tabela = soup.find('table')
    if tabela:
        # Extrair a tabela usando pandas
        df = pd.read_html(str(tabela))[0]
        print("Tabela extraída:")
        print(df)
        # Salvar em Excel
        caminho_excel = r'C:\dev\python_escritorios\codes\tabela_extraida.xlsx'
        df.to_excel(caminho_excel, index=False)
        print(f"Tabela salva em: {caminho_excel}")
    else:
        print("Nenhuma tabela encontrada na página.")
else:
    print(f"Erro ao acessar o site: {response.status_code}")
