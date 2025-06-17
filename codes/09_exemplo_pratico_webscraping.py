# ================================================================
# Exemplo Prático 09: Web Scraping Básico com BeautifulSoup
#
# Capítulo: 5. Web scraping e automação de sites
# Objetivo: Demonstrar como coletar informações de sites automaticamente,
#           aplicável a escritórios que precisam extrair dados públicos,
#           como acompanhamento de processos, preços, notícias, etc.
#
# Área de aplicação: Escritórios de advocacia, contabilidade, logística, vendas.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install requests beautifulsoup4
# ================================================================
import requests
from bs4 import BeautifulSoup

# URL do site de exemplo (substitua por um site real para testes)
url = 'https://example.com'

# Fazer uma requisição HTTP para obter o conteúdo da página
response = requests.get(url)
if response.status_code == 200:
    soup = BeautifulSoup(response.text, 'html.parser')
    # Exemplo: extrair todos os textos de divs com a classe 'informacao'
    dados = soup.find_all('div', class_='informacao')
    print("Informações extraídas:")
    for dado in dados:
        print(dado.text.strip())
else:
    print(f"Erro ao acessar o site: {response.status_code}")
