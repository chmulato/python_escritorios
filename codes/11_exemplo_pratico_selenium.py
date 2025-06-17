# ================================================================
# Exemplo Prático 11: Automação de Sites com Selenium
#
# Capítulo: 5. Web scraping e automação de sites
# Objetivo: Demonstrar como automatizar o login e a navegação em sites,
#           útil para escritórios que precisam acessar sistemas web internos,
#           portais de clientes, tribunais, ou baixar relatórios protegidos.
#
# Área de aplicação: Escritórios de advocacia, contabilidade, logística, vendas.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install selenium
# - Baixe o ChromeDriver em: https://sites.google.com/chromium.org/driver/
# ================================================================
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Configurar o driver do navegador (certifique-se de que o ChromeDriver está no PATH)
driver = webdriver.Chrome()  # ou webdriver.Firefox() para Firefox

try:
    # Abrir a página de login
    driver.get('https://example.com/login')

    # Encontrar os campos de login e senha
    username_field = driver.find_element(By.NAME, 'username')
    password_field = driver.find_element(By.NAME, 'password')

    # Preencher os campos de login
    username_field.send_keys('seu_usuario')
    password_field.send_keys('sua_senha')

    # Enviar o formulário (pressionar Enter)
    password_field.send_keys(Keys.RETURN)

    # Esperar alguns segundos para a página carregar após o login
    time.sleep(5)

    # Extrair informações da página após o login
    dados = driver.find_elements(By.CLASS_NAME, 'informacao')
    print("Informações extraídas após login:")
    for dado in dados:
        print(dado.text)

finally:
    # Fechar o navegador
    driver.quit()
