# ================================================================
# Exemplo Prático 20: Leitura e Geração de Manifestos (XML, PDF)
#
# Este script lê um manifesto em formato XML, extrai informações relevantes
# e gera um novo manifesto em formato PDF.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install fpdf
# ================================================================
import xml.etree.ElementTree as ET
from fpdf import FPDF

# Função para ler o manifesto em formato XML
def ler_manifesto_xml(caminho_arquivo):
    tree = ET.parse(caminho_arquivo)
    root = tree.getroot()
    dados_manifesto = {
        'remetente': root.find('remetente').text,
        'destinatario': root.find('destinatario').text,
        'peso': root.find('peso').text,
        'valor_frete': root.find('valor_frete').text
    }
    return dados_manifesto

# Função para gerar o manifesto em formato PDF
def gerar_manifesto_pdf(dados_manifesto, caminho_saida):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    for chave, valor in dados_manifesto.items():
        pdf.cell(0, 10, f"{chave}: {valor}", ln=True)
    pdf.output(caminho_saida)
    print(f"Manifesto gerado: {caminho_saida}")

# Exemplo de uso
caminho_xml = r'C:\dev\python_escritorios\codes\manifesto.xml'
caminho_pdf = r'C:\dev\python_escritorios\codes\manifesto.pdf'
dados_manifesto = ler_manifesto_xml(caminho_xml)
gerar_manifesto_pdf(dados_manifesto, caminho_pdf)
