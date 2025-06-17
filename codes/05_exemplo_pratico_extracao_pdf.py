# ================================================================
# Extração de Texto de PDF - Exemplo Prático
#
# Este script realiza a extração de texto de um arquivo PDF e salva
# o conteúdo em um arquivo .txt, facilitando o processamento de documentos
# em repartições públicas e escritórios.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - Certifique-se de instalar a biblioteca PyPDF2 antes de executar: pip install PyPDF2
# - Verifique se o arquivo PDF está no caminho correto.
# ================================================================
import PyPDF2

# Caminho do arquivo PDF
pdf_path = r'C:\dev\python_escritorios\codes\oficio_exemplo.pdf'
txt_path = r'C:\dev\python_escritorios\codes\oficio_extraido.txt'

# Abrir o PDF e extrair texto
with open(pdf_path, 'rb') as file:
    reader = PyPDF2.PdfReader(file)
    texto = ""
    for page in reader.pages:
        texto += page.extract_text()

# Salvar o texto extraído em um arquivo .txt
with open(txt_path, 'w', encoding='utf-8') as f:
    f.write(texto)

# Mostrar as primeiras linhas do texto extraído
print("Primeiras linhas do texto extraído:")
print('\n'.join(texto.splitlines()[:10]))