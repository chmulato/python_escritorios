# ================================================================
# Exemplo Prático 14: Gerador Automático de Procurações e Petições
#
# Este script lê um modelo de documento Word (.docx), substitui campos
# variáveis por dados fornecidos e salva um novo documento preenchido.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - pip install python-docx
# - O modelo deve conter campos como {{nome}}, {{cpf}}, {{data}}, etc.
# ================================================================
from docx import Document

# Função para ler o modelo de documento
def ler_modelo(caminho_modelo):
    return Document(caminho_modelo)

# Função para preencher o documento com os dados do cliente
def preencher_documento(doc, dados_cliente):
    for par in doc.paragraphs:
        for chave, valor in dados_cliente.items():
            if f'{{{{{chave}}}}}' in par.text:
                par.text = par.text.replace(f'{{{{{chave}}}}}', valor)
    return doc

# Função para salvar o documento preenchido
def salvar_documento(doc, caminho_saida):
    doc.save(caminho_saida)

# Exemplo de uso
modelo_path = r'C:\dev\python_escritorios\codes\modelo_procuracao.docx'
saida_path = r'C:\dev\python_escritorios\codes\procuracao_preenchida.docx'
dados_cliente = {
    'nome': 'João da Silva',
    'cpf': '123.456.789-00',
    'data': '01/06/2025'
}

doc = ler_modelo(modelo_path)
doc_preenchido = preencher_documento(doc, dados_cliente)
salvar_documento(doc_preenchido, saida_path)

print(f"Procuração gerada com sucesso: {saida_path}")
