# ================================================================
# Exemplo Prático 18: Geração Automática de Guias de Impostos
#
# Este script calcula o imposto devido com base em receitas e despesas,
# e gera um arquivo de guia de pagamento.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - Não requer bibliotecas externas.
# ================================================================

# Função para calcular os impostos devidos
def calcular_impostos(receitas, despesas, aliquota_imposto):
    lucro = receitas - despesas
    imposto_devido = lucro * aliquota_imposto
    return imposto_devido

# Função para gerar a guia de pagamento do imposto
def gerar_guia_pagamento(imposto_devido, caminho_saida):
    with open(caminho_saida, 'w') as file:
        file.write(f"Guia de Pagamento de Imposto\n")
        file.write(f"Valor do Imposto Devido: R$ {imposto_devido:.2f}\n")
    print(f"Guia de pagamento gerada: {caminho_saida}")

# Exemplo de uso
receitas = 10000
despesas = 5000
aliquota_imposto = 0.15  # 15%

imposto_devido = calcular_impostos(receitas, despesas, aliquota_imposto)
gerar_guia_pagamento(imposto_devido, r'C:\dev\python_escritorios\codes\guia_pagamento_imposto.txt')
