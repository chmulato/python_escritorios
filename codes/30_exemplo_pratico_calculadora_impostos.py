# ================================================================
# Exemplo Prático 30: Calculadora de Impostos com Interface Gráfica
#
# Este script cria uma interface gráfica simples para calcular o imposto devido
# com base em receitas, despesas e alíquota informadas pelo usuário.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - Tkinter já vem com o Python padrão.
# ================================================================
import tkinter as tk

def calcular_imposto():
    try:
        receitas = float(entrada_receitas.get())
        despesas = float(entrada_despesas.get())
        aliquota = float(entrada_aliquota.get()) / 100
        lucro = receitas - despesas
        imposto = lucro * aliquota
        label_resultado.config(text=f"Imposto devido: R$ {imposto:.2f}")
    except Exception as e:
        label_resultado.config(text=f"Erro: {str(e)}")

janela = tk.Tk()
janela.title("Calculadora de Impostos")

tk.Label(janela, text="Receitas (R$):").pack()
entrada_receitas = tk.Entry(janela)
entrada_receitas.pack()

tk.Label(janela, text="Despesas (R$):").pack()
entrada_despesas = tk.Entry(janela)
entrada_despesas.pack()

tk.Label(janela, text="Alíquota (%) :").pack()
entrada_aliquota = tk.Entry(janela)
entrada_aliquota.pack()

botao_calcular = tk.Button(janela, text="Calcular Imposto", command=calcular_imposto)
botao_calcular.pack()

label_resultado = tk.Label(janela, text="Imposto devido: R$ 0.00")
label_resultado.pack()

janela.mainloop()
