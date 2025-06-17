# ================================================================
# Exemplo Prático 12: Interface Gráfica Simples com Tkinter
#
# Capítulo: 6. Criação de interfaces gráficas simples (Tkinter ou PyWebIO)
# Objetivo: Demonstrar como criar uma interface gráfica básica para automação,
#           útil para escritórios que desejam facilitar o uso de scripts por
#           usuários não técnicos, como calculadoras, formulários, etc.
#
# Área de aplicação: Escritórios de contabilidade, advocacia, logística, vendas.
#
# Autor: Christian Mulato
# Data: Junho de 2025
#
# ATENÇÃO:
# - Tkinter já vem com o Python padrão.
# ================================================================
import tkinter as tk

# Função para realizar a operação
def calcular(operacao):
    try:
        num1 = float(entrada_num1.get())
        num2 = float(entrada_num2.get())
        if operacao == "adição":
            resultado = num1 + num2
        elif operacao == "subtração":
            resultado = num1 - num2
        elif operacao == "multiplicação":
            resultado = num1 * num2
        elif operacao == "divisão":
            resultado = num1 / num2
        label_resultado.config(text=f"Resultado: {resultado}")
    except Exception as e:
        label_resultado.config(text=f"Erro: {str(e)}")

# Criar a janela principal
janela = tk.Tk()
janela.title("Calculadora Simples")

# Criar campos de entrada
label_num1 = tk.Label(janela, text="Número 1:")
label_num1.pack()
entrada_num1 = tk.Entry(janela)
entrada_num1.pack()

label_num2 = tk.Label(janela, text="Número 2:")
label_num2.pack()
entrada_num2 = tk.Entry(janela)
entrada_num2.pack()

# Criar botões para as operações
botao_adicao = tk.Button(janela, text="+", command=lambda: calcular("adição"))
botao_adicao.pack()

botao_subtracao = tk.Button(janela, text="-", command=lambda: calcular("subtração"))
botao_subtracao.pack()

botao_multiplicacao = tk.Button(janela, text="*", command=lambda: calcular("multiplicação"))
botao_multiplicacao.pack()

botao_divisao = tk.Button(janela, text="/", command=lambda: calcular("divisão"))
botao_divisao.pack()

# Criar rótulo para exibir o resultado
label_resultado = tk.Label(janela, text="Resultado:")
label_resultado.pack()

# Iniciar o loop da interface
janela.mainloop()
