CAPA
```plaintext
 ___________________________________________________________
|                                                           |
|   ____        _   _                 _                     |
|  |  _ \ _   _| |_| |__   ___   ___ | | ___   _ _ __       |
|  | |_) | | | | __| '_ \ / _ \ / _ \| |/ / | | | '_ \      |
|  |  __/| |_| | |_| | | | (_) | (_) |   <| |_| | |_) |     |
|  |_|    \__, |\__|_| |_|\___/ \___/|_|\_\\__,_| .__/      |
|         |___/                                   |_|       |
|                                                           |
|   ⚖️   [ ]   🚚   🛒                                     |
|                                                           |
|   AUTOMAÇÃO DE TAREFAS PARA ESCRITÓRIOS                   |
|   COM PYTHON E INTELIGÊNCIA ARTIFICIAL                    |
|                                                           |
|   ┌───────────────────────────────────────────────┐       |
|   │ [  ][  ][  ][  ][  ][  ][  ][  ][  ][  ][  ]  │       |
|   │ [  ][  ][  ][  ][PY][  ][  ][  ][  ][  ][  ]  │       |
|   │ [  ][  ][  ][  ][  ][  ][  ][  ][  ][  ][  ]  │       |
|   └───────────────────────────────────────────────┘       |
|                                                           |
|   Christian Mulato - Programador de Computador            |
|___________________________________________________________|

# Automação de Tarefas para Escritórios com Python e Inteligência Artificial
**Christian Mulato**
Programador de Computador
**ISBN:** 978-65-0000-000-0
**Data de publicação:** Junho de 2025
**Direitos Autorais Reservados**
```

***

# Apresentação

Vivemos uma era em que a tecnologia e a inteligência artificial podem transformar o dia a dia dos escritórios, tornando tarefas repetitivas mais rápidas, precisas e menos cansativas. No entanto, muitos profissionais ainda gastam horas com atividades manuais que poderiam ser facilmente automatizadas.

Este livro nasceu da vontade de aproximar o poder da programação — especialmente do Python — e da inteligência artificial do ambiente de escritórios, sejam eles de advocacia, contabilidade, logística, vendas ou serviços para repartições públicas e órgãos governamentais. O objetivo é mostrar, de forma prática e acessível, como automatizar processos comuns, economizar tempo e aumentar a produtividade, mesmo para quem nunca programou antes.

Ao longo dos capítulos, você encontrará exemplos reais, projetos práticos e dicas para aplicar imediatamente no seu trabalho. Não é necessário ser um especialista em tecnologia: basta curiosidade e vontade de aprender.

O conteúdo está dividido em duas partes principais para facilitar o aprendizado e a aplicação dos conceitos:

- **Parte I – Fundamentos da Automação:** 
- Aqui você aprenderá os conceitos essenciais, desde a instalação do ambiente até a manipulação de arquivos, automação de e-mails, web scraping, criação de interfaces gráficas simples e introdução ao uso de inteligência artificial em tarefas do escritório.

- **Parte II – Casos Reais por Tipo de Escritório:**
-  Nesta parte, apresentamos exemplos práticos e projetos voltados para diferentes áreas, incluindo escritórios que prestam serviços para repartições públicas e órgãos governamentais, mostrando como adaptar as soluções para a sua realidade.

Espero que este material ajude você a transformar sua rotina profissional, abrindo portas para novas possibilidades e mostrando que a automação e a inteligência artificial estão ao alcance de todos.

Boa leitura e ótimos projetos!

**Christian Mulato**  
Programador de Computador

***

# Sumário

## Parte 1 – Fundamentos da Automação

1. [Introdução à automação com Python](#1-introdução-à-automação-com-python)  
2. [Instalação de ambiente: VS Code, Python, bibliotecas úteis](#2-instalação-de-ambiente-vs-code-python-e-bibliotecas-úteis)  
3. [Trabalhando com arquivos (Excel, CSV, PDF, Word)](#3-trabalhando-com-arquivos-excel-csv-pdf-word)  
4. [Automação de e-mails e notificações](#4-automação-de-e-mails-e-notificações) 
5. [Web scraping e automação de sites](#5-web-scraping-e-automação-de-sites) 
6. [Criação de interfaces gráficas simples (Tkinter ou PyWebIO)](#6-criação-de-interfaces-gráficas-simples-tkinter-ou-pywebio)  

## Parte 2 – Casos Reais por Tipo de Escritório

### Escritório de Advocacia

7. [Gerador automático de procurações e petições a partir de modelos](#7-gerador-automático-de-procurações-e-petições-a-partir-de-modelos)  
8. [Controle de prazos processuais (leitura de planilhas + envio de alertas por e-mail)](#8-controle-de-prazos-processuais-leitura-de-planilhas--envio-de-alertas-por-e-mail)  
9. [Consulta a sites de tribunais](#9-consulta-a-sites-de-tribunais)  

### Escritório de Contabilidade

10. [Leitura e consolidação de extratos bancários (CSV)](#10-leitura-e-consolidação-de-extratos-bancários-csv)  
11. [Geração automática de guias de impostos](#11-geraçã-automática-de-guias-de-impostos)  
12. [Envio automático de boletos por e-mail](#12-envio-automático-de-boletos-por-e-mail)  

### Escritório de Logística

13. [Leitura e geração de manifestos (XML, PDF)](#13-leitura-e-geraçã-de-manifestos-xml-pdf)  
14. [Roteirização com base em distância (API Google Maps ou OpenRoute)](#14-roteirização-com-base-em-distância-api-google-maps-ou-openroute)  
15. [Acompanhamento de entregas via planilhas atualizadas](#15-acompanhamento-de-entregas-via-planilhas-atualizadas)  

### E-commerce e Vendas Online

16. [Leitura de pedidos de marketplaces](#16-leitura-de-pedidos-de-marketplaces)  
17. [Atualização automática de estoque em Excel/ERP simples](#17-atualização-automática-de-estoque-em-excelerp-simples)  
18. [Envio de notas fiscais e respostas automáticas a clientes](#18-envio-de-notas-fiscais-e-respostas-automáticas-a-clientes)  

### Escritórios que prestam serviços para repartições públicas e órgãos governamentais

19. [Controle Automatizado de Protocolos em Repartição Pública](#19-controle-automatizado-de-protocolos-em-repartição-pública)
20. [Conversor de tabelas PDF → Excel](#20-conversor-de-tabelas-pdf--excel)
21. [Organizador de arquivos em pastas por cliente](#21-organizador-de-arquivos-em-pastas-por-cliente)
22. [Dashboard de pagamentos](#22-dashboard-de-pagamentos)

## Anexos

- Figuras e diagramas  
- Referências

## Códigos em Python

- Exemplo Prático 01 - Análise de vendas de produtos online (CSV)
   - `codes/01_exemplo_pratico_vendas_online.py`
- Exemplo Prático 02 - Geração de relatórios financeiros (Excel)
   - `codes/02_exemplo_pratico_relatorio_financeiro.py`
- Exemplo Prático 03 - Controle de inadimplência em aluguéis (Excel + E-mail)
   - `codes/03_exemplo_pratico_alugueis_com_email.py`
- Exemplo Prático 04 - Envio de e-mails automatizados (Python + SMTP)
   - `codes/04_exemplo_pratico_envio_email.py`
- Exemplo Prático 05 - Extração de texto de um PDF
   - `codes/05_exemplo_pratico_extracao_pdf.py`
- Exemplo Prático 06 - Leitura e modificação de um arquivo Word
   - `codes/06_exemplo_pratico_modificacao_word.py`
- Exemplo Prático 07 - Automação de Google Sheets
   - `codes/07_exemplo_pratico_automacao_excel.py`
- Exemplo Prático 08 - Envio de relatórios diários por e-mail
   - `codes/08_exemplo_pratico_envio_relatorio_diario.py`

***

# Introdução

A automação de tarefas rotineiras é uma das maiores oportunidades para escritórios que desejam ganhar produtividade, reduzir erros e liberar tempo para atividades mais estratégicas. Python, por ser uma linguagem acessível, poderosa e com vasta comunidade, tornou-se a escolha ideal para quem quer começar a automatizar processos no dia a dia profissional.

Este livro foi pensado para ser um guia prático, direto ao ponto, e acessível mesmo para quem nunca programou antes. O conteúdo está dividido em duas partes principais, para facilitar o aprendizado e a aplicação dos conceitos:

## Parte I – Fundamentos da Automação

Na primeira parte, você aprenderá os conceitos essenciais para começar a automatizar tarefas com Python. Aqui, abordamos desde a instalação do ambiente, passando pelo trabalho com arquivos, automação de e-mails, web scraping e até a criação de interfaces gráficas simples. O objetivo é construir uma base sólida, permitindo que você entenda como Python pode ser aplicado em diferentes situações do cotidiano de um escritório.

## Parte II – Casos Reais por Tipo de Escritório

Na segunda parte, apresentamos exemplos práticos e projetos voltados para áreas específicas: advocacia, contabilidade, logística e e-commerce. Cada capítulo traz situações reais enfrentadas nesses segmentos e mostra, passo a passo, como resolvê-las com Python. Assim, você poderá adaptar as soluções para a sua realidade, independentemente do ramo de atuação.

Ao final, você encontrará projetos práticos para consolidar o aprendizado e estimular a criatividade na busca por novas automações.

Esperamos que este livro seja um ponto de partida para transformar sua rotina profissional, tornando seu escritório mais inteligente, eficiente e preparado para os desafios

***

# Parte I – Fundamentos da Automação

Nesta primeira parte, você vai aprender os conceitos e ferramentas essenciais para começar a automatizar tarefas com Python no ambiente de escritório. O objetivo é construir uma base sólida, mesmo para quem nunca programou antes, mostrando passo a passo como instalar o ambiente, manipular arquivos, enviar e-mails automaticamente, coletar informações da internet e criar interfaces gráficas simples.

Ao dominar esses fundamentos, você estará preparado para aplicar a automação em diferentes áreas e desafios do dia a dia profissional.

***

# 1. Introdução à automação com Python

Automação é o processo de usar tecnologia para executar tarefas de forma automática, reduzindo a necessidade de intervenção manual. No contexto dos escritórios, isso significa transformar atividades repetitivas e demoradas em processos rápidos, precisos e eficientes.

Python é uma das linguagens de programação mais populares para automação devido à sua simplicidade, legibilidade e grande quantidade de bibliotecas prontas para uso. Com Python, é possível automatizar desde tarefas simples, como renomear arquivos em lote, até processos mais complexos, como gerar relatórios, enviar e-mails, manipular planilhas e interagir com sistemas web.

### Por que automatizar com Python?

- **Facilidade de aprendizado:** Python possui uma sintaxe clara e intuitiva, ideal para iniciantes.
- **Comunidade ativa:** Há muitos tutoriais, fóruns e exemplos disponíveis.
- **Bibliotecas poderosas:** Existem módulos prontos para manipular arquivos, enviar e-mails, acessar bancos de dados, trabalhar com planilhas, PDFs, web scraping, entre outros.
- **Versatilidade:** Python pode ser usado em diferentes sistemas operacionais e integrado a diversas ferramentas já utilizadas em escritórios.

### Exemplos de automação em escritórios

- Organizar e renomear arquivos automaticamente.
- Preencher e gerar documentos a partir de modelos.
- Ler e consolidar dados de planilhas.
- Enviar e-mails em massa ou com anexos personalizados.
- Coletar informações de sites automaticamente (web scraping).
- Gerar relatórios e dashboards de forma automática.

Ao longo deste livro, você verá como essas e outras tarefas podem ser automatizadas, tornando o seu dia a dia mais produtivo e eficiente. Vamos começar a jornada pela automação com Python!

***

# 2. Instalação de Ambiente: VS Code, Python e Bibliotecas Úteis

Antes de iniciar os exemplos práticos, é importante preparar o ambiente de desenvolvimento. Siga os passos abaixo:

## 2.1. Instale o Visual Studio Code (VS Code)

- Acesse: [https://code.visualstudio.com/](https://code.visualstudio.com/)
- Baixe e instale a versão adequada para o seu sistema operacional (Windows, macOS ou Linux).

## 2.2. Instale o Python

- Acesse: [https://www.python.org/downloads/](https://www.python.org/downloads/)
- Baixe a versão mais recente do Python 3.x.
- Durante a instalação, marque a opção **"Add Python to PATH"**.

## 2.3. Instale Bibliotecas Úteis

Abra o terminal do VS Code (atalho: `Ctrl + ``) e execute os comandos abaixo para instalar as principais bibliotecas que serão usadas nos exemplos:

```bash
pip install pandas openpyxl
```

- `pandas`: Manipulação de dados e leitura de planilhas Excel.
- `openpyxl`: Leitura e escrita de arquivos `.xlsx` (Excel).

## 2.4. (Opcional) Instale Extensões no VS Code

- **Python** (Microsoft): Suporte a sintaxe, depuração e execução de scripts Python.
- **Jupyter**: Para notebooks interativos, se desejar.

---

> Após esses passos, seu ambiente estará pronto para executar os exemplos práticos.

***

# 3. Trabalhando com arquivos (Excel, CSV, PDF, Word)

Nesta seção, você aprenderá a manipular diferentes tipos de arquivos comuns no ambiente de escritório, como Excel, CSV, PDF e Word. Essas habilidades são essenciais para automatizar tarefas que envolvem leitura, escrita e processamento de dados.

## 3.1. Arquivos CSV

Para ler e escrever arquivos CSV, use o módulo csv ou o pandas:
```python
import csv
# Lendo um arquivo CSV
with open('dados.csv', mode='r', encoding='utf-8') as file:
    leitor = csv.reader(file)
    for linha in leitor:
        print(linha)
# Escrevendo em um arquivo CSV
with open('saida.csv', mode='w', newline='', encoding='utf-8') as file:
    escritor = csv.writer(file)
    escritor.writerow(['Nome', 'Idade'])
    escritor.writerow(['João', 30])
    escritor.writerow(['Maria', 25])
```

## 3.2. Arquivos Excel

Para manipular arquivos Excel, use a biblioteca `pandas` ou `openpyxl`. Aqui está um exemplo com `pandas`:

```python
import pandas as pd
# Lendo um arquivo Excel
df = pd.read_excel('dados.xlsx', sheet_name='Planilha1')
print(df.head())  # Exibe as primeiras linhas do DataFrame
# Escrevendo em um arquivo Excel
df.to_excel('saida.xlsx', index=False, sheet_name='Resultados')
```

## 3.3. Arquivos PDF

Para ler arquivos PDF, você pode usar a biblioteca `PyPDF2`. Aqui está um exemplo simples:

```python
import PyPDF2
# Lendo um arquivo PDF
with open('documento.pdf', 'rb') as file:
    leitor = PyPDF2.PdfReader(file)
    for pagina in leitor.pages:
        print(pagina.extract_text())  # Extrai o texto de cada página
```

## 3.4. Arquivos Word

Para manipular arquivos Word, use a biblioteca `python-docx`. Veja como ler e escrever documentos:

```python
from docx import Document
# Lendo um arquivo Word
doc = Document('documento.docx')
for par in doc.paragraphs:
    print(par.text)  # Exibe o texto de cada parágrafo
# Escrevendo em um arquivo Word
doc_novo = Document()
doc_novo.add_heading('Título do Documento', level=1)
doc_novo.add_paragraph('Este é um parágrafo de exemplo.')
doc_novo.save('novo_documento.docx')  # Salva o novo documento
```

## 3.5. Configurando as dependências do ambiente Python

Para trabalhar com arquivos Excel, CSV, PDF e Word, você precisa instalar algumas bibliotecas extras. Siga os passos abaixo:

1. **Crie um ambiente virtual (opcional, mas recomendado):**
   No terminal, execute:
   ```
   python -m venv venv
   ```
   Ative o ambiente virtual:
   ```
   venv\Scripts\activate
   ```

2. **Instale as bibliotecas necessárias:**
   ```
   pip install pandas openpyxl pdfplumber python-docx
   ```

   - `pandas`: manipulação de dados (CSV, Excel)
   - `openpyxl`: leitura/escrita de arquivos Excel (.xlsx)
   - `pdfplumber`: leitura de arquivos PDF
   - `python-docx`: leitura/escrita de arquivos Word (.docx)

3. **(Opcional) Crie um arquivo `requirements.txt` para registrar as dependências:**
   ```
   pip freeze > requirements.txt
   ```

Assim, seu ambiente estará pronto para manipular arquivos desses tipos em Python. Você pode usar os exemplos acima como base para criar scripts que automatizam tarefas relacionadas a esses arquivos no seu dia a dia profissional.

***

## 3.6. Exercícios Práticos

Nesta seção, você encontrará exercícios práticos para aplicar os conceitos aprendidos sobre manipulação de arquivos. Esses exercícios são projetados para serem desafiadores e ajudarão a consolidar seu conhecimento em automação de tarefas comuns em escritórios.

***

### 3.6.1. Exercício Prático: Análise de Vendas de Produtos Online (CSV)

**História:**  
Você trabalha no setor de análise de dados de uma loja online. Recebeu um arquivo `vendas.csv` contendo o histórico de vendas do último mês. Cada linha do arquivo representa uma venda, com as seguintes colunas: `data`, `produto`, `quantidade`, `preco_unitario`.

**Desafio:**  
1. Leia o arquivo `vendas.csv` usando Python.
2. Calcule o total vendido (em reais) por produto.
3. Gere um novo arquivo `relatorio_vendas.csv` contendo duas colunas: `produto` e `total_vendido`.
4. (Opcional) Identifique qual produto teve o maior volume de vendas.

**Exemplo de entrada (`vendas.csv`):**
```
data,produto,quantidade,preco_unitario
2025-06-01,Mouse,2,50
2025-06-01,Teclado,1,120
2025-06-02,Mouse,1,50
2025-06-02,Monitor,1,900
```

**Código Python** para resolver o exercício:

```python
import os
import pandas as pd

# Caminhos dos arquivos
csv_path = r'C:\dev\python_escritorios\codes\vendas.csv'
relatorio_path = r'C:\dev\python_escritorios\codes\relatorio_vendas.csv'

# Apaga o relatório antigo, se existir
if os.path.exists(relatorio_path):
    os.remove(relatorio_path)

# Lê o arquivo de vendas
df = pd.read_csv(csv_path)

# Calcula o total vendido por produto
df['total'] = df['quantidade'] * df['preco_unitario']
df_relatorio = df.groupby('produto')['total'].sum().reset_index()
df_relatorio.rename(columns={'total': 'total_vendido'}, inplace=True)

# Salva o novo relatório
df_relatorio.to_csv(relatorio_path, index=False)

print("Relatório de vendas gerado com sucesso!")
```

**Resultado esperado (`relatorio_vendas.csv`)**:
```plaintext
produto,total_vendido
Mouse,150
Teclado,120
Monitor,900
```
**Conclusão:**

Com este exercício, você praticou a leitura de arquivos CSV, manipulação de dados com pandas e geração de relatórios. Essas habilidades são fundamentais para automatizar análises de vendas e outras tarefas relacionadas a dados em escritórios.


***

### 3.6.2. Exercício Prático: Automatizando Relatórios de Inadimplência em Aluguéis com Python e Excel

**História:**
Neste exercício, vamos aplicar a integração entre Python e Excel para resolver um problema comum em empresas de aluguel de imóveis: o controle de inadimplência.

**Cenário:**  
Uma empresa de aluguel de imóveis mantém uma planilha chamada `alugueis.xlsx` com os seguintes campos:

- Imóvel
- Inquilino
- Valor Mensal
- Data de Vencimento
- Pago (Sim/Não)

**Planilha de exemplo (`alugueis.xlsx`):**

```plaintext
|-----------|-------------------|--------------|--------------------|----------------|
| Imóvel    | Inquilino         | Valor Mensal | Data de Vencimento | Pago (Sim/Não) |
|-----------|-------------------|--------------|--------------------|----------------|
| Apto 101  | João da Silva     | 1800         | 2025-06-10         | Sim            |
| Casa 202  | Maria Souza       | 2500         | 2025-06-15         | Não            |
| Loja 303  | Pedro Lima        | 3200         | 2025-06-20         | Não            |
| Apto 104  | Ana Martins       | 2000         | 2025-06-12         | Sim            |
| Casa 205  | Bruno Carvalho    | 2700         | 2025-06-18         | Não            |
| Loja 307  | Carla Mendes      | 3500         | 2025-06-25         | Sim            |
| Apto 110  | Felipe Gonçalves  | 2100         | 2025-06-14         | Não            |
| Casa 208  | Luciana Ferreira  | 2600         | 2025-06-22         | Não            |
|-----------|-------------------|--------------|--------------------|----------------|
```

**Desafio:**  
Crie um script Python que:

1. Leia a planilha `alugueis.xlsx`.
2. Filtre os registros de imóveis cujo aluguel ainda não foi pago.
3. Gere um novo arquivo Excel chamado `relatorio_inadimplentes.xlsx` apenas com os inadimplentes.
4. Simule o envio de um e-mail para cada inquilino inadimplente, informando o valor devido e a data de vencimento.

**Exemplo de Código Python:**

```python
import pandas as pd
import os

relatorio_path = r'C:\dev\python_escritorios\codes\relatorio_inadimplentes.xlsx'
# 1. Importar bibliotecas
# Importar bibliotecas necessárias
# Apagar o relatório anterior, se existir
if os.path.exists(relatorio_path):
    os.remove(relatorio_path)

# 2. Ler a planilha de aluguéis
df = pd.read_excel(r'C:\dev\python_escritorios\codes\alugueis.xlsx')

# 3. Filtrar inadimplentes
inadimplentes = df[df['Pago (Sim/Não)'] == 'Não']

# 4. Salvar relatório dos inadimplentes
inadimplentes.to_excel(r'C:\dev\python_escritorios\codes\relatorio_inadimplentes.xlsx', index=False)

# 5. Simular envio de e-mails
for _, row in inadimplentes.iterrows():
    print(f"Enviando e-mail para {row['Inquilino']}:")
    print(f"Prezado(a) {row['Inquilino']}, seu aluguel do imóvel {row['Imóvel']} no valor de R$ {row['Valor Mensal']} está em aberto. Vencimento: {row['Data de Vencimento']}.")
    print("------------------------------\n")

# 6. Calcular total de aluguéis em aberto
total_inadimplentes = inadimplentes['Valor Mensal'].sum()
print(f"Total de aluguéis em aberto: R$ {total_inadimplentes}")

# 7. Calcular total de aluguéis pagos
total_pagos = df[df['Pago (Sim/Não)'] == 'Sim']['Valor Mensal'].sum()
print(f"Total de aluguéis pagos: R$ {total_pagos}")

# 8. Calcular total geral de aluguéis
total_geral = df['Valor Mensal'].sum()
print(f"Total geral de aluguéis: R$ {total_geral}")

# 9. Calcular porcentagem de inadimplência
porcentagem_inadimplencia = (total_inadimplentes / total_geral) * 100 if total_geral > 0 else 0
print(f"Porcentagem de inadimplência: {porcentagem_inadimplencia:.2f}%")

# 10. Exibir relatório final
print("\nRelatório Final:")
print(f"Total de aluguéis em aberto: R$ {total_inadimplentes}")
print(f"Total de aluguéis pagos: R$ {total_pagos}")
print(f"Total geral de aluguéis: R$ {total_geral}")

# 11. Salvar relatório final
relatorio_final = {
    'Total Aluguéis em Aberto': [total_inadimplentes],
    'Total Aluguéis Pagos': [total_pagos],
    'Total Geral de Aluguéis': [total_geral],
    'Porcentagem de Inadimplência': [porcentagem_inadimplencia]
}
relatorio_df = pd.DataFrame(relatorio_final)
relatorio_df.to_excel(r'C:\dev\python_escritorios\codes\relatorio_final.xlsx', index=False)

# 12. Exibir mensagem de conclusão
print("Relatório final salvo com sucesso em 'relatorio_final.xlsx'.")

# 13. Fim do script
print("Processamento concluído.")
```

### Adendo - Exercício Prático: Envio Real de E-mails para Inquilinos Inadimplentes

Além de gerar o relatório de inadimplentes, é possível automatizar o envio de e-mails reais para cada inquilino com aluguel em aberto utilizando Python.

**Como fazer:**

1. **Adicione uma coluna "E-mail"** na planilha `alugueis.xlsx` com o endereço de e-mail de cada inquilino.

2. **Utilize a biblioteca `smtplib`** do Python para enviar e-mails automáticos. É recomendado utilizar uma conta de e-mail dedicada e, no caso do Gmail, gerar uma senha de aplicativo para maior segurança.

3. **Exemplo de código Python:**  
O código completo para geração do relatório de inadimplentes e envio real de e-mails está disponível no arquivo:

**Código Python para envio de e-mails:**

`codes/03_exemplo_pratico_alugueis_com_email.py`

4. **Exemplo de Planilha de Aluguéis com E-mail dos Inquilinos**

```plaintext
|-----------|-------------------|--------------|--------------------|----------------|----------------------|
| Imóvel    | Inquilino         | Valor Mensal | Data de Vencimento | Pago (Sim/Não) | E-mail               |
|-----------|-------------------|--------------|--------------------|----------------|----------------------|
| Apto 101  | João da Silva     | 1800         | 2025-06-10         | Sim            | joao@email.com       |
| Casa 202  | Maria Souza       | 2500         | 2025-06-15         | Não            | maria@email.com      |
| Loja 303  | Pedro Lima        | 3200         | 2025-06-20         | Não            | pedro@email.com      |
| Apto 104  | Ana Martins       | 2000         | 2025-06-12         | Sim            | ana@email.com        |
| Casa 205  | Bruno Carvalho    | 2700         | 2025-06-18         | Não            | bruno@email.com      |
| Loja 307  | Carla Mendes      | 3500         | 2025-06-25         | Sim            | carla@email.com      |
| Apto 110  | Felipe Gonçalves  | 2100         | 2025-06-14         | Não            | felipe@email.com     |
| Casa 208  | Luciana Ferreira  | 2600         | 2025-06-22         | Não            | luciana@email.com    |
|-----------|-------------------|--------------|--------------------|----------------|----------------------|
```

**Atenção:**  
- Nunca compartilhe senhas em código público.
- Teste o envio com poucos destinatários antes de usar em produção.
- Certifique-se de que o envio de e-mails está de acordo com as políticas da sua empresa e com a legislação vigente (LGPD).

Com essa automação, a empresa pode agilizar a comunicação com os inquilinos inadimplentes, tornando o processo mais eficiente e reduzindo o tempo gasto com cobranças manuais.

**Conclusão:**

Com este exercício, você praticou a leitura e escrita de arquivos Excel, filtragem de dados com pandas. Essas habilidades são essenciais para automatizar o controle de inadimplência em empresas de aluguel de imóveis, tornando o processo mais eficiente e organizado.

***

### 3.6.3. Exercício Prático: Extração de Texto de um PDF

**História Fictícia:**  
A repartição pública “Prefeitura Municipal de Campo Largo, do Estado do Paraná” recebe diariamente diversos documentos em PDF, como ofícios, requerimentos e certidões. A servidora Ana precisa extrair rapidamente o texto de um ofício em PDF para copiar o conteúdo para um sistema interno, mas fazer isso manualmente toma muito tempo.

**Desafio:**  
Ajude Ana a automatizar esse processo criando um script Python que:

1. Leia um arquivo PDF chamado `oficio_exemplo.pdf` localizado na pasta `codes`.
2. Extraia todo o texto do documento.
3. Salve o texto extraído em um arquivo chamado `oficio_extraido.txt` na mesma pasta.
4. Mostre na tela as primeiras linhas do texto extraído.

**Dica:**  
Utilize a biblioteca `PyPDF2` para a extração de texto de arquivos PDF.

**Código Python**:

```python
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
```

**Observação:**  
Certifique-se de instalar a biblioteca com `pip install PyPDF2` antes de executar o script.

**Resultado esperado:**

```plaintext
OFÍCIO Nº 123/2025 – GAB/PMNE  
 
Campo Largo, 17 de junho de 2025.  
 
Ao   
Senhor Diretor do Departamento de Recursos Humanos   
Prefeitura Municipal de Campo Largo  
 
Assunto: Solicitação de atualização cadastral  
 
Prezado Senhor,  
 
Solicito, por meio deste ofício, a gentileza de proceder à atualização dos dados 
cadastrais do servidor João da Silva, matrícula 4567, lotado no setor de Protocolo. 
Informo que houve alteração de endereço residencial e número de telefone para 
contato.  
 
Segue em anexo a documentação comprobatória.  
 
Atenciosamente,  
 
Maria Souza   
Chefe de Gabinete   
Prefeitura Municipal de Campo Largo
```

**Conclusão:**

Com este exercício, você praticou a leitura de arquivos PDF e a extração de texto usando Python. Essa habilidade é útil para automatizar a coleta de informações de documentos oficiais, economizando tempo e aumentando a eficiência no trabalho com repartições públicas.


***

### 3.6.4. Exercício Prático: Leitura e Modificação de um Arquivo Word

**História Fictícia:**  
O escritório de advocacia “Silva & Associados” recebe frequentemente minutas de contratos em formato Word, enviadas por clientes e parceiros. A advogada Júlia precisa revisar rapidamente esses contratos para garantir que todos contenham uma cláusula padrão de confidencialidade. Para agilizar o processo, ela deseja um script Python que leia o arquivo Word do contrato, verifique se a cláusula está presente e, caso não esteja, adicione automaticamente a cláusula ao final do documento.

**Desafio:**  
Crie um script Python que:

1. Leia um arquivo Word chamado `contrato_exemplo.docx` localizado na pasta `codes`.
2. Verifique se o texto da cláusula de confidencialidade está presente.
3. Caso não esteja, adicione a seguinte cláusula ao final do documento:

   > **Cláusula de Confidencialidade:** As partes se comprometem a manter em sigilo todas as informações trocadas em razão deste contrato, não podendo divulgá-las a terceiros sem prévia autorização por escrito.

4. Salve o documento modificado como `contrato_confidencial.docx` na mesma pasta.

**Dica:**  
Utilize a biblioteca `python-docx` para manipulação de arquivos Word.

**Código Python**:

```python
from docx import Document
# Caminho do arquivo Word
doc_path = r'C:\dev\python_escritorios\codes\contrato_exemplo.docx'
doc_novo_path = r'C:\dev\python_escritorios\codes\contrato_confidencial.docx' 
# Cláusula de confidencialidade
clausa_confidencialidade = (
    "Cláusula de Confidencialidade: As partes se comprometem a manter em sigilo "
    "todas as informações trocadas em razão deste contrato, não podendo divulgá-las "
    "a terceiros sem prévia autorização por escrito."
)
# Abrir o documento Word
doc = Document(doc_path)
# Verificar se a cláusula já existe
clausula_presente = any(clausa_confidencialidade in par.text for par in doc.paragraphs)
if not clausula_presente:
    # Adicionar a cláusula de confidencialidade ao final do documento
    doc.add_paragraph(clausa_confidencialidade)
    print("Cláusula de confidencialidade adicionada ao contrato.")
else:
    print("Cláusula de confidencialidade já está presente no contrato.")
# Salvar o documento modificado
doc.save(doc_novo_path)
# Exibir mensagem de conclusão
print(f"Documento modificado salvo como '{doc_novo_path}'")
```

**Resultado esperado:**

```plaintext
Cláusula de confidencialidade adicionada ao contrato.
Documento modificado salvo como 'C:\dev\python_escritorios\codes\contrato_confidencial.docx'
```

**Conclusão:**

Com este exercício, você praticou a leitura e modificação de arquivos Word usando Python. Essa habilidade é útil para automatizar a revisão de documentos legais, garantindo que cláusulas importantes estejam sempre presentes, economizando tempo e aumentando a eficiência no trabalho de advogados e escritórios de advocacia.

***

# 4. Automação de E-mails e Notificações

Nesta seção, você aprenderá a automatizar o envio de e-mails e notificações usando Python. Essa habilidade é essencial para manter a comunicação eficiente em escritórios, seja para enviar lembretes, relatórios ou avisos importantes.

## 4.1. Enviando E-mails com Python

Para enviar e-mails com Python, você pode usar a biblioteca `smtplib`, que permite enviar mensagens através do protocolo SMTP. Abaixo está um exemplo básico de como enviar um e-mail:

```python
import smtplib
from email.mime.text import MIMEText
# Configurações do servidor SMTP
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_user = 'seu_email@gmail.com'
smtp_password = 'sua_senha'
# Função para enviar e-mail
def enviar_email(destinatario, assunto, corpo):
    # Criar o objeto de mensagem
    msg = MIMEText(corpo)
    msg['Subject'] = assunto
    msg['From'] = smtp_user
    msg['To'] = destinatario
    # Enviar o e-mail
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)    
# Exemplo de uso
enviar_email('destinatario@example.com', 'Assunto do E-mail', 'Corpo do e-mail')
```

## 4.2. Enviando E-mails com Anexos

Para enviar e-mails com anexos, você pode usar a biblioteca `email` para criar mensagens mais complexas. Veja o exemplo abaixo:
```python
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
# Configurações do servidor SMTP
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_user = 'seu_email@gmail.com'
smtp_password = 'sua_senha'
# Função para enviar e-mail com anexo
def enviar_email_com_anexo(destinatario, assunto, corpo, caminho_anexo):
    # Criar o objeto de mensagem
    msg = MIMEMultipart()
    msg['Subject'] = assunto
    msg['From'] = smtp_user
    msg['To'] = destinatario
    # Adicionar o corpo do e-mail
    msg.attach(MIMEText(corpo, 'plain'))
    # Adicionar o anexo
    with open(caminho_anexo, 'rb') as anexo:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(anexo.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f'attachment; filename={caminho_anexo.split("/")[-1]}')
        msg.attach(parte)
    # Enviar o e-mail
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)
# Exemplo de uso
enviar_email_com_anexo(
    'destinatario@example.com',
    'Assunto do E-mail com Anexo',
    'Corpo do e-mail com anexo',
    'caminho/para/o/anexo.txt'
)
```

## 4.3. Enviando Notificações por E-mail

Para enviar notificações por e-mail, você pode usar a função `enviar_email` que criamos anteriormente. Veja um exemplo de como enviar um lembrete de reunião:

```python
enviar_email(
    'destinatario@example.com',
    'Lembrete de Reunião',
    'Este é um lembrete de que você tem uma reunião agendada para amanhã às 10h.'
)
```

## 4.4. Configurando as Dependências do Ambiente Python

Para que os exemplos de envio de e-mails funcionem, você precisa ter o Python instalado em seu ambiente, bem como as bibliotecas necessárias. Você pode instalar as bibliotecas usando o gerenciador de pacotes `pip`. Execute o seguinte comando no terminal:

```bash
pip install secure-smtplib email
```

## 4.5. Exercícios Práticos

Nesta seção, você encontrará exercícios práticos para aplicar os conceitos aprendidos sobre automação de e-mails e notificações. Esses exercícios são projetados para serem desafiadores e ajudarão a consolidar seu conhecimento em automação de comunicação por e-mail.


### 4.5.1. Exercício Prático: Envio de Relatórios Diários por E-mail

**História:**
A empresa "Tech Solutions" precisa enviar relatórios diários de vendas para a equipe de gerência. O relatório é gerado em um arquivo Excel chamado `relatorio_vendas.xlsx`, que contém as vendas do dia anterior. A gerente Ana deseja automatizar o envio desse relatório por e-mail todas as manhãs.

**Desafio:**
Crie um script Python que:
1. Leia o arquivo `relatorio_vendas.xlsx` localizado na pasta `codes`.
2. Envie o arquivo como anexo para o e-mail da gerente Ana, cujo endereço é `ana@techsolutions.com`.
3. O assunto do e-mail deve ser "Relatório de Vendas - [Data Atual]".
4. O corpo do e-mail deve conter uma breve mensagem: "Prezada Ana, segue em anexo o relatório de vendas do dia anterior."
5. Certifique-se de que o e-mail seja enviado usando uma conta de e-mail válida (por exemplo, Gmail).
6. (Opcional) Configure o script para ser executado automaticamente todos os dias às 8h da manhã usando o agendador de tarefas do sistema operacional (Windows Task Scheduler ou cron no Linux).

**Código Python:**

```python
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from datetime import datetime
import os

# Configurações do servidor SMTP
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_user = 'seu_email@gmail.com'
smtp_password = 'sua_senha'

# Caminho do arquivo Excel
excel_path = r'C:\dev\python_escritorios\codes\relatorio_vendas.xlsx'

# 1. Geração de dados fictícios de vendas
fake = Faker('pt_BR')
dados_vendas = []
for _ in range(10):  # Gerar 10 registros fictícios
    dados_vendas.append({
        'data': fake.date_this_month(),
        'produto': fake.word(),
        'quantidade': fake.random_int(min=1, max=10),
        'preco_unitario': fake.random_int(min=50, max=500)
    })

# Criar DataFrame
df_vendas = pd.DataFrame(dados_vendas)

# 2. Salvando dados em um arquivo Excel (.xlsx)
df_vendas.to_excel(excel_path, index=False, sheet_name='Vendas')

# 3. Exibindo dados no console para cópia manual para o Excel
print("Dados de Vendas Gerados:")
print(df_vendas)

# 4. Envio do relatório como anexo por e-mail
data_atual = datetime.now().strftime('%d/%m/%Y')
assunto = f'Relatório de Vendas - {data_atual}'
corpo = 'Prezada Ana, segue em anexo o relatório de vendas do dia.'

# Função para enviar e-mail com anexo
def enviar_email_com_anexo(destinatario, assunto, corpo, caminho_anexo):
    # Criar o objeto de mensagem
    msg = MIMEMultipart()
    msg['Subject'] = assunto
    msg['From'] = smtp_user
    msg['To'] = destinatario
    # Adicionar o corpo do e-mail
    msg.attach(MIMEText(corpo, 'plain'))
    # Adicionar o anexo
    with open(caminho_anexo, 'rb') as anexo:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(anexo.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f'attachment; filename={caminho_anexo.split("/")[-1]}')
        msg.attach(parte)
    # Enviar o e-mail
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

# Enviar o e-mail com o relatório
enviar_email_com_anexo(
    'ana@techsolutions.com',
    assunto,
    corpo,
    excel_path
)
```

**Resultado esperado:**
```plaintext
Arquivo de vendas gerado em: C:\dev\python_escritorios\codes\relatorio_vendas.xlsx
Copie e cole os dados abaixo no Excel, se desejar:
Data	Produto	Quantidade	Valor Total
25/10/2023	Notebook	5	15000.00
E-mail enviado com sucesso!
```

**Observação:**
Certifique-se de substituir `seu_email@gmail.com` e `sua_senha` pelas suas credenciais de e-mail.

**Atenção:**
- Nunca compartilhe suas credenciais de e-mail em código público.
- Use uma senha de aplicativo se estiver usando o Gmail, pois o Google pode bloquear tentativas de login de aplicativos menos seguros.
- Considere usar OAuth2 para uma abordagem mais segura.

**Conclusão:**

Com este exercício, você praticou o envio de e-mails com anexos usando Python. Essa habilidade é essencial para automatizar a comunicação de relatórios e documentos importantes em ambientes corporativos, economizando tempo e aumentando a eficiência na troca de informações.

***

### 4.5.2. Exercício Prático: Envio Automático de Relatório Diário de Vendas

**História:**  
A empresa "Tech Solutions" deseja automatizar o envio diário de um relatório de vendas por e-mail para a equipe de gerência. O relatório deve conter dados fictícios de vendas, ser salvo em Excel e enviado como anexo para a gerente Ana.

**Desafio:**  
Automatizar o processo de geração e envio do relatório diário de vendas, seguindo os passos:

1. **Geração de Dados Fictícios de Vendas:**  
   - Gerar 10 registros fictícios de vendas usando a biblioteca `Faker`.
   - Cada registro deve conter: data, produto, quantidade e preço unitário.

2. **Salvando Dados em um Arquivo Excel (.xlsx):**  
   - Salvar os dados em um arquivo Excel chamado `relatorio_vendas.xlsx` usando o `pandas`.

3. **Exibindo Dados no Console para Cópia Manual para o Excel:**  
   - Exibir os dados gerados no console para facilitar a cópia.

4. **Envio do Relatório como Anexo por E-mail:**  
   - Enviar o arquivo Excel como anexo para o e-mail da gerente Ana (`ana@techsolutions.com`) usando `smtplib`.

**Código Python:**

```python
import pandas as pd
from faker import Faker
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import os

# Configurações do servidor SMTP
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_user = 'seu_email@gmail.com'
smtp_password = 'sua_senha'

# Caminho do arquivo Excel
excel_path = r'C:\dev\python_escritorios\codes\relatorio_vendas.xlsx'

# 1. Geração de dados fictícios de vendas
fake = Faker('pt_BR')
dados_vendas = []
for _ in range(10):
    dados_vendas.append({
        'data': fake.date_this_month(),
        'produto': fake.word(),
        'quantidade': fake.random_int(min=1, max=10),
        'preco_unitario': fake.random_int(min=50, max=500)
    })

# 2. Criar DataFrame e salvar em Excel
df_vendas = pd.DataFrame(dados_vendas)
df_vendas.to_excel(excel_path, index=False, sheet_name='Vendas')

# 3. Exibir dados no console
print("Dados de Vendas Gerados:")
print(df_vendas)

# 4. Envio do relatório como anexo por e-mail
data_atual = datetime.now().strftime('%d/%m/%Y')
assunto = f'Relatório de Vendas - {data_atual}'
corpo = 'Prezada Ana, segue em anexo o relatório de vendas do dia.'

def enviar_email_com_anexo(destinatario, assunto, corpo, caminho_anexo):
    msg = MIMEMultipart()
    msg['Subject'] = assunto
    msg['From'] = smtp_user
    msg['To'] = destinatario
    msg.attach(MIMEText(corpo, 'plain'))
    with open(caminho_anexo, 'rb') as anexo:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(anexo.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f'attachment; filename={caminho_anexo.split("/")[-1]}')
        msg.attach(parte)
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

enviar_email_com_anexo(
    'ana@techsolutions.com',
    assunto,
    corpo,
    excel_path
)

print("E-mail enviado com sucesso!")
```

**Resultado esperado:**
```plaintext
Dados de Vendas Gerados:
           data   produto  quantidade  preco_unitario
0   2025-06-01  exemplo1           3             120
1   2025-06-02  exemplo2           7             350
... (mais linhas) ...
E-mail enviado com sucesso!
```

**Observações:**
- Substitua `seu_email@gmail.com` e `sua_senha` pelas suas credenciais.
- Instale as dependências necessárias com:  
  `pip install pandas faker openpyxl`
- Nunca compartilhe suas credenciais em código público.
- Considere usar senha de aplicativo ou OAuth2 para maior segurança.

**Conclusão:**  
Com este exemplo, você praticou a geração de dados fictícios, manipulação de arquivos Excel, exibição de dados no console e envio automático de e-mails com anexo usando Python.

***

## 4.6. Considerações de Segurança (OAuth2, aplicativos menos seguros)

Ao automatizar o envio de e-mails, é importante considerar as implicações de segurança. O uso de senhas em texto claro, como mostrado nos exemplos, não é recomendado para ambientes de produção.

## OAuth2

Uma abordagem mais segura é usar OAuth2 para autenticação. O OAuth2 é um protocolo de autorização que permite que aplicativos de terceiros acessem suas informações sem precisar compartilhar suas senhas.

### Vantagens do OAuth2

- **Segurança:** Não é necessário armazenar senhas em texto claro.
- **Controle:** É possível revogar o acesso a qualquer momento.
- **Escopo:** É possível limitar o acesso a apenas determinadas informações.

### Como Usar OAuth2 com Python

Para usar OAuth2 com Python, você pode usar a biblioteca `oauth2client`. Abaixo está um exemplo básico de como configurar o OAuth2:

```python
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# Configurar o escopo e as credenciais
escopo = ['https://www.googleapis.com/auth/spreadsheets']
credenciais = ServiceAccountCredentials.from_json_keyfile_name('caminho/para/credenciais.json', escopo)

# Autenticar e acessar o Google Sheets
cliente = gspread.authorize(credenciais)
planilha = cliente.open('Nome da Planilha').sheet1

# Ler dados da planilha
dados = planilha.get_all_records()
print(dados)
```

## Aplicativos Menos Seguros

Se você estiver usando o Gmail, é possível que precise permitir o acesso de "aplicativos menos seguros" para que o envio de e-mails funcione. No entanto, essa opção reduz a segurança da sua conta e não é recomendada.

### Como Ativar

1. Acesse sua conta do Google.
2. Vá para "Segurança".
3. Na seção "Acesso a app menos seguro", ative a opção "Permitir aplicativos menos seguros".

### Atenção

- Essa opção pode não estar disponível para contas do Google Workspace (antigo G Suite).
- O Google pode bloquear o acesso de aplicativos menos seguros a qualquer momento, afetando a funcionalidade do seu aplicativo.

## Conclusão

Neste capítulo, você aprendeu sobre considerações de segurança ao automatizar o envio de e-mails. A utilização de OAuth2 é a abordagem recomendada para garantir a segurança das suas credenciais e informações.

***

