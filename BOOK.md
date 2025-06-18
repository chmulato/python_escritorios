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

## Parte I – Fundamentos da Automação

1. [Introdução à automação com Python](#1-introdução-à-automação-com-python)
2. [Instalação de Ambiente: VS Code, Python e Bibliotecas Úteis](#2-instalação-de-ambiente-vs-code-python-e-bibliotecas-úteis)
3. [Trabalhando com arquivos (Excel, CSV, PDF, Word)](#3-trabalhando-com-arquivos-excel-csv-pdf-word)
4. [Automação de e-mails e notificações](#4-automação-de-e-mails-e-notificações)
5. [Web scraping e automação de sites](#5-web-scraping-e-automação-de-sites)
6. [Criação de interfaces gráficas simples (Tkinter ou PyWebIO)](#6-criação-de-interfaces-gráficas-simples-tkinter-ou-pywebio)

## Parte II – Casos Reais por Tipo de Escritório

### Escritório de Advocacia

7. [Gerador automático de procurações e petições a partir de modelos](#7-gerador-automático-de-procurações-e-petições-a-partir-de-modelos)
8. [Controle de prazos processuais (leitura de planilhas + envio de alertas por e-mail)](#8-controle-de-prazos-processuais-leitura-de-planilhas--envio-de-alertas-por-e-mail)
9. [Consulta a sites de tribunais](#9-consulta-a-sites-de-tribunais)

### Escritório de Contabilidade

10. [Leitura e consolidação de extratos bancários (CSV)](#10-leitura-e-consolidação-de-extratos-bancários-csv)
11. [Geração automática de guias de impostos](#11-gera%C3%A7%C3%A3o-autom%C3%A1tica-de-guias-de-impostos)
12. [Envio automático de boletos por e-mail](#12-envio-autom%C3%A1tico-de-boletos-por-e-mail)

### Escritório de Logística

13. [Leitura e geração de manifestos (XML, PDF)](#13-leitura-e-gera%C3%A7%C3%A3o-de-manifestos-xml-pdf)
14. [Roteirização com base em distância (API Google Maps ou OpenRoute)](#14-roteiriza%C3%A7%C3%A3o-com-base-em-dist%C3%A2ncia-api-google-maps-ou-openroute)
15. [Acompanhamento de entregas via planilhas atualizadas](#15-acompanhamento-de-entregas-via-planilhas-atualizadas)

### E-commerce e Vendas Online

16. [Leitura de pedidos de marketplaces](#16-leitura-de-pedidos-de-marketplaces)
17. [Atualização automática de estoque em Excel/ERP simples](#17-atualiza%C3%A7%C3%A3o-autom%C3%A1tica-de-estoque-em-excelerp-simples)
18. [Envio de notas fiscais e respostas automáticas a clientes](#18-envio-de-notas-fiscais-e-respostas-autom%C3%A1ticas-a-clientes)

### Escritórios que prestam serviços para repartições públicas e órgãos governamentais

19. [Controle Automatizado de Protocolos em Repartição Pública](#19-controle-automatizado-de-protocolos-em-reparti%C3%A7%C3%A3o-p%C3%BAblica)
20. [Conversor de tabelas PDF → Excel](#20-conversor-de-tabelas-pdf--excel)
21. [Organizador de arquivos em pastas por cliente](#21-organizador-de-arquivos-em-pastas-por-cliente)
22. [Dashboard de pagamentos](#22-dashboard-de-pagamentos)

---

23. [Conclusão](#23-conclusão)
24. [Agradecimentos](#24-agradecimentos)
25. [Referências bibliográficas (formato ABNT)](#25-referências-bibliográficas-formato-abnt)
26. [Adendo: códigos em Python](#26-adendo-códigos-em-python)

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

# Caminho do arquivo Excel
excel_path = r'C:\dev\python_escritorios\codes\relatorio_vendas.xlsx'

# Ler o arquivo Excel existente
if os.path.exists(excel_path):
    df_vendas = pd.read_excel(excel_path)
    print(f"Arquivo de vendas encontrado em: {excel_path}")
    # Exibir os dados de vendas no console para copiar e colar no Excel
    print("\nCopie e cole os dados abaixo no Excel, se desejar:")
    print(df_vendas.to_csv(sep='\t', index=False))
else:
    print(f"Arquivo {excel_path} não encontrado. Certifique-se de que a planilha existe com os dados prontos.")
    exit(1)

# Enviar o e-mail com o relatório
data_atual = datetime.now().strftime('%d/%m/%Y')
assunto = f'Relatório de Vendas - {data_atual}'
corpo = 'Prezada Ana, segue em anexo o relatório de vendas do dia anterior.'
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
Arquivo de vendas gerado em: C:\dev\python_escritorios\codes\relatorio_vendas.xlsx
Copie e cole os dados abaixo no Excel, se desejar:
Data	Produto	Quantidade	Valor Total
25/10/2023	Notebook	5	15000.00
E-mail enviado com sucesso!
```

**Observações:**
- Substitua `seu_email@gmail.com` e `sua_senha` pelas suas credenciais.
- Instale as dependências necessárias com:  
  `pip install pandas faker openpyxl`
- Nunca compartilhe suas credenciais em código público.
- Considere usar senha de aplicativo ou OAuth2 para maior segurança.

**Conclusão:**

Com este exercício, você praticou o envio de e-mails com anexos usando Python. Essa habilidade é essencial para automatizar a comunicação de relatórios e documentos importantes em ambientes corporativos, economizando tempo e aumentando a eficiência na troca de informações.

***

### 4.5.2. Exercício Prático: Geração e Envio Automático de Relatórios Diários de Vendas (Múltiplos Dias)

**História:**  
A empresa "Tech Solutions" deseja automatizar o envio de relatórios diários de vendas, gerando arquivos separados para cada dia e enviando todos juntos em um único e-mail para a gerente Ana.

**Desafio:**  
Automatizar o processo de geração de dados fictícios de vendas para vários dias, salvar um arquivo Excel para cada dia e enviar todos os arquivos como anexos em um único e-mail.

1. **Geração de Dados Fictícios de Vendas para Múltiplos Dias:**  
   - Gerar registros de vendas para vários dias usando a biblioteca `Faker`.
   - Cada arquivo deve conter as vendas de um dia específico.

2. **Salvando Dados em Arquivos Excel Diários:**  
   - Salvar cada dia em um arquivo Excel separado, nomeando conforme a data.

3. **Exibindo Dados no Console:**  
   - Exibir os dados de cada dia no console para conferência.

4. **Envio de Todos os Relatórios em um Único E-mail:**  
   - Enviar todos os arquivos gerados como anexos em um único e-mail para a gerente Ana (`ana@techsolutions.com`).

**Código Python:**  
Veja o arquivo `codes/08_exemplo_pratico_envio_relatorio_diario.py` para o exemplo completo.

**Resumo do Aprendizado:**
- Geração de dados para múltiplos dias.
- Manipulação e criação dinâmica de múltiplos arquivos Excel.
- Envio de múltiplos anexos em um único e-mail.
- Organização de arquivos em pastas.

***

## 4.6. Considerações de Segurança (OAuth2, aplicativos menos seguros)

Ao automatizar o envio de e-mails, é importante considerar as implicações de segurança. O uso de senhas em texto claro, como mostrado nos exemplos, não é recomendado para ambientes de produção.

### 4.6.1. OAuth2

Uma abordagem mais segura é usar OAuth2 para autenticação. O OAuth2 é um protocolo de autorização que permite que aplicativos de terceiros acessem suas informações sem precisar compartilhar suas senhas.

**Vantagens do OAuth2:**
- **Segurança:** Não é necessário armazenar senhas em texto claro.
- **Controle:** É possível revogar o acesso a qualquer momento.
- **Escopo:** É possível limitar o acesso a apenas determinadas informações.

**Como Usar OAuth2 com Python:**

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

### 4.6.2. Aplicativos Menos Seguros

Se você estiver usando o Gmail, é possível que precise permitir o acesso de "aplicativos menos seguros" para que o envio de e-mails funcione. No entanto, essa opção reduz a segurança da sua conta e não é recomendada.

**Como Ativar:**
1. Acesse sua conta do Google.
2. Vá para "Segurança".
3. Na seção "Acesso a app menos seguro", ative a opção "Permitir aplicativos menos seguros".

**Atenção:**
- Essa opção pode não estar disponível para contas do Google Workspace (antigo G Suite).
- O Google pode bloquear o acesso de aplicativos menos seguros a qualquer momento, afetando a funcionalidade do seu aplicativo.

### 4.6.3. Conclusão

Neste capítulo, você aprendeu sobre considerações de segurança ao automatizar o envio de e-mails. A utilização de OAuth2 é a abordagem recomendada para garantir a segurança das suas credenciais e informações.

***

# 5. Web scraping e automação de sites

Nesta seção, você aprenderá como extrair dados de sites e automatizar interações com páginas web utilizando Python. O web scraping permite coletar informações de páginas da internet de forma automatizada, facilitando tarefas como atualização de preços, monitoramento de concorrentes, coleta de dados públicos, entre outros. Além disso, a automação de sites possibilita simular ações humanas, como preencher formulários, clicar em botões e navegar por páginas automaticamente, tornando processos repetitivos mais rápidos e eficientes.

Você verá exemplos práticos usando as bibliotecas `requests` e `BeautifulSoup` para extração de dados, além de dicas para lidar com tabelas HTML e boas práticas para respeitar as políticas dos sites.

*** 

## 5.1. Introdução ao Web Scraping

Web scraping é o processo de extrair dados de sites. Com Python, você pode usar bibliotecas como `BeautifulSoup` e `requests` para realizar essa tarefa. Abaixo está um exemplo básico de como fazer web scraping:

**Código Python:**

```python
import requests
from bs4 import BeautifulSoup
# URL do site que você deseja extrair dados
url = 'https://example.com'
# Fazer uma requisição HTTP para obter o conteúdo da página
response = requests.get(url)
# Verificar se a requisição foi bem-sucedida
if response.status_code == 200:
    # Analisar o conteúdo da página
    soup = BeautifulSoup(response.text, 'html.parser')
    # Extrair informações específicas
    dados = soup.find_all('div', class_='informacao')
    for dado in dados:
        print(dado.text)
else:
    print(f"Erro ao acessar o site: {response.status_code}")
```

**Resultado esperado**:
```plaintext
Informação 1
Informação 2
Informação 3
```

### 5.1.1. Boas Práticas de Web Scraping

Ao realizar web scraping, é importante seguir algumas boas práticas para garantir que você não cause problemas ao site que está acessando e para respeitar as regras de uso. Aqui estão algumas dicas:

- **Respeite o `robots.txt`:** Verifique se o site permite scraping e quais páginas podem ser acessadas.
- **Não sobrecarregue o servidor:** Faça requisições com intervalos de tempo razoáveis para evitar sobrecarga no site.
- **Use User-Agent:** Inclua um cabeçalho `User-Agent` na sua requisição para identificar seu script como um navegador legítimo.
- **Armazene os dados de forma estruturada:** Utilize formatos como CSV, JSON ou bancos de dados para armazenar os dados extraídos.
- **Mantenha a ética:** Não use scraping para coletar informações sensíveis ou violar os termos de serviço do site.


## 5.2. Extraindo Dados de Tabelas HTML

Para extrair dados de tabelas HTML, você pode usar o `pandas` junto com `BeautifulSoup`. Veja como fazer isso:

**Código Python:**

```python
import requests
from bs4 import BeautifulSoup
import pandas as pd
# URL do site com a tabela
url = 'https://example.com/tabela'
# Fazer uma requisição HTTP para obter o conteúdo da página
response = requests.get(url)
# Verificar se a requisição foi bem-sucedida
if response.status_code == 200:
    # Analisar o conteúdo da página
    soup = BeautifulSoup(response.text, 'html.parser')
    # Encontrar a tabela
    tabela = soup.find('table')
    # Ler a tabela usando pandas
    df = pd.read_html(str(tabela))[0]
    print(df)  # Exibir os dados da tabela
else:
    print(f"Erro ao acessar o site: {response.status_code}")
```

**Resultado esperado:**
```plaintext
   Coluna1  Coluna2  Coluna3
0  Dado1    Dado2    Dado3
1  Dado4    Dado5    Dado6
2  Dado7    Dado8    Dado9
```

### 5.2.1. Manipulando Dados Extraídos

Após extrair os dados de uma tabela HTML, você pode manipulá-los usando as funcionalidades do `pandas`. Por exemplo, você pode filtrar, agrupar ou transformar os dados conforme necessário.

**Exemplo de manipulação de dados:**
```python
# Filtrar dados onde Coluna1 é igual a 'Dado1'
dados_filtrados = df[df['Coluna1'] == 'Dado1']
print(dados_filtrados)
```

**Resultado esperado:**
```plaintext
   Coluna1  Coluna2  Coluna3
0  Dado1    Dado2    Dado3
```

## 5.3. Automação de Sites com Selenium

O Selenium é uma das ferramentas mais poderosas para automação de navegadores web. Com ele, é possível simular praticamente qualquer ação que um usuário humano faria em um site: clicar em botões, preencher formulários, rolar páginas, fazer login, baixar arquivos e até mesmo interagir com elementos dinâmicos carregados por JavaScript.

Essa abordagem é especialmente útil para automatizar tarefas em sistemas web internos, portais de clientes, ou para coletar dados de sites que exigem autenticação ou navegação dinâmica. O Selenium suporta diversos navegadores, como Chrome, Firefox e Edge, e pode ser integrado a scripts Python para criar robôs de automação robustos.

### Principais recursos do Selenium:

- **Automação de login e navegação:** Preencha campos de usuário e senha, clique em botões de login e navegue por diferentes páginas automaticamente.
- **Interação com elementos dinâmicos:** Clique em menus, marque checkboxes, selecione opções em listas suspensas e interaja com pop-ups.
- **Extração de dados após autenticação:** Acesse áreas restritas de sites e colete informações que não estão disponíveis publicamente.
- **Execução de tarefas repetitivas:** Automatize processos como cadastro de informações, download de relatórios, envio de formulários e muito mais.

### Quando usar Selenium?

- Quando o site depende fortemente de JavaScript para carregar dados.
- Quando é necessário simular ações humanas completas (login, navegação, cliques).
- Quando outras bibliotecas de scraping (como requests + BeautifulSoup) não conseguem acessar o conteúdo desejado.

### 5.3.1. Instalando o Selenium

Para usar o Selenium, instale a biblioteca e baixe o driver do navegador correspondente (por exemplo, ChromeDriver para Google Chrome):

```bash
pip install selenium
```
Baixe o ChromeDriver em: https://sites.google.com/chromium.org/driver/

### 5.3.2. Exemplo de Automação com Selenium

O exemplo abaixo mostra como automatizar o login em um site e extrair informações após o acesso:

```python
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Configurar o driver do navegador (certifique-se de que o ChromeDriver está no PATH)
driver = webdriver.Chrome()  # ou webdriver.Firefox() para Firefox

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
for dado in dados:
    print(dado.text)

# Fechar o navegador
driver.quit()
```

**Resultado esperado:**
```plaintext
Informação 1
Informação 2
Informação 3
```

### 5.3.3. Dicas e Boas Práticas com Selenium

- **Espere elementos carregarem:** Use `time.sleep()` para aguardar carregamentos simples, ou prefira `WebDriverWait` para esperar elementos específicos.
- **Evite hardcoding de caminhos:** Utilize seletores robustos (por exemplo, `By.ID`, `By.NAME`, `By.CSS_SELECTOR`) para localizar elementos.
- **Automatize downloads:** É possível configurar o navegador para baixar arquivos automaticamente para uma pasta específica.
- **Cuidado com bloqueios:** Sites podem detectar automação. Use User-Agent customizado e intervalos realistas entre ações.
- **Feche o navegador ao final:** Sempre chame `driver.quit()` para liberar recursos do sistema.

### 5.3.4. Exemplos de Interação com Elementos

- **Clicar em botões:**
```python
botao = driver.find_element(By.ID, 'botaoEnviar')
botao.click()
```
- **Preencher campos de texto:**
```python
campo_texto = driver.find_element(By.NAME, 'campoTexto')
campo_texto.send_keys('Texto a ser preenchido')
```
- **Selecionar opções em dropdowns:**
```python
from selenium.webdriver.support.ui import Select
select_element = driver.find_element(By.ID, 'dropdown')
select = Select(select_element)
select.select_by_visible_text('Opção 1')
```
- **Navegar para outra página:**
```python
driver.get('https://example.com/outra_pagina')
```

Com o Selenium, você pode criar robôs que automatizam tarefas repetitivas em sistemas web, economizando tempo e reduzindo erros manuais em processos do escritório.

***

# 6. Criação de interfaces gráficas simples (Tkinter ou PyWebIO)

Nesta seção, você aprenderá a criar interfaces gráficas simples para suas aplicações de automação. Interfaces gráficas são importantes para tornar suas ferramentas mais amigáveis e acessíveis, permitindo que usuários interajam com suas automações de forma mais intuitiva.

Você verá exemplos de como criar janelas, botões, campos de texto e outros elementos gráficos usando as bibliotecas Tkinter e PyWebIO. Além disso, aprenderá a coletar entradas de usuários e exibir resultados de forma visual.

***

## 6.1. Introdução ao Tkinter

Tkinter é a biblioteca padrão do Python para criação de interfaces gráficas. Com ela, é possível criar janelas, diálogos, botões, menus e outros elementos gráficos de forma simples e rápida.

### Exemplo básico de uma janela Tkinter:

```python
import tkinter as tk

# Função chamada ao clicar no botão
def ao_clicar():
    print("Botão clicado!")

# Criar a janela principal
janela = tk.Tk()
janela.title("Minha Primeira Janela")

# Criar um botão e adicionar à janela
botao = tk.Button(janela, text="Clique Aqui", command=ao_clicar)
botao.pack()

# Iniciar o loop da interface
janela.mainloop()
```

**Resultado esperado:**

Uma janela com um botão. Ao clicar no botão, a mensagem "Botão clicado!" deve aparecer no console.

## 6.2. Coletando Entradas de Usuário

É possível coletar entradas de usuário através de campos de texto, caixas de seleção, botões de opção, entre outros.

### Exemplo de coleta de entrada:

```python
import tkinter as tk

# Função chamada ao clicar no botão
def mostrar_nome():
    nome = entrada_nome.get()
    label_resultado.config(text=f"Olá, {nome}!")

# Criar a janela principal
janela = tk.Tk()
janela.title("Coletando Entradas")

# Criar um rótulo (label)
label_instrucoes = tk.Label(janela, text="Digite seu nome:")
label_instrucoes.pack()

# Criar um campo de entrada de texto
entrada_nome = tk.Entry(janela)
entrada_nome.pack()

# Criar um botão
botao = tk.Button(janela, text="Enviar", command=mostrar_nome)
botao.pack()

# Criar um rótulo para exibir o resultado
label_resultado = tk.Label(janela, text="")
label_resultado.pack()

# Iniciar o loop da interface
janela.mainloop()
```

**Resultado esperado:**

Uma janela com um campo de texto, um botão e um rótulo. O usuário deve digitar seu nome no campo de texto, clicar no botão e ver seu nome sendo exibido no rótulo.

## 6.3. Introdução ao PyWebIO

PyWebIO é uma biblioteca que permite criar interfaces web interativas diretamente em scripts Python, sem a necessidade de conhecimentos em HTML ou CSS. É uma ótima opção para quem deseja criar rapidamente interfaces para suas aplicações de automação.

### Exemplo básico de uso do PyWebIO:

```python
from pywebio import start_server
from pywebio.output import put_text
from pywebio.input import input

# Função principal da aplicação
def app():
    nome = input("Qual é o seu nome?")
    put_text(f"Olá, {nome}! Bem-vindo à automação com Python.")

# Iniciar o servidor web
start_server(app, port=8080)
```

**Resultado esperado:**

Ao executar o script, um servidor web será iniciado. Acesse `http://localhost:8080` no navegador para interagir com a aplicação.

## 6.4. Criando Formulários Simples

Tanto no Tkinter quanto no PyWebIO, é possível criar formulários para coletar informações de usuários.

### Exemplo de formulário no Tkinter:

```python
import tkinter as tk

# Função chamada ao enviar o formulário
def enviar_formulario():
    nome = entrada_nome.get()
    email = entrada_email.get()
    print(f"Nome: {nome}, Email: {email}")

# Criar a janela principal
janela = tk.Tk()
janela.title("Formulário de Contato")

# Criar campos de entrada
label_nome = tk.Label(janela, text="Nome:")
label_nome.pack()
entrada_nome = tk.Entry(janela)
entrada_nome.pack()

label_email = tk.Label(janela, text="Email:")
label_email.pack()
entrada_email = tk.Entry(janela)
entrada_email.pack()

# Criar botão de envio
botao_enviar = tk.Button(janela, text="Enviar", command=enviar_formulario)
botao_enviar.pack()

# Iniciar o loop da interface
janela.mainloop()
```

### Exemplo de formulário no PyWebIO:

```python
from pywebio import start_server
from pywebio.input import input_group
from pywebio.output import put_text

# Função principal da aplicação
def app():
    dados = input_group("Formulário de Contato", [
        input("Nome:"),
        input("Email:")
    ])
    put_text(f"Nome: {dados[0]}, Email: {dados[1]}")

# Iniciar o servidor web
start_server(app, port=8080)
```

# 7. Gerador automático de procurações e petições a partir de modelos

A criação de procurações e petições é uma tarefa comum e essencial no dia a dia de um escritório de advocacia. No entanto, preencher esses documentos manualmente pode ser demorado e sujeito a erros. Neste exemplo, vamos mostrar como automatizar a geração de procurações e petições a partir de modelos pré-definidos, utilizando Python e a biblioteca `docx` para manipulação de arquivos do Word.

## 7.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Leia um modelo de procuração ou petição, previamente elaborado em um arquivo do Word.
- Substitua automaticamente os campos variáveis (como nome do cliente, CPF, data, etc.) pelas informações fornecidas pelo usuário.
- Salve o documento preenchido com um novo nome, indicando que se trata de um documento personalizado.

## 7.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `ler_modelo()`: Para ler o arquivo modelo da procuração ou petição.
- `preencher_documento()`: Para substituir os campos variáveis pelas informações do cliente.
- `salvar_documento()`: Para salvar o documento preenchido com um novo nome.

## 7.3. Exemplo de Código

```python
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
modelo_path = 'modelo_procuracao.docx'
saida_path = 'procuracao_preenchida.docx'
dados_cliente = {
    'nome': 'João da Silva',
    'cpf': '123.456.789-00',
    'data': '01/06/2025'
}

# Ler o modelo
doc = ler_modelo(modelo_path)
# Preencher o documento
doc_preenchido = preencher_documento(doc, dados_cliente)
# Salvar o documento preenchido
salvar_documento(doc_preenchido, saida_path)

print(f"Procuração gerada com sucesso: {saida_path}")
```

## 7.4. Considerações Finais

A automação da geração de procurações e petições pode trazer ganhos significativos de produtividade para escritórios de advocacia, além de reduzir a possibilidade de erros na elaboração desses documentos. Com o uso de modelos e a automação do preenchimento, é possível agilizar o trabalho e garantir maior precisão e padronização nos documentos gerados.

***

# 8. Controle de prazos processuais (leitura de planilhas + envio de alertas por e-mail)

O controle de prazos processuais é uma atividade crítica para escritórios de advocacia, pois o não cumprimento de prazos pode resultar em prejuízos financeiros e danos à reputação do escritório. Neste exemplo, vamos mostrar como automatizar o controle de prazos processuais, utilizando Python para ler planilhas com os prazos e enviar alertas por e-mail quando os prazos estiverem se aproximando.

## 8.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Leia uma planilha com os prazos processuais, contendo informações como número do processo, descrição do prazo, data de início e data de término.
- Verifique quais prazos estão próximos do vencimento (por exemplo, faltando 3 dias ou menos).
- Envie um e-mail de alerta para o responsável pelo processo, informando sobre o prazo iminente.

## 8.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `ler_planilha()`: Para ler o arquivo da planilha com os prazos processuais.
- `verificar_prazos()`: Para identificar quais prazos estão próximos do vencimento.
- `enviar_alertas()`: Para enviar os e-mails de alerta.

## 8.3. Exemplo de Código

```python
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from datetime import datetime, timedelta

# Função para ler a planilha de prazos
def ler_planilha(caminho_arquivo):
    return pd.read_excel(caminho_arquivo)

# Função para verificar prazos próximos do vencimento
def verificar_prazos(df, dias_aviso=3):
    hoje = datetime.now()
    proximos_prazos = df[(df['Data de Término'] - hoje).dt.days <= dias_aviso]
    return proximos_prazos

# Função para enviar e-mail de alerta
def enviar_alertas(df_prazos, smtp_user, smtp_password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    for _, row in df_prazos.iterrows():
        destinatario = row['Email']
        assunto = f"Aviso de Prazo Processual - {row['Número do Processo']}"
        corpo = f"Prezado(a),\n\nEste é um aviso de que o prazo processual referente ao processo {row['Número do Processo']} está se aproximando.\nData de Término: {row['Data de Término'].strftime('%d/%m/%Y')}\n\nAtenciosamente,"
        
        msg = MIMEText(corpo)
        msg['Subject'] = assunto
        msg['From'] = smtp_user
        msg['To'] = destinatario
        
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(msg)

# Exemplo de uso
caminho_planilha = 'prazos_processuais.xlsx'
df = ler_planilha(caminho_planilha)
df_prazos_proximos = verificar_prazos(df)

# Enviar alertas (descomente a linha abaixo para enviar os e-mails)
# enviar_alertas(df_prazos_proximos, 'seu_email@gmail.com', 'sua_senha')

print("Processamento concluído.")
```

## 8.4. Considerações Finais

A automação do controle de prazos processuais pode trazer maior eficiência e segurança para escritórios de advocacia, reduzindo o risco de perda de prazos importantes e melhorando a comunicação com os clientes. Com o uso de planilhas e o envio automático de alertas por e-mail, é possível manter o controle dos prazos de forma organizada e prática.

***

# 9. Consulta a sites de tribunais

A consulta a sites de tribunais é uma atividade rotineira em escritórios de advocacia, utilizada para acompanhar o andamento de processos, verificar publicações e obter informações atualizadas sobre casos judiciais. Neste exemplo, vamos mostrar como automatizar a consulta a sites de tribunais, utilizando Python para acessar as informações de forma rápida e eficiente.

## 9.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Acesse o site de um tribunal (por exemplo, Tribunal de Justiça de São Paulo).
- Realize a busca por um processo utilizando o número do processo informado pelo usuário.
- Extraia e exiba informações relevantes sobre o andamento do processo.

## 9.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `acessar_site()`: Para acessar o site do tribunal.
- `buscar_processo()`: Para realizar a busca pelo processo.
- `extrair_informacoes()`: Para extrair as informações relevantes sobre o processo.

## 9.3. Exemplo de Código

```python
import requests
from bs4 import BeautifulSoup

# Função para acessar o site do tribunal
def acessar_site(url):
    response = requests.get(url)
    if response.status_code == 200:
        return response.text
    else:
        raise Exception(f"Erro ao acessar o site: {response.status_code}")

# Função para buscar o processo
def buscar_processo(html, numero_processo):
    soup = BeautifulSoup(html, 'html.parser')
    campo_busca = soup.find('input', {'name': 'numero_processo'})
    if campo_busca:
        # Simular a busca pelo processo
        return f"Processo {numero_processo} encontrado."
    else:
        return "Campo de busca não encontrado."

# Função para extrair informações do processo
def extrair_informacoes(html):
    soup = BeautifulSoup(html, 'html.parser')
    informacoes = {}
    # Extrair informações relevantes (exemplo: número do processo, partes, advogado, etc.)
    informacoes['numero_processo'] = soup.find('h1').text
    informacoes['partes'] = [parte.text for parte in soup.find_all('div', class_='parte')]
    return informacoes

# Exemplo de uso
url = 'https://www.tjsp.jus.br/'
numero_processo = '1234567890'
html_site = acessar_site(url)
resultado_busca = buscar_processo(html_site, numero_processo)

print(resultado_busca)
```

## 9.4. Considerações Finais

A automação da consulta a sites de tribunais pode trazer maior agilidade e eficiência para escritórios de advocacia, permitindo o acompanhamento em tempo real dos processos e a obtenção rápida de informações relevantes. Com o uso de web scraping e automação de navegação, é possível simplificar e otimizar essa tarefa rotineira.

***

# Escritório de Contabilidade

Os escritórios de contabilidade desempenham um papel crucial na gestão financeira de empresas e indivíduos, lidando com uma variedade de tarefas, como a preparação de declarações de impostos, a elaboração de demonstrações financeiras e o gerenciamento de folha de pagamento. A automação pode trazer benefícios significativos para a rotina desses escritórios, permitindo maior agilidade, precisão e conformidade.

Nesta seção, apresentamos exemplos práticos de automação voltados para escritórios de contabilidade, com o objetivo de demonstrar como as técnicas de programação e automação podem ser aplicadas para resolver problemas específicos dessa área.

Os exemplos foram escolhidos com base em situações comuns enfrentadas por contadores e profissionais da área contábil e buscam mostrar, de forma prática e didática, como a automação pode trazer ganhos significativos de produtividade, precisão e eficiência.

Ao final de cada exemplo, você encontrará dicas e sugestões para adaptar as soluções apresentadas à sua realidade, além de exercícios práticos para consolidar o aprendizado.

Esperamos que esta seção inspire você a identificar novas oportunidades de automação em sua rotina profissional e a explorar todo o potencial da programação e da inteligência artificial para transformar o seu trabalho.

***

# 10. Leitura e consolidação de extratos bancários (CSV)

A leitura e consolidação de extratos bancários é uma tarefa essencial para a gestão financeira de empresas, permitindo o acompanhamento de receitas, despesas e saldos. Neste exemplo, vamos mostrar como automatizar a leitura de extratos bancários em formato CSV e a consolidação das informações em um relatório financeiro.

## 10.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Leia um ou mais arquivos CSV contendo extratos bancários.
- Consolide as informações em um único relatório, com o resumo das receitas e despesas.

## 10.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `ler_extrato()`: Para ler o arquivo CSV do extrato bancário.
- `consolidar_extratos()`: Para consolidar as informações de múltiplos extratos.
- `gerar_relatorio()`: Para gerar o relatório financeiro.

## 10.3. Exemplo de Código

```python
import pandas as pd
import glob

# Função para ler o extrato bancário
def ler_extrato(caminho_arquivo):
    return pd.read_csv(caminho_arquivo, sep=';', encoding='latin1')

# Função para consolidar extratos
def consolidar_extratos(pasta_extratos):
    arquivos = glob.glob(f"{pasta_extratos}/*.csv")
    extratos = [ler_extrato(arquivo) for arquivo in arquivos]
    return pd.concat(extratos, ignore_index=True)

# Função para gerar o relatório financeiro
def gerar_relatorio(df_consolidado):
    relatorio = df_consolidado.groupby('Tipo')['Valor'].sum().reset_index()
    relatorio.to_csv('relatorio_financeiro.csv', index=False, sep=';')
    print("Relatório financeiro gerado: relatorio_financeiro.csv")

# Exemplo de uso
pasta_extratos = 'extratos_bancarios'
df_consolidado = consolidar_extratos(pasta_extratos)
gerar_relatorio(df_consolidado)
```

## 10.4. Considerações Finais

A automação da leitura e consolidação de extratos bancários pode trazer maior eficiência e precisão para a gestão financeira de empresas, facilitando o acompanhamento de receitas e despesas. Com o uso de arquivos CSV e a automação do processamento de dados, é possível simplificar e otimizar essa tarefa contábil.

***

# 11. Geração automática de guias de impostos

A geração de guias de impostos é uma obrigação periódica para empresas, sendo essencial para a regularidade fiscal e o cumprimento das obrigações tributárias. Neste exemplo, vamos mostrar como automatizar a geração de guias de impostos, utilizando Python para calcular os valores devidos e gerar os documentos necessários.

## 11.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Calcule os valores de impostos devidos com base nas receitas e despesas informadas.
- Gere as guias de pagamento dos impostos de forma automática.

## 11.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `calcular_impostos()`: Para calcular os impostos devidos.
- `gerar_guia_pagamento()`: Para gerar a guia de pagamento do imposto.

## 11.3. Exemplo de Código

```python
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
gerar_guia_pagamento(imposto_devido, 'guia_pagamento_imposto.txt')
```

## 11.4. Considerações Finais

A automação da geração de guias de impostos pode trazer maior agilidade e precisão para o cumprimento das obrigações tributárias das empresas, reduzindo o risco de erros e atrasos. Com o uso de cálculos automáticos e a geração de documentos, é possível simplificar e otimizar essa tarefa contábil.

***

# 12. Envio automático de boletos por e-mail

O envio de boletos por e-mail é uma prática comum e eficiente para a cobrança de clientes, facilitando o recebimento e melhorando o fluxo de caixa das empresas. Neste exemplo, vamos mostrar como automatizar o envio de boletos por e-mail, utilizando Python para gerar os boletos e enviá-los automaticamente para os clientes.

## 12.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Gere boletos bancários com base nas informações de cobrança.
- Envie os boletos gerados como anexos em e-mails para os clientes.

## 12.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `gerar_boleto()`: Para gerar o boleto bancário.
- `enviar_boleto_email()`: Para enviar o boleto por e-mail.

## 12.3. Exemplo de Código

```python
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Função para gerar o boleto bancário
def gerar_boleto(dados_boleto, caminho_saida):
    df_boleto = pd.DataFrame([dados_boleto])
    df_boleto.to_csv(caminho_saida, index=False, sep=';')
    print(f"Boleto gerado: {caminho_saida}")

# Função para enviar o boleto por e-mail
def enviar_boleto_email(destinatario, assunto, corpo, caminho_boleto, smtp_user, smtp_password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    # Criar o objeto de mensagem
    msg = MIMEMultipart()
    msg['Subject'] = assunto
    msg['From'] = smtp_user
    msg['To'] = destinatario
    # Adicionar o corpo do e-mail
    msg.attach(MIMEText(corpo, 'plain'))
    # Adicionar o anexo
    with open(caminho_boleto, 'rb') as anexo:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(anexo.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f'attachment; filename={caminho_boleto.split("/")[-1]}')
        msg.attach(parte)
    # Enviar o e-mail
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

# Exemplo de uso
dados_boleto = {
    'Cliente': 'João da Silva',
    'Valor': 150.00,
    'Vencimento': '2025-06-10'
}
caminho_boleto = 'boleto_joao_silva.csv'
gerar_boleto(dados_boleto, caminho_boleto)

# Enviar o boleto por e-mail (descomente a linha abaixo para enviar o e-mail)
# enviar_boleto_email('joao@email.com', 'Seu Boleto', 'Segue em anexo o boleto.', caminho_boleto, 'seu_email@gmail.com', 'sua_senha')
```

## 12.4. Considerações Finais

A automação do envio de boletos por e-mail pode trazer maior eficiência e agilidade para o processo de cobrança, melhorando o fluxo de caixa e a relação com os clientes. Com o uso de geração automática de boletos e envio programático por e-mail, é possível simplificar e otimizar essa tarefa financeira.

***

# Escritório de Logística

Os escritórios de logística são responsáveis pelo planejamento, execução e controle de operações de transporte e armazenamento de mercadorias, visando a eficiência e a redução de custos. A automação pode trazer benefícios significativos para a rotina desses escritórios, permitindo maior agilidade, precisão e rastreabilidade.

Nesta seção, apresentamos exemplos práticos de automação voltados para escritórios de logística, com o objetivo de demonstrar como as técnicas de programação e automação podem ser aplicadas para resolver problemas específicos dessa área.

Os exemplos foram escolhidos com base em situações comuns enfrentadas por profissionais de logística e buscam mostrar, de forma prática e didática, como a automação pode trazer ganhos significativos de produtividade, precisão e eficiência.

Ao final de cada exemplo, você encontrará dicas e sugestões para adaptar as soluções apresentadas à sua realidade, além de exercícios práticos para consolidar o aprendizado.

Esperamos que esta seção inspire você a identificar novas oportunidades de automação em sua rotina profissional e a explorar todo o potencial da programação e da inteligência artificial para transformar o seu trabalho.

***

# 13. Leitura e geração de manifestos (XML, PDF)

A leitura e geração de manifestos é uma atividade importante para o transporte de mercadorias, sendo essencial para a conformidade fiscal e o rastreamento das operações de transporte. Neste exemplo, vamos mostrar como automatizar a leitura de manifestos em formato XML e PDF, e a geração de novos manifestos com base nas informações lidas.

## 13.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Leia um manifesto em formato XML, extraindo as informações relevantes.
- Gere um novo manifesto em formato PDF, com as informações extraídas.

## 13.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `ler_manifesto_xml()`: Para ler o manifesto em formato XML.
- `gerar_manifesto_pdf()`: Para gerar o novo manifesto em formato PDF.

## 13.3. Exemplo de Código

```python
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
caminho_xml = 'manifesto.xml'
caminho_pdf = 'manifesto.pdf'
dados_manifesto = ler_manifesto_xml(caminho_xml)
gerar_manifesto_pdf(dados_manifesto, caminho_pdf)
```

## 13.4. Considerações Finais

A automação da leitura e geração de manifestos pode trazer maior eficiência e precisão para as operações de transporte, facilitando o cumprimento das obrigações fiscais e o rastreamento das mercadorias. Com o uso de arquivos XML e PDF, e a automação do processamento e geração de documentos, é possível simplificar e otimizar essa tarefa logística.

***

# 14. Roteirização com base em distância (API Google Maps ou OpenRoute)

A roteirização é uma atividade fundamental para a otimização do transporte de mercadorias, permitindo a definição das melhores rotas a serem seguidas pelos veículos. Neste exemplo, vamos mostrar como automatizar a roteirização com base em distância, utilizando APIs como Google Maps ou OpenRoute.

## 14.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Receba uma lista de endereços de origem e destino.
- Calcule a rota mais curta ou mais rápida entre os pontos, utilizando a API escolhida.
- Exiba o resultado com as informações da rota, como distância e tempo estimado.

## 14.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `calcular_rota_google_maps()`: Para calcular a rota utilizando a API do Google Maps.
- `calcular_rota_openroute()`: Para calcular a rota utilizando a API do OpenRoute.

## 14.3. Exemplo de Código

```python
import requests

# Função para calcular a rota utilizando a API do Google Maps
def calcular_rota_google_maps(origem, destino, chave_api):
    url = f"https://maps.googleapis.com/maps/api/directions/json?origin={origem}&destination={destino}&key={chave_api}"
    response = requests.get(url)
    dados = response.json()
    if dados['status'] == 'OK':
        rota = dados['routes'][0]
        distancia = rota['legs'][0]['distance']['text']
        tempo = rota['legs'][0]['duration']['text']
        return f"Distância: {distancia}, Tempo: {tempo}"
    else:
        return "Erro ao calcular a rota."

# Função para calcular a rota utilizando a API do OpenRoute
def calcular_rota_openroute(origem, destino):
    url = f"https://api.openrouteservice.org/v2/directions/driving-car?api_key=sua_chave_api&start={origem}&end={destino}"
    response = requests.get(url)
    dados = response.json()
    if dados['routes']:
        rota = dados['routes'][0]
        distancia = rota['summary']['length']
        tempo = rota['summary']['duration']
        return f"Distância: {distancia} metros, Tempo: {tempo} segundos"
    else:
        return "Erro ao calcular a rota."

# Exemplo de uso
origem = "-23.550520, -46.633308"  # São Paulo
destino = "-22.906847, -43.172896"  # Rio de Janeiro
# Calcular rota pelo Google Maps
resultado_google_maps = calcular_rota_google_maps(origem, destino, 'sua_chave_api_google_maps')
print("Google Maps:", resultado_google_maps)
# Calcular rota pelo OpenRoute
resultado_openroute = calcular_rota_openroute(origem, destino)
print("OpenRoute:", resultado_openroute)
```

## 14.4. Considerações Finais

A automação da roteirização com base em distância pode trazer ganhos significativos de eficiência e economia para as operações de transporte, permitindo a definição de rotas otimizadas e a redução de custos com combustível e tempo de viagem. Com o uso de APIs de mapeamento e a automação do processamento das informações, é possível simplificar e otimizar essa tarefa logística.

***

# 15. Acompanhamento de entregas via planilhas

O acompanhamento de entregas é uma atividade importante para garantir a satisfação dos clientes e a eficiência das operações de logística. Neste exemplo, vamos mostrar como automatizar o acompanhamento de entregas, utilizando Python para ler planilhas com os dados das entregas e enviar notificações automáticas.

## 15.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Leia uma planilha com os dados das entregas, contendo informações como código da entrega, cliente, endereço, data e status.
- Envie notificações automáticas para os clientes, informando sobre o status da entrega.

## 15.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `ler_planilha_entregas()`: Para ler a planilha com os dados das entregas.
- `enviar_notificacoes()`: Para enviar as notificações automáticas.

## 15.3. Exemplo de Código

```python
import pandas as pd
import smtplib
from email.mime.text import MIMEText

# Função para ler a planilha de entregas
def ler_planilha_entregas(caminho_arquivo):
    return pd.read_excel(caminho_arquivo)

# Função para enviar notificações automáticas
def enviar_notificacoes(df_entregas, smtp_user, smtp_password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    for _, row in df_entregas.iterrows():
        destinatario = row['Email']
        assunto = f"Atualização sobre sua entrega - {row['Código da Entrega']}"
        corpo = f"Prezado(a),\n\nInformamos que o status da sua entrega {row['Código da Entrega']} foi atualizado para: {row['Status']}.\n\nAtenciosamente,"
        
        msg = MIMEText(corpo)
        msg['Subject'] = assunto
        msg['From'] = smtp_user
        msg['To'] = destinatario
        
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(msg)

# Exemplo de uso
caminho_planilha = 'entregas.xlsx'
df_entregas = ler_planilha_entregas(caminho_planilha)
# Enviar notificações (descomente a linha abaixo para enviar os e-mails)
# enviar_notificacoes(df_entregas, 'seu_email@gmail.com', 'sua_senha')

print("Processamento concluído.")
```

## 15.4. Considerações Finais

A automação do acompanhamento de entregas pode trazer maior eficiência e agilidade para as operações de logística, melhorando a comunicação com os clientes e a gestão das entregas. Com o uso de planilhas e o envio automático de notificações por e-mail, é possível simplificar e otimizar essa tarefa logística.

***

# 16. Leitura de pedidos de marketplaces

A leitura e processamento de pedidos de marketplaces é uma tarefa importante para empresas que vendem produtos em plataformas como Mercado Livre, Amazon e outros. Neste exemplo, vamos mostrar como automatizar a leitura de pedidos recebidos em um marketplace, utilizando Python para processar as informações e gerar relatórios.

## 16.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Leia um arquivo com os pedidos recebidos, contendo informações como produto, quantidade, preço e dados do cliente.
- Gere um relatório com o resumo dos pedidos, incluindo o total de vendas e a lista de produtos mais vendidos.

## 16.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `ler_pedidos()`: Para ler o arquivo com os pedidos recebidos.
- `gerar_relatorio_pedidos()`: Para gerar o relatório com o resumo dos pedidos.

## 16.3. Exemplo de Código

```python
import pandas as pd

# Função para ler os pedidos recebidos
def ler_pedidos(caminho_arquivo):
    return pd.read_csv(caminho_arquivo, sep=';', encoding='latin1')

# Função para gerar o relatório com o resumo dos pedidos
def gerar_relatorio_pedidos(df_pedidos):
    relatorio = df_pedidos.groupby('Produto').agg({'Quantidade': 'sum', 'Preço': 'mean'}).reset_index()
    relatorio['Total Vendas'] = relatorio['Quantidade'] * relatorio['Preço']
    relatorio.to_csv('relatorio_pedidos.csv', index=False, sep=';')
    print("Relatório de pedidos gerado: relatorio_pedidos.csv")

# Exemplo de uso
caminho_arquivo = 'pedidos_marketplace.csv'
df_pedidos = ler_pedidos(caminho_arquivo)
gerar_relatorio_pedidos(df_pedidos)
```

## 16.4. Considerações Finais

A automação da leitura e processamento de pedidos de marketplaces pode trazer maior eficiência e agilidade para as operações de vendas online, facilitando o acompanhamento dos pedidos e a geração de relatórios. Com o uso de arquivos CSV e a automação do processamento de dados, é possível simplificar e otimizar essa tarefa comercial.

***

# 17. Atualização automática de estoque

A atualização de estoque é uma atividade crítica para empresas de e-commerce e vendas online, sendo essencial para garantir a disponibilidade dos produtos e evitar perdas financeiras. Neste exemplo, vamos mostrar como automatizar a atualização de estoque em planilhas Excel ou em um ERP simples, utilizando Python para ler os dados de vendas e atualizar as quantidades em estoque.

## 17.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Leia um arquivo com os dados de vendas, contendo informações como produto, quantidade vendida e data da venda.
- Atualize automaticamente as quantidades em estoque em uma planilha Excel ou em um ERP simples.

## 17.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `ler_vendas()`: Para ler o arquivo com os dados de vendas.
- `atualizar_estoque()`: Para atualizar as quantidades em estoque.

## 17.3. Exemplo de Código

```python
import pandas as pd

# Função para ler os dados de vendas
def ler_vendas(caminho_arquivo):
    return pd.read_csv(caminho_arquivo, sep=';', encoding='latin1')

# Função para atualizar o estoque
def atualizar_estoque(df_vendas, caminho_estoque):
    df_estoque = pd.read_excel(caminho_estoque)
    for _, row in df_vendas.iterrows():
        produto = row['Produto']
        quantidade_vendida = row['Quantidade']
        df_estoque.loc[df_estoque['Produto'] == produto, 'Estoque'] -= quantidade_vendida
    df_estoque.to_excel(caminho_estoque, index=False)
    print("Estoque atualizado com sucesso.")

# Exemplo de uso
caminho_vendas = 'vendas.csv'
caminho_estoque = 'estoque.xlsx'
df_vendas = ler_vendas(caminho_vendas)
atualizar_estoque(df_vendas, caminho_estoque)
```

## 17.4. Considerações Finais

A automação da atualização de estoque pode trazer maior eficiência e precisão para a gestão de estoques em empresas de e-commerce e vendas online, facilitando o acompanhamento das vendas e a atualização das quantidades em estoque. Com o uso de arquivos CSV e Excel, e a automação do processamento de dados, é possível simplificar e otimizar essa tarefa comercial.

***

# 18. Envio de notas fiscais e respostas automáticas a clientes

O envio de notas fiscais e respostas automáticas a clientes é uma prática importante para garantir a conformidade fiscal e melhorar o atendimento ao cliente. Neste exemplo, vamos mostrar como automatizar o envio de notas fiscais eletrônicas e respostas automáticas a clientes, utilizando Python para gerar os documentos e enviá-los por e-mail.

## 18.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Gere notas fiscais eletrônicas com base nas informações de vendas.
- Envie as notas fiscais geradas e respostas automáticas por e-mail para os clientes.

## 18.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `gerar_nota_fiscal()`: Para gerar a nota fiscal eletrônica.
- `enviar_nota_fiscal_email()`: Para enviar a nota fiscal e a resposta automática por e-mail.

## 18.3. Exemplo de Código

```python
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Função para gerar a nota fiscal eletrônica
def gerar_nota_fiscal(dados_venda, caminho_saida):
    df_nota = pd.DataFrame([dados_venda])
    df_nota.to_csv(caminho_saida, index=False, sep=';')
    print(f"Nota fiscal gerada: {caminho_saida}")

# Função para enviar a nota fiscal por e-mail
def enviar_nota_fiscal_email(destinatario, assunto, corpo, caminho_nota, smtp_user, smtp_password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    # Criar o objeto de mensagem
    msg = MIMEMultipart()
    msg['Subject'] = assunto
    msg['From'] = smtp_user
    msg['To'] = destinatario
    # Adicionar o corpo do e-mail
    msg.attach(MIMEText(corpo, 'plain'))
    # Adicionar o anexo
    with open(caminho_nota, 'rb') as anexo:
        parte = MIMEBase('application', 'octet-stream')
        parte.set_payload(anexo.read())
        encoders.encode_base64(parte)
        parte.add_header('Content-Disposition', f'attachment; filename={caminho_nota.split("/")[-1]}')
        msg.attach(parte)
    # Enviar o e-mail
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)

# Exemplo de uso
dados_venda = {
    'Cliente': 'João da Silva',
    'Produto': 'Notebook',
    'Quantidade': 1,
    'Valor Unitário': 5000.00,
    'CNPJ': '12.345.678/0001-95'
}
caminho_nota = 'nota_fiscal_joao_silva.csv'
gerar_nota_fiscal(dados_venda, caminho_nota)

# Enviar a nota fiscal por e-mail (descomente a linha abaixo para enviar o e-mail)
# enviar_nota_fiscal_email('joao@email.com', 'Sua Nota Fiscal', 'Segue em anexo a nota fiscal.', caminho_nota, 'seu_email@gmail.com', 'sua_senha')
```

## 18.4. Considerações Finais

A automação do envio de notas fiscais e respostas automáticas a clientes pode trazer maior eficiência e agilidade para o processo de faturamento e atendimento ao cliente, garantindo a conformidade fiscal e melhorando a experiência do cliente. Com o uso de geração automática de notas fiscais e envio programático por e-mail, é possível simplificar e otimizar essa tarefa financeira e comercial.

***

# Escritórios que prestam serviços para repartições públicas e órgãos governamentais

Os escritórios que prestam serviços para repartições públicas e órgãos governamentais enfrentam desafios específicos relacionados à gestão de documentos, cumprimento de prazos, atendimento a requisitos legais e normativos, entre outros. A automação pode trazer benefícios significativos para a rotina desses escritórios, permitindo maior agilidade, precisão e conformidade.

Nesta seção, apresentamos exemplos práticos de automação voltados para escritórios que prestam serviços para repartições públicas e órgãos governamentais, com o objetivo de demonstrar como as técnicas de programação e automação podem ser aplicadas para resolver problemas específicos dessa área.

Os exemplos foram escolhidos com base em situações comuns enfrentadas por profissionais que atuam com serviços públicos e buscam mostrar, de forma prática e didática, como a automação pode trazer ganhos significativos de produtividade, precisão e eficiência.

Ao final de cada exemplo, você encontrará dicas e sugestões para adaptar as soluções apresentadas à sua realidade, além de exercícios práticos para consolidar o aprendizado.

Esperamos que esta seção inspire você a identificar novas oportunidades de automação em sua rotina profissional e a explorar todo o potencial da programação e da inteligência artificial para transformar o seu trabalho.

***

# 19. Controle Automatizado de Protocolos em Repartição Pública

O controle de protocolos em repartições públicas é uma atividade importante para garantir a tramitação adequada de processos e documentos. Neste exemplo, vamos mostrar como automatizar o controle de protocolos em uma repartição pública, utilizando Python para ler planilhas com os dados dos protocolos e enviar notificações automáticas.

## 19.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Leia uma planilha com os dados dos protocolos, contendo informações como número do protocolo, descrição, data de entrada e situação.
- Verifique quais protocolos estão com a situação "Pendente".
- Envie um e-mail de notificação para o responsável pelo protocolo, informando sobre a pendência.

## 19.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `ler_planilha_protocolos()`: Para ler a planilha com os dados dos protocolos.
- `verificar_protocolos_pendentes()`: Para verificar quais protocolos estão pendentes.
- `enviar_notificacoes_protocolos()`: Para enviar as notificações por e-mail.

## 19.3. Exemplo de Código

```python
import pandas as pd
import smtplib
from email.mime.text import MIMEText

# Função para ler a planilha de protocolos
def ler_planilha_protocolos(caminho_arquivo):
    return pd.read_excel(caminho_arquivo)

# Função para verificar protocolos pendentes
def verificar_protocolos_pendentes(df_protocolos):
    return df_protocolos[df_protocolos['Situação'] == 'Pendente']

# Função para enviar notificações por e-mail
def enviar_notificacoes_protocolos(df_protocolos, smtp_user, smtp_password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    for _, row in df_protocolos.iterrows():
        destinatario = row['Email']
        assunto = f"Pendência no Protocolo - {row['Número do Protocolo']}"
        corpo = f"Prezado(a),\n\nIdentificamos uma pendência no protocolo {row['Número do Protocolo']}. Por favor, regularize a situação o quanto antes.\n\nAtenciosamente,"
        
        msg = MIMEText(corpo)
        msg['Subject'] = assunto
        msg['From'] = smtp_user
        msg['To'] = destinatario
        
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(msg)

# Exemplo de uso
caminho_planilha = 'protocolos_reparticao_publica.xlsx'
df_protocolos = ler_planilha_protocolos(caminho_planilha)
df_protocolos_pendentes = verificar_protocolos_pendentes(df_protocolos)

# Enviar notificações (descomente a linha abaixo para enviar os e-mails)
# enviar_notificacoes_protocolos(df_protocolos_pendentes, 'seu_email@gmail.com', 'sua_senha')

print("Processamento concluído.")
```

## 19.4. Considerações Finais

A automação do controle de protocolos em repartições públicas pode trazer maior eficiência e segurança para a tramitação de processos e documentos, garantindo o cumprimento de prazos e a comunicação adequada entre as partes envolvidas. Com o uso de planilhas e o envio automático de notificações por e-mail, é possível simplificar e otimizar essa tarefa administrativa.

***

# 20. Conversor de tabelas PDF → Excel

A conversão de tabelas em PDF para Excel é uma tarefa comum em escritórios que lidam com dados financeiros, orçamentos, relatórios e outros documentos que contêm tabelas. Neste exemplo, vamos mostrar como automatizar a conversão de tabelas em PDF para Excel, utilizando Python para extrair os dados das tabelas e salvá-los em um arquivo Excel.

## 20.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Leia um arquivo PDF contendo uma tabela.
- Extraia os dados da tabela e os salve em um arquivo Excel.

## 20.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `ler_tabela_pdf()`: Para ler a tabela do arquivo PDF.
- `salvar_tabela_excel()`: Para salvar a tabela em um arquivo Excel.

## 20.3. Exemplo de Código

```python
import pdfplumber
import pandas as pd

# Função para ler a tabela do PDF
def ler_tabela_pdf(caminho_arquivo):
    with pdfplumber.open(caminho_arquivo) as pdf:
        primeira_pagina = pdf.pages[0]
        tabela = primeira_pagina.extract_table()
    return tabela

# Função para salvar a tabela em um arquivo Excel
def salvar_tabela_excel(tabela, caminho_saida):
    df_tabela = pd.DataFrame(tabela[1:], columns=tabela[0])
    df_tabela.to_excel(caminho_saida, index=False)
    print(f"Tabela salva em Excel: {caminho_saida}")

# Exemplo de uso
caminho_pdf = 'tabela.pdf'
caminho_excel = 'tabela.xlsx'
tabela_extraida = ler_tabela_pdf(caminho_pdf)
salvar_tabela_excel(tabela_extraida, caminho_excel)
```

## 20.4. Considerações Finais

A automação da conversão de tabelas em PDF para Excel pode trazer maior eficiência e precisão para o trabalho com dados em escritórios, facilitando a análise e o compartilhamento das informações. Com o uso de arquivos PDF e Excel, e a automação do processamento e conversão de dados, é possível simplificar e otimizar essa tarefa de manipulação de dados.

***

# 21. Roteirização com base em distância (API Google Maps ou OpenRoute)

A roteirização é uma atividade fundamental para a otimização do transporte de mercadorias, permitindo a definição das melhores rotas a serem seguidas pelos veículos. Neste exemplo, vamos mostrar como automatizar a roteirização com base em distância, utilizando APIs como Google Maps ou OpenRoute.

## 21.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Receba uma lista de endereços de origem e destino.
- Calcule a rota mais curta ou mais rápida entre os pontos, utilizando a API escolhida.
- Exiba o resultado com as informações da rota, como distância e tempo estimado.

## 21.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `calcular_rota_google_maps()`: Para calcular a rota utilizando a API do Google Maps.
- `calcular_rota_openroute()`: Para calcular a rota utilizando a API do OpenRoute.

## 21.3. Exemplo de Código

```python
import requests

# Função para calcular a rota utilizando a API do Google Maps
def calcular_rota_google_maps(origem, destino, chave_api):
    url = f"https://maps.googleapis.com/maps/api/directions/json?origin={origem}&destination={destino}&key={chave_api}"
    response = requests.get(url)
    dados = response.json()
    if dados['status'] == 'OK':
        rota = dados['routes'][0]
        distancia = rota['legs'][0]['distance']['text']
        tempo = rota['legs'][0]['duration']['text']
        return f"Distância: {distancia}, Tempo: {tempo}"
    else:
        return "Erro ao calcular a rota."

# Função para calcular a rota utilizando a API do OpenRoute
def calcular_rota_openroute(origem, destino):
    url = f"https://api.openrouteservice.org/v2/directions/driving-car?api_key=sua_chave_api&start={origem}&end={destino}"
    response = requests.get(url)
    dados = response.json()
    if dados['routes']:
        rota = dados['routes'][0]
        distancia = rota['summary']['length']
        tempo = rota['summary']['duration']
        return f"Distância: {distancia} metros, Tempo: {tempo} segundos"
    else:
        return "Erro ao calcular a rota."

# Exemplo de uso
origem = "-23.550520, -46.633308"  # São Paulo
destino = "-22.906847, -43.172896"  # Rio de Janeiro
# Calcular rota pelo Google Maps
resultado_google_maps = calcular_rota_google_maps(origem, destino, 'sua_chave_api_google_maps')
print("Google Maps:", resultado_google_maps)
# Calcular rota pelo OpenRoute
resultado_openroute = calcular_rota_openroute(origem, destino)
print("OpenRoute:", resultado_openroute)
```

## 21.4. Considerações Finais

A automação da roteirização com base em distância pode trazer ganhos significativos de eficiência e economia para as operações de transporte, permitindo a definição de rotas otimizadas e a redução de custos com combustível e tempo de viagem. Com o uso de APIs de mapeamento e a automação do processamento das informações, é possível simplificar e otimizar essa tarefa logística.

***

# 22. Acompanhamento de entregas via planilhas

O acompanhamento de entregas é uma atividade importante para garantir a satisfação dos clientes e a eficiência das operações de logística. Neste exemplo, vamos mostrar como automatizar o acompanhamento de entregas, utilizando Python para ler planilhas com os dados das entregas e enviar notificações automáticas.

## 22.1. Objetivo

O objetivo deste exemplo é criar um script que:
- Leia uma planilha com os dados das entregas, contendo informações como código da entrega, cliente, endereço, data e status.
- Envie notificações automáticas para os clientes, informando sobre o status da entrega.

## 22.2. Estrutura do Código

O código será estruturado em funções, para facilitar a leitura e a manutenção. As principais funções serão:
- `ler_planilha_entregas()`: Para ler a planilha com os dados das entregas.
- `enviar_notificacoes()`: Para enviar as notificações automáticas.

## 22.3. Exemplo de Código

```python
import pandas as pd
import smtplib
from email.mime.text import MIMEText

# Função para ler a planilha de entregas
def ler_planilha_entregas(caminho_arquivo):
    return pd.read_excel(caminho_arquivo)

# Função para enviar notificações automáticas
def enviar_notificacoes(df_entregas, smtp_user, smtp_password):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    for _, row in df_entregas.iterrows():
        destinatario = row['Email']
        assunto = f"Atualização sobre sua entrega - {row['Código da Entrega']}"
        corpo = f"Prezado(a),\n\nInformamos que o status da sua entrega {row['Código da Entrega']} foi atualizado para: {row['Status']}.\n\nAtenciosamente,"
        
        msg = MIMEText(corpo)
        msg['Subject'] = assunto
        msg['From'] = smtp_user
        msg['To'] = destinatario
        
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(msg)

# Exemplo de uso
caminho_planilha = 'entregas.xlsx'
df_entregas = ler_planilha_entregas(caminho_planilha)
# Enviar notificações (descomente a linha abaixo para enviar os e-mails)
# enviar_notificacoes(df_entregas, 'seu_email@gmail.com', 'sua_senha')

print("Processamento concluído.")
```

## 22.4. Considerações Finais

A automação do acompanhamento de entregas pode trazer maior eficiência e agilidade para as operações de logística, melhorando a comunicação com os clientes e a gestão das entregas. Com o uso de planilhas e o envio automático de notificações por e-mail, é possível simplificar e otimizar essa tarefa logística.

***

# 23. Conclusão

A automação de processos em escritórios que prestam serviços para repartições públicas e órgãos governamentais pode trazer ganhos significativos de eficiência, precisão e conformidade. Com o uso de ferramentas como Python, planilhas e APIs, é possível simplificar tarefas administrativas, melhorar a comunicação com os clientes e otimizar a gestão de documentos e processos.
A implementação de soluções automatizadas pode reduzir o tempo gasto em atividades repetitivas, minimizar erros humanos e garantir que os prazos sejam cumpridos de forma mais eficaz.
Além disso, a automação pode facilitar o acesso e a análise de dados, permitindo uma tomada de decisão mais informada e ágil. Com o avanço da tecnologia e a crescente demanda por eficiência nos serviços públicos, a automação se torna uma ferramenta indispensável para os profissionais que atuam nesse setor.

***

# 24. Agradecimentos

Agradecemos a todos os leitores e profissionais que contribuíram para a elaboração deste material. Esperamos que os exemplos apresentados tenham sido úteis e inspiradores, e que possam ser aplicados na prática para melhorar a eficiência e a produtividade em seus escritórios.

***

# 25. Referências Bibliográficas (Formato ABNT)

FAKER. Faker Documentation. Disponível em: <https://faker.readthedocs.io/>. Acesso em: 17 jun. 2025.

FPDF. FPDF Documentation. Disponível em: <http://www.fpdf.org/>. Acesso em: 17 jun. 2025.

GOOGLE DEVELOPERS. Google Maps Platform Documentation. Disponível em: <https://developers.google.com/maps/documentation>. Acesso em: 17 jun. 2025.

LUTZ, Mark. Learning Python. 5. ed. Sebastopol: O’Reilly Media, 2013.

MCKINNEY, Wes. Python for Data Analysis. 3. ed. Sebastopol: O'Reilly Media, 2022.

OPENPYXL. OpenPyXL Documentation. Disponível em: <https://openpyxl.readthedocs.io/>. Acesso em: 17 jun. 2025.

OPENROUTESERVICE. API Documentation. Disponível em: <https://openrouteservice.org/sign-up/>. Acesso em: 17 jun. 2025.

PANDAS DEVELOPMENT TEAM. Pandas Documentation. Disponível em: <https://pandas.pydata.org/docs/>. Acesso em: 17 jun. 2025.

PDFPLUMBER. pdfplumber Documentation. Disponível em: <https://github.com/jsvine/pdfplumber>. Acesso em: 17 jun. 2025.

PYTHON SOFTWARE FOUNDATION. Python Documentation. Disponível em: <https://docs.python.org/3/>. Acesso em: 17 jun. 2025.

PYTHON-DOCX. python-docx Documentation. Disponível em: <https://python-docx.readthedocs.io/>. Acesso em: 17 jun. 2025.

PYPDF2. PyPDF2 Documentation. Disponível em: <https://pypdf2.readthedocs.io/>. Acesso em: 17 jun. 2025.

PYWEBIO. PyWebIO Documentation. Disponível em: <https://pywebio.readthedocs.io/>. Acesso em: 17 jun. 2025.

REITZ, Kenneth; SCHLUSSER, Tanya. The Hitchhiker’s Guide to Python. Sebastopol: O’Reilly Media, 2021.

REQUESTS. Requests: HTTP for Humans. Disponível em: <https://docs.python-requests.org/en/latest/>. Acesso em: 17 jun. 2025.

RICHARDSON, Leonard. BeautifulSoup Documentation. Disponível em: <https://www.crummy.com/software/BeautifulSoup/bs4/doc/>. Acesso em: 17 jun. 2025.

SELENIUMHQ. Selenium with Python Documentation. Disponível em: <https://selenium-python.readthedocs.io/>. Acesso em: 17 jun. 2025.

SWEIGART, Al. Automate the Boring Stuff with Python. 2. ed. San Francisco: No Starch Press, 2019.

TKDOCS. Tkinter Documentation. Disponível em: <https://tkdocs.com/>. Acesso em: 17 jun. 2025.

VANDERPLAS, Jake. Python Data Science Handbook. Sebastopol: O'Reilly Media, 2016.

***

# 26. ADENDO: Códigos em Python

- Exemplo Prático 01 – Análise de vendas de produtos online (CSV)  
  `01_analise_vendas_csv.py`
- Exemplo Prático 02 – Geração de relatórios financeiros (Excel)  
  `02_relatorio_financeiro_excel.py`
- Exemplo Prático 03 – Controle de inadimplência em aluguéis (Excel + E-mail)  
  `03_inadimplencia_alugueis_email.py`
- Exemplo Prático 04 – Envio de e-mails automatizados (Python + SMTP)  
  `04_envio_email_automatizado.py`
- Exemplo Prático 05 – Extração de texto de um PDF  
  `05_extracao_texto_pdf.py`
- Exemplo Prático 06 – Leitura e modificação de um arquivo Word  
  `06_modificacao_word.py`
- Exemplo Prático 07 – Automação de Excel com Pandas  
  `07_automacao_excel_pandas.py`
- Exemplo Prático 08 – Geração e envio de múltiplos relatórios diários por e-mail  
  `08_relatorios_diarios_email.py`
- Exemplo Prático 09 – Web scraping básico com BeautifulSoup  
  `09_webscraping_beautifulsoup.py`
- Exemplo Prático 10 – Extração de tabelas HTML para Excel  
  `10_extracao_tabelas_html_excel.py`
- Exemplo Prático 11 – Automação de sites com Selenium  
  `11_automacao_sites_selenium.py`
- Exemplo Prático 12 – Interface gráfica simples com Tkinter  
  `12_interface_tkinter.py`
- Exemplo Prático 13 – Interface web simples com PyWebIO  
  `13_interface_pywebio.py`
- Exemplo Prático 14 – Gerador automático de procurações e petições  
  `14_gerador_procuracoes_peticiones.py`
- Exemplo Prático 15 – Controle de prazos processuais com alertas  
  `15_controle_prazos_alertas.py`
- Exemplo Prático 16 – Consulta automatizada a sites de tribunais  
  `16_consulta_tribunais.py`
- Exemplo Prático 17 – Leitura e consolidação de extratos bancários  
  `17_extratos_bancarios.py`
- Exemplo Prático 18 – Geração automática de guias de impostos  
  `18_geracao_guias_impostos.py`
- Exemplo Prático 19 – Envio automático de boletos por e-mail  
  `19_envio_boletos_email.py`
- Exemplo Prático 20 – Leitura e geração de manifestos (XML, PDF)  
  `20_manifestos_xml_pdf.py`
- Exemplo Prático 21 – Roteirização com base em distância (API)  
  `21_roteirizacao_api.py`
- Exemplo Prático 22 – Acompanhamento de entregas via planilhas  
  `22_acompanhamento_entregas_planilha.py`
- Exemplo Prático 23 – Leitura de pedidos de marketplaces  
  `23_leitura_pedidos_marketplace.py`
- Exemplo Prático 24 – Atualização automática de estoque  
  `24_atualizacao_estoque.py`
- Exemplo Prático 25 – Envio de notas fiscais e respostas automáticas  
  `25_envio_notas_respostas.py`
- Exemplo Prático 26 – Controle automatizado de protocolos  
  `26_controle_protocolos.py`
- Exemplo Prático 27 – Conversor de tabelas PDF para Excel  
  `27_pdf_para_excel.py`
- Exemplo Prático 28 – Organizador de arquivos em pastas por cliente  
  `28_organizador_arquivos_clientes.py`
- Exemplo Prático 29 – Dashboard de pagamentos  
  `29_dashboard_pagamentos.py`
- Exemplo Prático 30 – Calculadora de impostos com interface  
  `30_calculadora_impostos_interface.py`