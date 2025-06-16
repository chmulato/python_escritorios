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
|   PYTHON PARA ESCRITÓRIOS INTELIGENTES                    |
|                                                           |
|   ┌───────────────────────────────────────────────┐       |
|   │ [  ][  ][  ][  ][  ][  ][  ][  ][  ][  ][  ]  │       |
|   │ [  ][  ][  ][  ][PY][  ][  ][  ][  ][  ][  ]  │       |
|   │ [  ][  ][  ][  ][  ][  ][  ][  ][  ][  ][  ]  │       |
|   └───────────────────────────────────────────────┘       |
|                                                           |
|   Christian Mulato - Programador de Computador            |
|___________________________________________________________|

# Apresentação

Vivemos uma era em que a tecnologia pode transformar o dia a dia dos escritórios, tornando tarefas repetitivas mais rápidas, precisas e menos cansativas. No entanto, muitos profissionais ainda gastam horas com atividades manuais que poderiam ser facilmente automatizadas.

Este livro nasceu da vontade de aproximar o poder da programação — especialmente do Python — do ambiente de escritórios, sejam eles de advocacia, contabilidade, logística, vendas ou serviços para repartições públicas e órgãos governamentais. O objetivo é mostrar, de forma prática e acessível, como automatizar processos comuns, economizar tempo e aumentar a produtividade, mesmo para quem nunca programou antes.

Ao longo dos capítulos, você encontrará exemplos reais, projetos práticos e dicas para aplicar imediatamente no seu trabalho. Não é necessário ser um especialista em tecnologia: basta curiosidade e vontade de aprender.

O conteúdo está dividido em duas partes principais para facilitar o aprendizado e a aplicação dos conceitos:

- **Parte I – Fundamentos da Automação:** Aqui você aprenderá os conceitos essenciais, desde a instalação do ambiente até a manipulação de arquivos, automação de e-mails, web scraping e criação de interfaces gráficas simples.
- **Parte II – Casos Reais por Tipo de Escritório:** Nesta parte, apresentamos exemplos práticos e projetos voltados para diferentes áreas, incluindo escritórios que prestam serviços para repartições públicas e órgãos governamentais, mostrando como adaptar as soluções para a sua realidade.

Espero que este material ajude você a transformar sua rotina profissional, abrindo portas para novas possibilidades e mostrando que a automação está ao alcance de todos.

Boa leitura e ótimos projetos!

**Christian Mulato**  
Programador de Computador

***

# Sumário

## Parte 1 – Fundamentos da Automação

1. Introdução à automação com Python  
2. Instalação de ambiente: VS Code, Python, bibliotecas úteis  
3. Trabalhando com arquivos (Excel, CSV, PDF, Word)  
4. Automação de e-mails e notificações  
5. Web scraping e automação de sites  
6. Criação de interfaces gráficas simples (Tkinter ou PyWebIO)  

## Parte 2 – Casos Reais por Tipo de Escritório

### Escritório de Advocacia
7. Gerador automático de procurações e petições a partir de modelos  
8. Controle de prazos processuais (leitura de planilhas + envio de alertas por e-mail)  
9. Consulta a sites de tribunais  

### Escritório de Contabilidade
10. Leitura e consolidação de extratos bancários (CSV)  
11. Geração automática de guias de impostos  
12. Envio automático de boletos por e-mail  

### Escritório de Logística
13. Leitura e geração de manifestos (XML, PDF)  
14. Roteirização com base em distância (API Google Maps ou OpenRoute)  
15. Acompanhamento de entregas via planilhas atualizadas  

### E-commerce e Vendas Online
16. Leitura de pedidos de marketplaces  
17. Atualização automática de estoque em Excel/ERP simples  
18. Envio de notas fiscais e respostas automáticas a clientes  

## Projetos Práticos

19. Organizador de arquivos em pastas por cliente  
20. Conversor de tabelas PDF → Excel  
21. Dashboard de pagamentos  
22. Web scraper de preços da concorrência  
23. Calculadora de impostos com interface simples  

## Anexos

- Figuras e diagramas  
- Referências

## Códigos em Python

- Exemplo Prático 01
- Exemplo Prático 02

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

# Parte 1 – Fundamentos da Automação

Nesta primeira parte, você vai aprender os conceitos e ferramentas essenciais para começar a automatizar tarefas com Python no ambiente de escritório. O objetivo é construir uma base sólida, mesmo para quem nunca programou antes, mostrando passo a passo como instalar o ambiente, manipular arquivos, enviar e-mails automaticamente, coletar informações da internet e criar interfaces gráficas simples.

Ao dominar esses fundamentos, você estará preparado para aplicar a automação em diferentes áreas e desafios do dia a dia profissional.

***

## 1. Introdução à automação com Python

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

## 2. Instalação de ambiente: VS Code, Python, bibliotecas úteis

Antes de começar a programar e automatizar tarefas, é importante preparar o ambiente de desenvolvimento. Nesta etapa, você vai instalar o Python, o Visual Studio Code (VS Code) e as principais bibliotecas que serão usadas ao longo do livro.

### 2.1 Instalando o Python

1. Acesse o site oficial: [python.org/downloads](https://www.python.org/downloads/)
2. Baixe a versão recomendada para o seu sistema operacional (Windows, macOS ou Linux).
3. Execute o instalador e marque a opção **"Add Python to PATH"** antes de clicar em "Install Now".
4. Para verificar se a instalação foi bem-sucedida, abra o terminal (Prompt de Comando no Windows) e digite:
   ```
   python --version
   ```
   O terminal deve exibir a versão instalada do Python.

### 2.2 Instalando o Visual Studio Code (VS Code)

1. Acesse: [code.visualstudio.com](https://code.visualstudio.com/)
2. Baixe e instale o VS Code para o seu sistema operacional.
3. Após a instalação, abra o VS Code e, na barra lateral esquerda, clique em **Extensões** (ícone de quadrado).
4. Busque e instale a extensão **Python** (desenvolvida pela Microsoft).

### 2.3 Instalando bibliotecas úteis

As bibliotecas Python ampliam as funcionalidades da linguagem e facilitam a automação de tarefas. Algumas das principais bibliotecas que usaremos são:

- **pandas** (manipulação de dados e planilhas)
- **openpyxl** (trabalho com arquivos Excel)
- **python-docx** (manipulação de arquivos Word)
- **PyPDF2** (leitura de PDFs)
- **requests** (requisições web)
- **beautifulsoup4** (web scraping)
- **tkinter** ou **PyWebIO** (interfaces gráficas)

Para instalar essas bibliotecas, abra o terminal e execute:

```
pip install pandas openpyxl python-docx PyPDF2 requests beautifulsoup4
```

> **Obs.:** O `tkinter` já vem instalado com o Python na maioria das distribuições. O `PyWebIO` pode ser instalado separadamente, se preferir usar interfaces web:
>
> ```
> pip install pywebio
> ```

### 2.4 Testando o ambiente

Abra o VS Code, crie um novo arquivo chamado `teste.py` e cole o seguinte código:

```python
print("Ambiente Python configurado com sucesso!")
```

Salve o arquivo e execute-o (clique com o botão direito e escolha "Run Python File in Terminal" ou use o terminal com `python teste.py`).  
Se aparecer a mensagem, seu ambiente está pronto para começar!

***

## 3. Trabalhando com arquivos (Excel, CSV, PDF, Word)

Nesta seção, você aprenderá a manipular diferentes tipos de arquivos comuns no ambiente de escritório, como Excel, CSV, PDF e Word. Essas habilidades são essenciais para automatizar tarefas que envolvem leitura, escrita e processamento de dados.

### 3.1. Arquivos CSV

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

### 3.2. Arquivos Excel

Para manipular arquivos Excel, use a biblioteca `pandas` ou `openpyxl`. Aqui está um exemplo com `pandas`:

```python
import pandas as pd
# Lendo um arquivo Excel
df = pd.read_excel('dados.xlsx', sheet_name='Planilha1')
print(df.head())  # Exibe as primeiras linhas do DataFrame
# Escrevendo em um arquivo Excel
df.to_excel('saida.xlsx', index=False, sheet_name='Resultados')
```

### 3.3. Arquivos PDF

Para ler arquivos PDF, você pode usar a biblioteca `PyPDF2`. Aqui está um exemplo simples:

```python
import PyPDF2
# Lendo um arquivo PDF
with open('documento.pdf', 'rb') as file:
    leitor = PyPDF2.PdfReader(file)
    for pagina in leitor.pages:
        print(pagina.extract_text())  # Extrai o texto de cada página
```

### 3.4. Arquivos Word

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

### 3.5. Configurando as dependências do ambiente Python

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

### 3.6. Exercícios Práticos

#### 3.6.1. Exercício prático: Análise de vendas de produtos online (CSV)

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

Código Python para resolver o exercício:

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

Resultado esperado (`relatorio_vendas.csv`):
```
produto,total_vendido
Mouse,150
Teclado,120
Monitor,900
```
**Conclusão**
Com este exercício, você praticou a leitura de arquivos CSV, manipulação de dados com pandas e geração de relatórios. Essas habilidades são fundamentais para automatizar análises de vendas e outras tarefas relacionadas a dados em escritórios.

#### 3.6.2. **Manipulação de planilhas Excel:**  
   Abrir uma planilha Excel, adicionar uma nova coluna calculada (ex: imposto sobre vendas), e salvar a planilha modificada.

#### 3.6.3. **Extração de texto de um PDF:**  
   Ler um arquivo PDF e extrair todo o texto, salvando o conteúdo em um arquivo `.txt`.

#### 3.6.4. **Leitura e modificação de um arquivo Word:**  
   Abrir um documento Word, listar todos os parágrafos, adicionar um novo parágrafo ao final e salvar o documento com outro nome.
