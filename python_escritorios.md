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

### Escritórios que prestam serviços para repartições públicas e órgãos governamentais

19. Controle Automatizado de Protocolos em Repartição Pública
20. Conversor de tabelas PDF → Excel
21. Organizador de arquivos em pastas por cliente
22. Dashboard de pagamentos

## Anexos

- Figuras e diagramas  
- Referências

## Códigos em Python

- Exemplo Prático 01 - Análise de vendas de produtos online (CSV)
   `codes/01_exemplo_pratico_vendas_online.py`
- Exemplo Prático 02 - Geração de relatórios financeiros (Excel)
   `codes/02_exemplo_pratico_relatorio_financeiro.py`
- Exemplo Prático 03 - Controle de inadimplência em aluguéis (Excel + E-mail)
   `codes/03_exemplo_pratico_alugueis_com_email.py`
- Exemplo Prático 04 - Envio de e-mails automatizados (Python + SMTP)
   `codes/04_exemplo_pratico_envio_email.py`

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

## 2. Instalação de Ambiente: VS Code, Python e Bibliotecas Úteis

Antes de iniciar os exemplos práticos, é importante preparar o ambiente de desenvolvimento. Siga os passos abaixo:

### 2.1. Instale o Visual Studio Code (VS Code)

- Acesse: [https://code.visualstudio.com/](https://code.visualstudio.com/)
- Baixe e instale a versão adequada para o seu sistema operacional (Windows, macOS ou Linux).

### 2.2. Instale o Python

- Acesse: [https://www.python.org/downloads/](https://www.python.org/downloads/)
- Baixe a versão mais recente do Python 3.x.
- Durante a instalação, marque a opção **"Add Python to PATH"**.

### 2.3. Instale Bibliotecas Úteis

Abra o terminal do VS Code (atalho: `Ctrl + ``) e execute os comandos abaixo para instalar as principais bibliotecas que serão usadas nos exemplos:

```bash
pip install pandas openpyxl
```

- `pandas`: Manipulação de dados e leitura de planilhas Excel.
- `openpyxl`: Leitura e escrita de arquivos `.xlsx` (Excel).

### 2.4. (Opcional) Instale Extensões no VS Code

- **Python** (Microsoft): Suporte a sintaxe, depuração e execução de scripts Python.
- **Jupyter**: Para notebooks interativos, se desejar.

---

> Após esses passos, seu ambiente estará pronto para executar os exemplos práticos.

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

Nesta seção, você encontrará exercícios práticos para aplicar os conceitos aprendidos sobre manipulação de arquivos. Esses exercícios são projetados para serem desafiadores e ajudarão a consolidar seu conhecimento em automação de tarefas comuns em escritórios.

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

## 3.6.2. Automatizando Relatórios de Inadimplência em Aluguéis com Python e Excel

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

### Adendo: Envio Real de E-mails para Inquilinos Inadimplentes

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
```

**Atenção:**  
- Nunca compartilhe senhas em código público.
- Teste o envio com poucos destinatários antes de usar em produção.
- Certifique-se de que o envio de e-mails está de acordo com as políticas da sua empresa e com a legislação vigente (LGPD).

Com essa automação, a empresa pode agilizar a comunicação com os inquilinos inadimplentes, tornando o processo mais eficiente e reduzindo o tempo gasto com cobranças manuais.

**Conclusão:**
Com este exercício, você praticou a leitura e escrita de arquivos Excel, filtragem de dados com pandas. Essas habilidades são essenciais para automatizar o controle de inadimplência em empresas de aluguel de imóveis, tornando o processo mais eficiente e organizado.

***

#### 3.6.3. **Extração de texto de um PDF:**  
   Ler um arquivo PDF e extrair todo o texto, salvando o conteúdo em um arquivo `.txt`.


***

#### 3.6.4. **Leitura e modificação de um arquivo Word:**  
   Abrir um documento Word, listar todos os parágrafos, adicionar um novo parágrafo ao final e salvar o documento com outro nome.


***

