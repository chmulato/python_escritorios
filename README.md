![Capa do livro Python para Escritórios Inteligentes](images/capa.png)

***

# Automação de Tarefas para Escritórios com Python e Inteligência Artificial

Exemplos práticos para advocacia, contabilidade, logística, repartições públicas e e-commerce.

## Sobre

Este repositório reúne exemplos e projetos práticos de automação de tarefas em escritórios utilizando Python e inteligência artificial. O objetivo é mostrar, de forma acessível e prática, como profissionais de diferentes áreas podem automatizar rotinas repetitivas, economizar tempo e aumentar a produtividade — mesmo sem experiência prévia em programação.

Os exemplos abrangem desde a configuração do ambiente, manipulação de arquivos (Excel, CSV, PDF, Word), automação de e-mails, web scraping, até casos reais aplicados em advocacia, contabilidade, logística, repartições públicas e e-commerce. Cada projeto foi pensado para ser facilmente adaptado à realidade do escritório, servindo como ponto de partida para soluções personalizadas.

Se você busca tornar seu trabalho mais eficiente, reduzir erros manuais e explorar o potencial da automação com Python, este repositório é para você.

## Estrutura do Conteúdo

- **Parte I – Fundamentos da Automação:**  
  Instalação do ambiente (VS Code, Python, bibliotecas), manipulação de arquivos (Excel, CSV, PDF, Word), automação de e-mails, web scraping e criação de interfaces gráficas simples.

- **Parte II – Casos Reais por Tipo de Escritório:**  
  Exemplos práticos para advocacia, contabilidade, logística, e-commerce e repartições públicas, incluindo projetos como controle de inadimplência, geração de relatórios, automação de protocolos e envio automatizado de e-mails.

## Exemplos de Projetos Disponíveis

Os exemplos práticos estão organizados na pasta `codes/`:

- Exemplo Prático 01 – Análise de vendas de produtos online (CSV):  
  `codes/01_exemplo_pratico_vendas_online.py`
- Exemplo Prático 02 – Geração de relatórios financeiros (Excel):  
  `codes/02_exemplo_pratico_relatorio_financeiro.py`
- Exemplo Prático 03 – Controle de inadimplência em aluguéis (Excel + E-mail):  
  `codes/03_exemplo_pratico_alugueis_com_email.py`
- Exemplo Prático 04 – Envio de e-mails automatizados (Python + SMTP):  
  `codes/04_exemplo_pratico_envio_email.py`
- Exemplo Prático 05 – Extração de texto de um PDF:  
  `codes/05_exemplo_pratico_extracao_pdf.py`
- Exemplo Prático 06 – Leitura e modificação de um arquivo Word:  
  `codes/06_exemplo_pratico_modificacao_word.py`
- Exemplo Prático 07 – Automação de Excel com Pandas:  
  `codes/07_exemplo_pratico_automacao_excel.py`
- Exemplo Prático 08 – Geração e envio de múltiplos relatórios diários por e-mail:  
  `codes/08_exemplo_pratico_envio_relatorio_diario.py`
- Exemplo Prático 09 – Web scraping básico com BeautifulSoup:  
  `codes/09_exemplo_pratico_webscraping.py`
- Exemplo Prático 10 – Extração de tabelas HTML para Excel:  
  `codes/10_exemplo_pratico_tabela_html_excel.py`
- Exemplo Prático 11 – Automação de sites com Selenium:  
  `codes/11_exemplo_pratico_selenium.py`
- Exemplo Prático 12 – Interface gráfica simples com Tkinter:  
  `codes/12_exemplo_pratico_tkinter.py`
- Exemplo Prático 13 – Interface web simples com PyWebIO:  
  `codes/13_exemplo_pratico_pywebio.py`
- Exemplo Prático 14 – Gerador automático de procurações e petições:  
  `codes/14_exemplo_pratico_gerador_procuracao.py`
- Exemplo Prático 15 – Controle de prazos processuais com alertas:  
  `codes/15_exemplo_pratico_prazos_alerta.py`
- Exemplo Prático 16 – Consulta automatizada a sites de tribunais:  
  `codes/16_exemplo_pratico_consulta_tribunal.py`
- Exemplo Prático 17 – Leitura e consolidação de extratos bancários:  
  `codes/17_exemplo_pratico_extratos_bancarios.py`
- Exemplo Prático 18 – Geração automática de guias de impostos:  
  `codes/18_exemplo_pratico_guia_imposto.py`
- Exemplo Prático 19 – Envio automático de boletos por e-mail:  
  `codes/19_exemplo_pratico_envio_boleto.py`
- Exemplo Prático 20 – Leitura e geração de manifestos (XML, PDF):  
  `codes/20_exemplo_pratico_manifesto.py`
- Exemplo Prático 21 – Roteirização com base em distância (API):  
  `codes/21_exemplo_pratico_roteirizacao.py`
- Exemplo Prático 22 – Acompanhamento de entregas via planilhas:  
  `codes/22_exemplo_pratico_entregas.py`
- Exemplo Prático 23 – Leitura de pedidos de marketplaces:  
  `codes/23_exemplo_pratico_pedidos_marketplace.py`
- Exemplo Prático 24 – Atualização automática de estoque:  
  `codes/24_exemplo_pratico_estoque.py`
- Exemplo Prático 25 – Envio de notas fiscais e respostas automáticas:  
  `codes/25_exemplo_pratico_nota_fiscal.py`
- Exemplo Prático 26 – Controle automatizado de protocolos:  
  `codes/26_exemplo_pratico_protocolos.py`
- Exemplo Prático 27 – Conversor de tabelas PDF para Excel:  
  `codes/27_exemplo_pratico_pdf_para_excel.py`
- Exemplo Prático 28 – Organizador de arquivos em pastas por cliente:  
  `codes/28_exemplo_pratico_organizador_arquivos.py`
- Exemplo Prático 29 – Dashboard de pagamentos:  
  `codes/29_exemplo_pratico_dashboard_pagamentos.py`
- Exemplo Prático 30 – Calculadora de impostos com interface:  
  `codes/30_exemplo_pratico_calculadora_impostos.py`

## Como usar

1. Clone este repositório:
   ```
   git clone https://github.com/chmulato/python_escritorios.git
   ```
2. Instale as dependências necessárias:
   ```
   pip install pandas openpyxl requests beautifulsoup4 smtplib email python-docx PyPDF2 pdfplumber fpdf selenium pywebio faker
   ```
3. Navegue até a pasta `codes/` e execute o exemplo desejado conforme as instruções em cada arquivo.

## Observações Importantes

- Nunca compartilhe senhas em código público.
- Teste o envio de e-mails com poucos destinatários antes de usar em produção.
- Certifique-se de que o envio de e-mails está de acordo com as políticas da sua empresa e com a legislação vigente (ex: LGPD).
- Para automações envolvendo e-mail, adicione uma coluna "E-mail" nas planilhas de exemplo.
- Alguns exemplos podem exigir configuração adicional, como credenciais de API ou drivers de navegador (Selenium).

## Licença

Este projeto está licenciado sob a licença MIT.

A licença MIT é uma licença de software livre e permissiva. Ela permite que qualquer pessoa use, copie, modifique, mescle, publique, distribua, sublicencie e/ou venda cópias do software, desde que o aviso de copyright e a permissão sejam incluídos em todas as cópias ou partes substanciais do software. O software é fornecido "no estado em que se encontra", sem garantias de qualquer tipo.

## Curiosidades

### O que é Pythookup?

**Pythookup** é uma palavra criada especialmente para este livro, unindo “Python” — a linguagem de programação que democratizou a automação — com “hookup”, termo em inglês que significa conexão, ligação ou ponto de encontro.

Neste contexto, **Pythookup** representa a ideia de conectar pessoas, ideias e soluções por meio do Python. É o ponto de encontro entre profissionais de escritório e a tecnologia, onde a criatividade se une à automação para transformar rotinas, otimizar processos e abrir portas para novas possibilidades.

Ao longo deste livro, você verá que **Pythookup** é mais do que um nome: é um convite para se conectar ao universo da automação, aprender de forma prática e descobrir como o Python pode ser o elo entre o seu conhecimento e o futuro do trabalho.

***

![Contra-Capa do livro Python para Escritórios Inteligentes](images/contra_capa.png)