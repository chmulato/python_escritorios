# Arquivos Necessários para Executar os Exemplos em Python do Livro

Este documento lista todos os arquivos e pastas de dados que você deve criar ou garantir que existam para executar **todos os exemplos práticos** do livro _Automação de Tarefas para Escritórios com Python e Inteligência Artificial_.

> **Dica:**  
> Os scripts que geram relatórios, boletos, notas fiscais, etc., normalmente criam os arquivos de saída automaticamente.  
> Para arquivos de entrada (como PDFs, modelos Word, planilhas), crie arquivos fictícios com os campos e formatos esperados, conforme descrito nos exemplos do livro.

---

## Estrutura Recomendada de Pastas

```
c:\dev\python_escritorios\
│
├── codes\
│   ├── arquivos_entrada\
│   ├── extratos_bancarios\
│   ├── relatorios_diarios\
│   └── ... (scripts .py e arquivos de dados)
│
├── docs\
│   └── arquivos_necessarios.md
├── python_escritorios.md
└── cronograma.md
```

---

## Parte 1 – Fundamentos

- **vendas.csv**  
  Arquivo de vendas para exemplos de análise de vendas.
- **relatorio_vendas.xlsx**  
  Arquivo Excel de relatório de vendas (gerado pelo script).
- **alugueis.xlsx**  
  Planilha de aluguéis para inadimplência.
- **oficio_exemplo.pdf**  
  PDF de exemplo para extração de texto.
- **contrato_exemplo.docx**  
  Documento Word para modificação de cláusulas.
- **dados.xlsx**, **dados.csv**  
  Arquivos genéricos para exemplos de leitura de Excel/CSV.
- **documento.pdf**, **documento.docx**  
  Arquivos genéricos para exemplos de leitura de PDF/Word.

---

## Parte 2 – Casos Reais

### Escritório de Advocacia

- **modelo_procuracao.docx**  
  Modelo Word com campos `{{nome}}`, `{{cpf}}`, `{{data}}`.
- **prazos_processuais.xlsx**  
  Planilha com prazos processuais e e-mails.
- **protocolos_reparticao_publica.xlsx**  
  Planilha de protocolos para repartição pública.

### Escritório de Contabilidade

- **extratos_bancarios\\***  
  Pasta com arquivos CSV de extratos bancários.
- **pagamentos.xlsx**  
  Planilha de pagamentos para dashboard.
- **guia_pagamento_imposto.txt**  
  Arquivo de saída gerado pelo script de impostos.

### Escritório de Logística

- **manifesto.xml**  
  Arquivo XML de manifesto de transporte.
- **entregas.xlsx**  
  Planilha de entregas para acompanhamento.
- **arquivos_entrada\\***  
  Pasta com arquivos para organizar por cliente.

### E-commerce/Vendas

- **pedidos_marketplace.csv**  
  Arquivo CSV de pedidos de marketplace.
- **estoque.xlsx**  
  Planilha de estoque para atualização.
- **vendas.csv**  
  Arquivo CSV de vendas para atualização de estoque.

### Outros

- **tabela.pdf**  
  PDF com tabela para conversão em Excel.
- **tabela.xlsx**  
  Arquivo Excel de saída da conversão.
- **relatorios_diarios\\***  
  Pasta para relatórios diários de vendas (gerada pelo script).
- **boleto_joao_silva.csv**  
  Exemplo de boleto gerado.
- **nota_fiscal_joao_silva.csv**  
  Exemplo de nota fiscal gerada.

---

## Observações Importantes

- **Geração automática:**  
  Muitos arquivos de saída (relatórios, boletos, notas fiscais) são criados ao rodar os scripts correspondentes.
- **Arquivos de entrada:**  
  Crie arquivos fictícios com os campos e formatos esperados, conforme descrito nos exemplos do livro.
- **Ajuste de caminhos:**  
  Se necessário, ajuste os caminhos dos arquivos nos scripts para refletir a sua estrutura de pastas.

---

**Com todos esses arquivos e pastas criados, você conseguirá executar todos os exemplos do livro sem dificuldades!**
