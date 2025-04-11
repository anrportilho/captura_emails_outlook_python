# Captura e tabula emails no outlook com python
Este projeto realiza a varredura da caixa de entrada do Outlook, identifica mensagens com um assunto específico (ex: "PÓS SENTENÇA") e extrai automaticamente campos relevantes do conteúdo do e-mail. O resultado é salvo em um arquivo Excel para posterior análise e organização jurídica.

# 📧 Extrator de E-mails - Pós Sentença (Jurídico)

Uma ferramenta desenvolvida para automatizar a extração de e-mails com dados relacionados a processos jurídicos recebidos via Microsoft Outlook, com interface gráfica amigável utilizando `customtkinter`.

---

## 🔍 Visão Geral

Este projeto realiza a varredura da caixa de entrada do Outlook, identifica mensagens com um assunto específico (ex: "PÓS SENTENÇA") e extrai automaticamente campos relevantes do conteúdo do e-mail. O resultado é salvo em um arquivo Excel para posterior análise e organização jurídica.

---

## 🧠 Funcionalidades

- Conexão automática com o Outlook
- Extração de e-mails com prefixo de assunto customizável
- Limpeza e padronização de texto
- Extração de campos estruturados do corpo do e-mail
- Registro de logs de execução em `log.txt`
- Exportação para planilha Excel (`.xlsx`)
- Interface gráfica com `customtkinter`
- Multithreading para não travar a interface durante a execução

---

## 📋 Campos Extraídos

A extração busca pelos seguintes campos estruturados dentro do corpo do e-mail:

- `CODIGO_INTERNO`
- `NUMERO_PROCESSO`
- `NOME_JUIZADO`
- `NOME_PARTES`
- `CPF_CNPJ_AUTORES`
- `VALOR_CONDENACAO_SIMPLES`
- `VALOR_ATUALIZADO_CONDENACAO`
- `VALOR_MULTA_OU_DANOS`
- `DATA_FATO_GERADOR`
- `OBSERVACOES`

---

## 🛠️ Tecnologias Utilizadas

- [Python 3.10+](https://www.python.org/)
- [win32com.client (pywin32)](https://pypi.org/project/pywin32/)
- [pandas](https://pandas.pydata.org/)
- [customtkinter](https://github.com/TomSchimansky/CustomTkinter)

---

## ▶️ Como Usar

### 1. Pré-requisitos

- Microsoft Outlook instalado e configurado
- Python 3.10 ou superior instalado
- Pacotes Python necessários:

```bash
pip install pandas pywin32 customtkinter
