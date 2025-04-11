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


-----------------------------------------

2. Executar a aplicação
Execute o script principal com:

bash
Copiar
Editar
python extrator_emails.py
A interface gráfica será carregada. Basta clicar em "Iniciar Extração" para processar os e-mails.

🗂️ Saída
Um arquivo chamado dados_extraidos_pos_sentenca.xlsx será gerado na mesma pasta do script, contendo os dados extraídos.

Todos os eventos e erros são registrados no arquivo log.txt.

🧑‍💻 Desenvolvedor
Anderson Rocha
Desenvolvedor Python apaixonado por automações jurídicas e soluções eficientes.
📅 Abril/2025 – Versão 1.1.0

📃 Licença
Este projeto é de uso interno e profissional. Caso deseje utilizar ou adaptar para outro contexto, entre em contato com o desenvolvedor.

📌 Observações
A aplicação ignora mensagens com prefixos como RE:, FW:, ENC:, etc.

A data de recebimento do e-mail é processada em string para evitar problemas de fuso horário com o win32timezone.
