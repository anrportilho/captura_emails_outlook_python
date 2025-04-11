# 📧 Extrator de E-mails - Pós Sentença

Uma aplicação desenvolvida em Python com interface gráfica que automatiza a leitura de e-mails no Microsoft Outlook, extrai informações jurídicas relevantes e exporta os dados em planilha Excel. Projetada especialmente para automatizar atividades em operações de back-Office em equipes jurídicas que lidam com grande volume de emails. Pode ser adaptada a qualquer necessidade e setor, basta ajustar o assunto do email e os campos a serem monitorados na mensagem.

---

## 🚀 Funcionalidades

- Conexão automática com o Outlook via `win32com.client`
- Extração de campos personalizados formatados entre colchetes
- Interface gráfica intuitiva com `customtkinter`
- Exportação dos dados extraídos em `.xlsx`
- Registro de logs e erros em `log.txt`
- Filtragem por prefixo no assunto do e-mail (`PÓS SENTENÇA`)
- Ignora e-mails encaminhados ou respondidos (RE:, FW:, ENC:, etc.)

---

## 🔧 Requisitos

- Python 3.10+
- Microsoft Outlook instalado e configurado
- Sistema operacional Windows

### 📦 Dependências

Instale os requisitos com:

```bash
pip install pandas customtkinter pywin32
```

---

## ▶️ Como usar

### 1. Clone o repositório

```bash
git clone https://github.com/seuusuario/extrator-emails-juridico.git
cd extrator-emails-juridico
```

### 2. Executar a aplicação

Execute o script principal com:

```bash
python extrator_emails.py
```

A interface gráfica será carregada. Basta clicar em **"Iniciar Extração"** para processar os e-mails.

---

### 🗂️ Saída

- Um arquivo chamado `dados_extraidos_pos_sentenca.xlsx` será gerado na mesma pasta do script, contendo os dados extraídos.
- Todos os eventos e erros são registrados no arquivo `log.txt`.

---

## 🧪 Exemplo de Campos Extraídos

A aplicação extrai automaticamente os seguintes campos do corpo dos e-mails (entre colchetes):

- `[CODIGO_INTERNO]`
- `[NUMERO_PROCESSO]`
- `[NOME_JUIZADO]`
- `[NOME_PARTES]`
- `[CPF_CNPJ_AUTORES]`
- `[VALOR_CONDENACAO_SIMPLES]`
- `[VALOR_ATUALIZADO_CONDENACAO]`
- `[VALOR_MULTA_OU_DANOS]`
- `[DATA_FATO_GERADOR]`
- `[OBSERVACOES]`

---

## 🧑‍💻 Desenvolvedor

**Anderson Rocha**  
Desenvolvedor Python apaixonado por automações jurídicas e soluções eficientes.  
📅 **Abril/2025** – **Versão 1.1.0**

---

## 📃 Licença

Este projeto é de uso interno e profissional.  
Caso deseje utilizar ou adaptar para outro contexto, entre em contato com o desenvolvedor.

---

## 📌 Observações

- A aplicação ignora mensagens com prefixos como `RE:`, `FW:`, `ENC:`, etc.
- A data de recebimento do e-mail é processada como string para evitar problemas de fuso horário com o `win32timezone`.

---

## 💡 Sugestões ou melhorias?

Sinta-se à vontade para abrir uma issue ou pull request.  
Automação é a chave para produtividade! ⚙️
