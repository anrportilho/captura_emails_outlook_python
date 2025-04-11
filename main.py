import win32com.client
import pandas as pd
import re
import unicodedata
import customtkinter as ctk
import threading
import pythoncom
import traceback
from datetime import datetime

# ===== Funções de utilidade =====

def log_evento(texto):
    with open("log.txt", "a", encoding="utf-8") as f:
        f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | {texto}\n")

def limpar_texto(texto):
    if not texto:
        return ""
    texto = texto.strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = re.sub(r"\s+", " ", texto)
    texto = texto.replace("–", "-")
    return texto.upper()

def extrair_campos_formatados(texto, campos_interesse):
    padrao = re.compile(r"\[(.+?)\]\s*(.+)")
    encontrados = dict(re.findall(padrao, texto))
    return {campo: encontrados.get(campo, "").strip() for campo in campos_interesse}

# ===== Extração de e-mails =====

def localizar_emails_por_assunto_prefixo(prefixo_assunto):
    log_evento("Iniciando conexão com Outlook")
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # Caixa de entrada
    mensagens = inbox.Items
    mensagens.Sort("[ReceivedTime]", True)

    dados_emails = []
    mensagens_lista = [msg for msg in mensagens]

    prefixo_normalizado = limpar_texto(prefixo_assunto)

    campos_desejados = [
        "CODIGO_INTERNO",
        "NUMERO_PROCESSO",
        "NOME_JUIZADO",
        "NOME_PARTES",
        "CPF_CNPJ_AUTORES",
        "VALOR_CONDENACAO_SIMPLES",
        "VALOR_ATUALIZADO_CONDENACAO",
        "VALOR_MULTA_OU_DANOS",
        "DATA_FATO_GERADOR",
        "OBSERVACOES"
    ]

    log_evento(f"Total de mensagens: {len(mensagens_lista)}")

    for mensagem in mensagens_lista:
        try:
            assunto_original = mensagem.Subject or ""
            assunto = limpar_texto(assunto_original)

            if re.match(r"^(RE|FW|ENC|RES):", assunto, flags=re.IGNORECASE):
                continue

            if assunto.startswith(prefixo_normalizado):
                corpo = mensagem.Body or ""
                campos_extraidos = extrair_campos_formatados(corpo, campos_desejados)

                try:
                    remetente = mensagem.Sender
                    if remetente and remetente.Class == 23:  # ExchangeUser
                        exchange_user = remetente.GetExchangeUser()
                        if exchange_user:
                            email = exchange_user.PrimarySmtpAddress
                        else:
                            email = remetente.Address
                    else:
                        email = remetente.Address if remetente else "DESCONHECIDO"
                except Exception as e:
                    log_evento(f"Erro ao acessar remetente: {str(e)}")
                    email = "ERRO AO OBTER EMAIL"

                # Aqui está a modificação principal para evitar a conversão de datetime
                # que requer win32timezone
                try:
                    # Obter a data como string diretamente
                    data_recebimento = str(mensagem.ReceivedTime)
                    # Remover a parte de timezone se existir
                    data_recebimento = data_recebimento.split('+')[0].strip()
                except Exception as e:
                    log_evento(f"Erro ao processar data: {str(e)}")
                    data_recebimento = "DATA INDISPONÍVEL"

                dados_emails.append({
                    "Remetente": email,
                    "Assunto": assunto_original.strip(),
                    "Data": data_recebimento,
                    **campos_extraidos
                })

        except Exception as e:
            log_evento("Erro ao processar mensagem:")
            log_evento(traceback.format_exc())
            continue

    log_evento(f"Total extraído: {len(dados_emails)}")
    return dados_emails

# ===== Interface Gráfica com customtkinter =====

def iniciar_extracao():
    def run():
        pythoncom.CoInitialize()
        try:
            btn_start.configure(state="disabled")
            label_status.configure(text="Extraindo e-mails, aguarde...")
            log_evento("Extração iniciada via GUI")

            dados = localizar_emails_por_assunto_prefixo("PÓS SENTENÇA")
            df = pd.DataFrame(dados)
            df.to_excel("dados_extraidos_pos_sentenca.xlsx", index=False)

            label_status.configure(text=f"{len(df)} e-mails exportados com sucesso.")
            log_evento("Extração finalizada com sucesso.")
        except Exception as e:
            erro = traceback.format_exc()
            log_evento("Erro durante execução da thread:")
            log_evento(erro)
            label_status.configure(text="Erro ao extrair e-mails. Verifique o log.")
        finally:
            btn_start.configure(state="normal")

    thread = threading.Thread(target=run)
    thread.start()

# ======= GUI =======

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Extrator de E-mails - Jurídico")
app.geometry("520x250")

label_titulo = ctk.CTkLabel(app, text="Automação Pós Sentença", font=("Arial", 18))
label_titulo.pack(pady=10)

btn_start = ctk.CTkButton(app, text="Iniciar Extração", command=iniciar_extracao)
btn_start.pack(pady=10)

label_status = ctk.CTkLabel(app, text="")
label_status.pack(pady=20)

rodape = ctk.CTkLabel(
    app,
    text="Desenvolvido com ❤ por Anderson Rocha • Abril/2025 • Versão 1.1.0",
    font=("Arial", 10),
    text_color="#888"
)
rodape.pack(side="bottom", pady=5)

app.mainloop()
