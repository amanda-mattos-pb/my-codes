# versão beta 1.0
import pandas as pd
import pywhatkit
import time
import random

# === CONFIGURAÇÕES ===
ARQUIVO_PLANILHA = r"C:\Users\Amanda\OneDrive\Documents\GitHub\my-codes\mensagem_whatsapp\alunos.xlsx" # nome do arquivo da planilha
TEMPO_MINIMO = 5    # segundos mínimos entre mensagens
TEMPO_MAXIMO = 30   # segundos máximos entre mensagens

# Número para receber a mensagem final (substitua pelo seu número com DDD e código +55)
NUMERO_RESUMO = "+5583999944947"  

# === LEITURA DA PLANILHA ===
try:
    df = pd.read_excel(ARQUIVO_PLANILHA)
except Exception as e:
    print(f"Erro ao abrir a planilha: {e}")
    exit()

# === PREPARAÇÃO DAS MENSAGENS ===

mensagens_ativos = []
mensagens_pendentes = []

for _, row in df.iterrows():
    nome = row.get("Nome", "Aluno")
    telefone = str(row.get("Telefone", "")).strip()
    status = str(row.get("Status", "")).strip().lower()
    pagamento = str(row.get("Pagamento", "")).strip().lower()

    if status != "ativo" or not telefone:
        continue

    numero_formatado = "+55" + telefone

    if pagamento == "atrasado" or pagamento == "pendente":
        mensagem = (
            f"Olá {nome}, tudo bem? 😊\n"
            "Notamos que seu pagamento está pendente. Por favor, regularize para continuar aproveitando nossos serviços. "
            "Se precisar de ajuda, estamos à disposição! 💳"
        )
        mensagens_pendentes.append({"telefone": numero_formatado, "mensagem": mensagem})
    else:
        mensagem = (
            f"Olá {nome}, tudo bem? 😊\n"
            "Passando para te lembrar das novidades desta semana! Fique ligado(a)! 🚀"
        )
        mensagens_ativos.append({"telefone": numero_formatado, "mensagem": mensagem})

# === FUNÇÃO PARA ENVIAR MENSAGENS COM TEMPO ALEATÓRIO ENTRE ELAS ===

def enviar_mensagens(lista_mensagens, grupo_nome):
    print(f"\n🟢 Iniciando envio de mensagens para {grupo_nome} ({len(lista_mensagens)} alunos).")
    base_hora = time.localtime().tm_hour
    base_minuto = time.localtime().tm_min + 1  # começa 1 minuto depois do horário atual

    for i, m in enumerate(lista_mensagens):
        numero = m["telefone"]
        texto = m["mensagem"]

        hora = base_hora
        minuto = base_minuto + i

        print(f"\n📤 Enviando para {numero} às {hora}:{minuto} ...\nMensagem:\n{texto}\n")

        try:
            pywhatkit.sendwhatmsg(
                phone_no=numero,
                message=texto,
                time_hour=hora,
                time_min=minuto,
                wait_time=10,
                tab_close=True,
                close_time=3
            )
        except Exception as e:
            print(f"Erro ao enviar para {numero}: {e}")

        espera = random.randint(TEMPO_MINIMO, TEMPO_MAXIMO)
        print(f"Aguardando {espera} segundos antes da próxima mensagem...")
        time.sleep(espera)

# === EXECUTA O ENVIO ===

total_enviadas = 0

if mensagens_ativos:
    enviar_mensagens(mensagens_ativos, "Alunos Ativos (Pagamento em dia)")
    total_enviadas += len(mensagens_ativos)

if mensagens_pendentes:
    enviar_mensagens(mensagens_pendentes, "Alunos com Pagamento Pendente")
    total_enviadas += len(mensagens_pendentes)

print("\n✅ Todos os envios foram agendados!")

# === ENVIAR MENSAGEM DE RESUMO PARA O NÚMERO ESPECÍFICO ===

hora = time.localtime().tm_hour
minuto = time.localtime().tm_min + 1

mensagem_resumo = (
    f"✅ Envio de mensagens concluído!\n\n"
    f"Total de mensagens enviadas: {total_enviadas}\n"
    f"Alunos ativos (pagamento em dia): {len(mensagens_ativos)}\n"
    f"Alunos com pagamento pendente: {len(mensagens_pendentes)}"
)

print(f"\n📤 Enviando mensagem de resumo para {NUMERO_RESUMO} às {hora}:{minuto}...\nMensagem:\n{mensagem_resumo}\n")

try:
    pywhatkit.sendwhatmsg(
        phone_no=NUMERO_RESUMO,
        message=mensagem_resumo,
        time_hour=hora,
        time_min=minuto,
        wait_time=10,
        tab_close=True,
        close_time=3
    )
except Exception as e:
    print(f"Erro ao enviar mensagem de resumo: {e}")
