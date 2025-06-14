import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
import time
import random

# === CONFIGURAÇÕES ===
CHROMEDRIVER_PATH = r"C:\Users\Amanda\OneDrive\Documents\GitHub\my-codes\mensagem_whatsapp\chromedriver-win64\chromedriver-win64\chromedriver.exe"  # Caminho para o chromedriver
PLANILHA = r"C:\Users\Amanda\OneDrive\Documents\mensagem_whatsapp\alunos.xlsx"  # Caminho da planilha
TEMPO_MIN = 5
TEMPO_MAX = 30
NUMERO_RESUMO = "+5583999944947"  # Número que recebe o resumo final

# === INICIAR O NAVEGADOR ===
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service)
driver.get("https://web.whatsapp.com")
input("\n✅ Escaneie o QR Code e pressione ENTER para continuar...\n")

# === FUNÇÃO PARA ENVIAR MENSAGEM ===
def enviar_mensagem(numero, mensagem):
    try:
        url = f"https://web.whatsapp.com/send?phone={numero}&text={mensagem}"
        driver.get(url)
        time.sleep(10)  # Tempo para o WhatsApp carregar
        caixa = driver.find_element(By.XPATH, '//div[@title="Digite uma mensagem"]')
        caixa.send_keys(Keys.ENTER)
        print(f"✅ Mensagem enviada para {numero}")
    except Exception as e:
        print(f"❌ Erro ao enviar para {numero}: {e}")

# === LEITURA DA PLANILHA ===
try:
    df = pd.read_excel(PLANILHA)
except Exception as e:
    print(f"❌ Erro ao abrir a planilha: {e}")
    driver.quit()
    exit()

ativos = 0
pendentes = 0
total = 0

# === ENVIO DAS MENSAGENS ===
for _, row in df.iterrows():
    nome = row.get("Nome", "Aluno")
    telefone = str(row.get("Telefone", "")).strip()
    status = str(row.get("Status", "")).strip().lower()
    pagamento = str(row.get("Pagamento", "")).strip().lower()

    if not telefone or status != "ativo":
        continue

    numero = "+55" + telefone

    if pagamento in ["pendente", "atrasado"]:
        mensagem = (
            f"Olá {nome}, tudo bem? 😊\n"
            "Notamos que seu pagamento está pendente. Por favor, regularize o quanto antes para continuar aproveitando nossos serviços. 💳"
        )
        pendentes += 1
    else:
        mensagem = (
            f"Olá {nome}, tudo bem? 😊\n"
            "Estamos passando para compartilhar as novidades desta semana. Fique de olho! 🚀"
        )
        ativos += 1

    enviar_mensagem(numero, mensagem)
    espera = random.randint(TEMPO_MIN, TEMPO_MAX)
    print(f"⏳ Aguardando {espera} segundos antes da próxima mensagem...\n")
    time.sleep(espera)
    total += 1

# === MENSAGEM DE RESUMO FINAL ===
resumo = (
    f"✅ Envio concluído!\n"
    f"Total de mensagens: {total}\n"
    f"Alunos ativos: {ativos}\n"
    f"Alunos pendentes: {pendentes}"
)

print("\n📢 Enviando resumo final...\n")
enviar_mensagem(NUMERO_RESUMO, resumo)

# === FINALIZAÇÃO ===
print("\n✅ Processo encerrado. Navegador será fechado.")
driver.quit()
