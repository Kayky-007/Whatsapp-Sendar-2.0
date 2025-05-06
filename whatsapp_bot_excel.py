import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
import os

# Cria o serviço com o ChromeDriver correto (automático)
service = Service(ChromeDriverManager().install())

# Inicializa o driver com esse serviço
driver = webdriver.Chrome(service=service)


# Abre o WhatsApp Web
driver.get("https://web.whatsapp.com/")
print("Por favor, escaneie o QR Code para entrar no WhatsApp Web.")
while len(driver.find_elements(By.ID, 'side')) < 1:
    time.sleep(1)
time.sleep(2)

# Carrega os dados do Excel
contatos = pd.read_excel("Enviar.xlsx")

# Função principal de envio
def enviar_mensagem_por_numero(numero, mensagem, arquivo=None):
    try:
        # Monta a URL e acessa o número no WhatsApp Web
        url = f"https://web.whatsapp.com/send?phone={numero}&text={mensagem}"
        driver.get(url)
        while len(driver.find_elements(By.ID, 'side')) < 1:
            time.sleep(1)
        time.sleep(5)

        if arquivo and os.path.exists(arquivo):
            # Clica no botão de anexar
            botao_anexar = driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div/button/span')
            botao_anexar.click()
            time.sleep(1)

            # Faz o upload do arquivo
            campo_upload = driver.find_element(By.XPATH, '//*[@id="app"]/div/span[5]/div/ul/div/div/div[2]/li/div/input')
            campo_upload.send_keys(os.path.abspath(arquivo))
            time.sleep(3)

            # Clica no botão de enviar anexo
            botao_enviar = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span')
            botao_enviar.click()
            print(f"Mensagem e arquivo enviados para {numero}")

        elif mensagem:
            # Clica no botão de enviar mensagem
            botao_enviar_mensagem = driver.find_element(By.XPATH, "//*[@id='main']/footer/div[1]/div/span/div/div[2]/div[2]/button/span")
            botao_enviar_mensagem.click()
            print(f"Mensagem enviada para {numero}: {mensagem}")

        time.sleep(1)

    except Exception as e:
        print(f"Erro ao enviar mensagem para {numero}: {e}")

# Loop para envio automático com base no Excel
for index, linha in contatos.iterrows():
    nome = linha.get('Pessoa', '')  
    numero = str(linha['Número']).strip().replace("+", "").replace(" ", "")
    mensagem = str(linha['Mensagem']).strip()
    arquivo = str(linha.get('Arquivo', '')).strip()

    if nome:
        mensagem = f"Olá {nome}, {mensagem}"

    print(f"Enviando para {nome if nome else numero}...")
    enviar_mensagem_por_numero(numero, mensagem, arquivo)
driver.quit()
