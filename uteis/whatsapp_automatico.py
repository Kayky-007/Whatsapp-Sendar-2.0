# uteis/whatsapp_automatico.py
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os

class WhatsAppAutomatico:
    def __init__(self):
        self.driver = None
        self.wait = None

    def iniciar(self):
        caminho = './chromedriver/chromedriver.exe'
        if not os.path.exists(caminho):
            print(f"[ERRO] chromedriver.exe não encontrado: {caminho}")
            return False

        service = Service(caminho)
        chrome_options = Options()
        chrome_options.add_argument('--disable-gcm')
        chrome_options.add_argument('--log-level=3')
        chrome_options.add_argument('--disable-logging')
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')

        try:
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.wait = WebDriverWait(self.driver, 30)
            print("[INFO] Chrome iniciado com sucesso!")
        except Exception as e:
            print(f"[ERRO] Falha ao iniciar Chrome: {e}")
            return False

        self.driver.get("https://web.web.whatsapp.com/")
        print("[INFO] Escaneie o QR Code...")

        try:
            self.wait.until(EC.presence_of_element_located((By.ID, "side")))
            time.sleep(3)
            print("[SUCESSO] WhatsApp conectado!")
            return True
        except Exception as e:
            print(f"[ERRO] QR Code não escaneado: {e}")
            return False

    def enviar(self, numero, mensagem=None, arquivo=None, legenda=None, delay=2):
        try:
            numero = numero.replace(" ", "").replace("-", "")
            print(f"[ENVIANDO] {numero}")

            if mensagem and not arquivo:
                msg = mensagem.replace("\n", "%0A").replace(" ", "%20")
                url = f"https://web.whatsapp.com/send?phone={numero}&text={msg}"
                self.driver.get(url)
                self.wait.until(EC.presence_of_element_located((By.ID, "main")))
                time.sleep(3)

                xpath_enviar = '//*[@id="main"]/footer/div[1]/div/span/div/div[2]/div/div[4]/div/span/div/div/div[1]/div[1]/span'
                botao = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath_enviar)))
                botao.click()
                print("[SUCESSO] Mensagem enviada!")
                time.sleep(delay)
                return "Enviado"

        except Exception as e:
            print(f"[ERRO] {str(e)}")
            return f"Erro: {str(e)}"

    def fechar(self):
        if self.driver:
            self.driver.quit()
            self.driver = None
            self.wait = None
            print("[FECHADO] Navegador fechado.")