from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
import os       

# Configuração inicial do Selenium
chromedriver_path = './chromedriver.exe'
service = Service(chromedriver_path)
driver = webdriver.Chrome(service=service)

# Abre o WhatsApp Web
driver.get("https://web.whatsapp.com/")
print("Por favor, escaneie o QR Code para entrar no WhatsApp Web.")
while len(driver.find_elements(By.ID, 'side')) < 1:
    time.sleep(1)
time.sleep(2)

# Funções para envio de mensagens
def enviar_mensagem_com_texto_e_anexo(numero, mensagem, arquivo):
    enviar_mensagem_por_numero(numero, mensagem, arquivo)

def enviar_mensagem_com_legenda_e_anexo(numero, legenda, arquivo):
    enviar_mensagem_por_numero(numero, legenda, arquivo)

def enviar_mensagem_somente_texto(numero, mensagem):
    enviar_mensagem_por_numero(numero, mensagem)

def enviar_mensagem_somente_anexo(numero, arquivo):
    enviar_mensagem_por_numero(numero, mensagem="", arquivo=arquivo)

def enviar_mensagem_por_numero(numero, mensagem, arquivo=None, legenda=None):
    try:
        # Monta a URL e acessa o número no WhatsApp Web
        url = f"https://web.whatsapp.com/send?phone={numero}&text={mensagem}"
        driver.get(url)
        while len(driver.find_elements(By.ID, 'side')) < 1:
            time.sleep(1)
        time.sleep(2)

        if arquivo:
            # Clica no botão de anexar
            botao_anexar = driver.find_element(By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div/button/span')
            botao_anexar.click()
            time.sleep(1)

            # Faz o upload do arquivo
            campo_upload = driver.find_element(By.XPATH, '//*[@id="app"]/div/span[5]/div/ul/div/div/div[2]/li/div/input')
            campo_upload.send_keys(arquivo)
            time.sleep(3)

            # Insere a legenda, se houver
            if legenda:
                campo_legenda = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[1]/div[2]')
                campo_legenda.click()
                campo_legenda.send_keys(legenda)
                time.sleep(2)

            # Clica no botão de enviar
            botao_enviar = driver.find_element(By.XPATH, '//*[@id="app"]/div/div[3]/div/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span')
            botao_enviar.click()
            print(f"Arquivo e legenda enviados para {numero}")
        
        elif mensagem:
            # Clica no botão de enviar mensagem
            botao_enviar_mensagem = driver.find_element(By.XPATH, "//*[@id='main']/footer/div[1]/div/span/div/div[2]/div[2]/button/span")
            botao_enviar_mensagem.click()
            print(f"Mensagem enviada para {numero}: {mensagem}")

    except Exception as e:
        print(f"Erro ao enviar mensagem para {numero}: {e}")


# Sistema de menu
def menu():
    while True:
        print("\nEscolha uma opção:")
        print("1 - Enviar mensagem com texto e anexo")
        print("======================================")
        print("2 - Enviar mensagem com legenda e anexo")
        print("======================================")
        print("3 - Enviar mensagem somente com texto")
        print("======================================")
        print("4 - Enviar mensagem somente com anexo")
        print("======================================")
        print("5 - Sair")
        
        opcao = input("Digite o número da opção desejada: ")
        
        if opcao == "1":
            numero = input("Digite o número (formato internacional, ex: +5511999999999): ")
            mensagem = input("Digite a mensagem: ")
            arquivo = input("Digite o caminho do arquivo para anexo: ")
            enviar_mensagem_com_texto_e_anexo(numero, mensagem, arquivo)
        
        elif opcao == "2":
            numero = input("Digite o número (formato internacional, ex: +5511999999999): ")
            legenda = input("Digite a legenda: ")
            arquivo = input("Digite o caminho do arquivo para anexo: ")
            enviar_mensagem_com_legenda_e_anexo(numero, legenda, arquivo)

        elif opcao == "3":
            numero = input("Digite o número (formato internacional, ex: +5511999999999): ")
            mensagem = input("Digite a mensagem: ")
            enviar_mensagem_somente_texto(numero, mensagem)
        
        elif opcao == "4":
            numero = input("Digite o número (formato internacional, ex: +5511999999999): ")
            arquivo = input("Digite o caminho do arquivo para anexo: ")
            enviar_mensagem_somente_anexo(numero, arquivo)
        
        elif opcao == "5":
            print("Encerrando o programa...")
            driver.quit()  # Fecha o navegador
            break
        
        else:
            print("Opção inválida. Tente novamente.")

menu()