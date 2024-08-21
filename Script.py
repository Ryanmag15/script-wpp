import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import urllib

def iniciar_navegador():
    navegador = webdriver.Chrome()
    navegador.get("https://web.whatsapp.com/")
    print("Aguardando login no WhatsApp Web...")
    time.sleep(15)  # Tempo para o usuário escanear o QR Code e logar
    return navegador

def enviar_mensagem(navegador, pessoa, numero, mensagem):
    try:
        texto = urllib.parse.quote(f"Oi {pessoa}! {mensagem}")
        link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
        navegador.get(link)
        print(f"Mandando mensagem para {pessoa} ({numero})")
        
        WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@data-testid='send']"))
        )
        time.sleep(3)  # Pode ajustar conforme necessário

    except Exception as e:
        print(f"Erro ao enviar mensagem para {pessoa}: {e}")

def anexar_arquivo(navegador, caminho_imagem):
    try:
        anexar_btn = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='main']/footer//div[@title='Anexar']"))
        )
        anexar_btn.click()
        time.sleep(3)
        
        file_input = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="file"]'))
        )
        file_input.send_keys(caminho_imagem)
        time.sleep(3)

        enviar_btn = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='app']/div/div[2]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span"))
        )
        enviar_btn.click()
        print("Arquivo anexado e enviado com sucesso.")
        time.sleep(3)

    except Exception as e:
        print(f"Erro ao anexar arquivo: {e}")

def main():
    contatos_df = pd.read_excel("./Arquivo.xlsx")
    navegador = iniciar_navegador()

    for i, mensagem in contatos_df.iterrows():
        pessoa = mensagem["Pessoa"]
        numero = mensagem["Numero"]
        mensagem_texto = mensagem["Mensagem"]
        
        enviar_mensagem(navegador, pessoa, numero, mensagem_texto)
        
        caminho_imagem = "C:\\Users\\Ryan\\Documents\\Script\\script-wpp\\Imagem.jpg"
        anexar_arquivo(navegador, caminho_imagem)

    navegador.quit()

if __name__ == "__main__":
    main()
