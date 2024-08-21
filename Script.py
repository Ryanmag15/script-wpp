import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import time
import urllib

contatos_df = pd.read_excel("./Arquivo.xlsx")

navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com/")

time.sleep(15)   

#Login Feito no WPP WEB

for i, mensagem in enumerate(contatos_df['Mensagem']):

    pessoa = contatos_df.loc[i, "Pessoa"]
    numero = contatos_df.loc[i, "Numero"]
    texto = urllib.parse.quote(f"Oi {pessoa}! {mensagem}")
    link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
    navegador.get(link)

    print("Mandando mensagem para " + pessoa)
    time.sleep(15) 

    # Clique no botão "de + para abrir o botao de selecionar o anexar"
    anexar_btn = navegador.find_element(By.XPATH, "//*[@id='main']/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/div/span")
    anexar_btn.click()
    time.sleep(3)

    # Clique no botão "Anexar"
    anexar_btn = navegador.find_element(By.XPATH, "//*[@id='main']/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[2]/li/div")
    anexar_btn.click()
    time.sleep(3)

    # Enviar caminho da imagem para o elemento de entrada do tipo arquivo
    caminho_imagem = "C:\\Users\\Ryan\\Documents\\Script\\Imagem.jpg"
    file_input = WebDriverWait(navegador, 10).until(
        lambda navegador: navegador.find_element(By.XPATH, '//input[@type="file"]')
    )
    file_input.send_keys(caminho_imagem)
    time.sleep(3)

    # Clique no botão "Enviar"
    enviar_btn = navegador.find_element(By.XPATH, "//*[@id='app']/div/div[2]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span")
    enviar_btn.click()
    time.sleep(3)