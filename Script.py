import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import urllib

def iniciar_navegador():
    """
    Esta função inicia uma instância do navegador Google Chrome usando o Selenium WebDriver.
    Ela abre o WhatsApp Web e aguarda o usuário realizar o login por meio do QR Code.
    
    Retorno:
        navegador (webdriver.Chrome): Instância do navegador Chrome após o login no WhatsApp Web.
    """
    navegador = webdriver.Chrome()
    navegador.get("https://web.whatsapp.com/")
    print("Aguardando login no WhatsApp Web...")
    time.sleep(15)  # Pausa para o usuário escanear o QR Code e logar no WhatsApp Web
    return navegador

def enviar_mensagem(navegador, pessoa, numero, mensagem):
    """
    Envia uma mensagem para um número de WhatsApp específico.
    
    Parâmetros:
        navegador (webdriver.Chrome): Instância do navegador já logada no WhatsApp Web.
        pessoa (str): Nome da pessoa para a qual a mensagem será enviada.
        numero (str): Número de telefone no formato internacional (com DDI e DDD).
        mensagem (str): Texto da mensagem a ser enviada.
    """
    try:
        # Codifica a mensagem para ser usada na URL do WhatsApp
        texto = urllib.parse.quote(f"Oi {pessoa}! {mensagem}")
        # Cria o link para abrir a conversa com o contato no WhatsApp Web
        link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
        navegador.get(link)
        print(f"Mandando mensagem para {pessoa} ({numero})")
        
        # Aguarda até que o botão de enviar mensagem esteja clicável
        WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@data-testid='send']"))
        )
        time.sleep(3)  # Espera adicional para garantir o carregamento

    except Exception as e:
        print(f"Erro ao enviar mensagem para {pessoa}: {e}")

def anexar_arquivo(navegador, caminho_imagem):
    """
    Anexa e envia um arquivo para o contato atual no WhatsApp Web.
    
    Parâmetros:
        navegador (webdriver.Chrome): Instância do navegador já logada no WhatsApp Web.
        caminho_imagem (str): Caminho completo para o arquivo que será anexado e enviado.
    """
    try:
        # Aguarda até que o botão de anexar esteja clicável
        anexar_btn = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='main']/footer//div[@title='Anexar']"))
        )
        anexar_btn.click()  # Clica no botão de anexar
        time.sleep(3)
        
        # Localiza o input para enviar o arquivo
        file_input = WebDriverWait(navegador, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@type="file"]'))
        )
        file_input.send_keys(caminho_imagem)  # Envia o arquivo
        time.sleep(3)

        # Aguarda até que o botão de enviar esteja clicável e então clica para enviar o arquivo
        enviar_btn = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='app']/div/div[2]/div[2]/div[2]/span/div/div/div/div[2]/div/div[2]/div[2]/div/div/span"))
        )
        enviar_btn.click()
        print("Arquivo anexado e enviado com sucesso.")
        time.sleep(3)

    except Exception as e:
        print(f"Erro ao anexar arquivo: {e}")

def main():
    """
    Função principal do script que coordena o processo de leitura dos contatos e envio de mensagens.
    """
    # Carrega o arquivo Excel contendo os contatos e as mensagens
    contatos_df = pd.read_excel("./Arquivo.xlsx")
    
    # Inicia o navegador e faz login no WhatsApp Web
    navegador = iniciar_navegador()

    # Itera sobre cada linha do DataFrame para enviar mensagens e arquivos
    for i, mensagem in contatos_df.iterrows():
        pessoa = mensagem["Pessoa"]
        numero = mensagem["Numero"]
        mensagem_texto = mensagem["Mensagem"]
        
        # Envia a mensagem para o contato
        enviar_mensagem(navegador, pessoa, numero, mensagem_texto)
        
        # Caminho do arquivo a ser anexado
        caminho_imagem = "C:\\Users\\Ryan\\Documents\\Script\\script-wpp\\Imagem.jpg"
        anexar_arquivo(navegador, caminho_imagem)

    # Fecha o navegador ao final do processo
    navegador.quit()

if __name__ == "__main__":
    main()
