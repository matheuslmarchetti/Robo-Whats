from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ChromeOptions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from glob import glob
from datetime import datetime
import time
from urllib.parse import urlencode
import pandas as pd
import os
import subprocess
import sys

def save_progress(message_list, file_path):
    """Função auxiliar para salvar o progresso em um arquivo Excel."""
    message_list.to_excel(file_path, index=False)

def mandar_anexos(driver, imagens: list, pdfs: list, audios: list, videos: list, grupo : str, texto = None) -> str:
    # Lista de possíveis XPATHs
    xpaths = [
        '//span[@data-icon="plus"]',
        '//span[@data-icon="attach-menu-plus"]'
    ]
    time.sleep(2)
    if texto:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.XPATH, "//span[@data-icon='send']"))
            ).click()
        
    for imagem in imagens:
        time.sleep(1)
        try:
            driver.maximize_window()
        except:
            print('Falhou Maximizar')
        # Tenta clicar em um dos botões disponíveis
        for xpath in xpaths:
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath))).click()
                print(f'Botão encontrado e clicado: {xpath}')
                break  # Sai do loop se encontrar um botão válido
            except:
                print(f'Não encontrou o botão: {xpath}')
        # Coloca o caminho da imagem
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))).send_keys(imagem)
        # Espera 3 Segundos
        time.sleep(3)
        # Envia a imagem
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[@data-icon='send']"))).click()
        # Printa a data e hora de envio do arquivo
        agora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        print(fr'Enviado imagem {imagem} as {agora}')
    
    for pdf in pdfs:
        time.sleep(1)
        try:
            driver.maximize_window()
        except:
            print('Falhou Maximizar')
        # Tenta clicar em um dos botões disponíveis
        for xpath in xpaths:
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath))).click()
                print(f'Botão encontrado e clicado: {xpath}')
                break  # Sai do loop se encontrar um botão válido
            except:
                print(f'Não encontrou o botão: {xpath}')
        # Coloca o caminho da pdf
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@accept="*"]'))).send_keys(pdf)
        # Espera 3 Segundos
        time.sleep(3)
        # Envia a pdf
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[@data-icon='send']"))).click()
        # Printa a data e hora de envio do arquivo
        agora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        print(fr'Enviado pdf {pdf} as {agora}')

    for audio in audios:
        time.sleep(1)
        try:
            driver.maximize_window()
        except:
            print('Falhou Maximizar')
        # Tenta clicar em um dos botões disponíveis
        for xpath in xpaths:
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath))).click()
                print(f'Botão encontrado e clicado: {xpath}')
                break  # Sai do loop se encontrar um botão válido
            except:
                print(f'Não encontrou o botão: {xpath}')
        # Coloca o caminho da audio
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))).send_keys(audio)
        # Espera 3 Segundos
        time.sleep(3)
        # Envia a audio
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[@data-icon='send']"))).click()
        # Printa a data e hora de envio do arquivo
        agora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        print(fr'Enviado pdf {pdf} as {agora}')

    for video in videos:
        time.sleep(1)
        try:
            driver.maximize_window()
        except:
            print('Falhou Maximizar')
        # Tenta clicar em um dos botões disponíveis
        for xpath in xpaths:
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath))).click()
                print(f'Botão encontrado e clicado: {xpath}')
                break  # Sai do loop se encontrar um botão válido
            except:
                print(f'Não encontrou o botão: {xpath}')
        # Coloca o caminho da video
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))).send_keys(video)
        time.sleep(3)
        # Envia a video
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[@data-icon='send']"))).click()
        # Printa a data e hora de envio do arquivo
        agora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        print(fr'Enviado pdf {pdf} as {agora}')

    return fr'Anexos enviados com sucesso!'

def ler_base():
    base = pd.read_excel(fr'.\mensagens\mensagens.xlsx')
    # Concatenar colunas com "." no meio
    base['MENSAGEM'] = base['NOME'].astype(str).str.cat(base['TEXTO'].astype(str), sep='. ')
    base = base.loc[~base['NUMERO'].isna()].copy()
    base['NUMERO'] = base['NUMERO'].astype(str).str.replace(r'[^\d]', '', regex=True)
    
    return base

if __name__ == '__main__':
    #Configura o Navegador
    message_list = pd.read_excel(fr'mensagens\mensagens.xlsx')
    current_time = datetime.now().strftime('%d_%m_%Y-%H_%M').replace(':', '_')
    result_file_path = os.path.join("relatorios", f"resultado_de_envio_das_mensagens_{current_time}.xlsx")
    # Inicialize o arquivo de resultados, se ainda não existir
    if not os.path.exists(result_file_path):
        message_list['Status'] = 'Pendente'  # Adiciona uma coluna de Status
        save_progress(message_list, result_file_path)
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()
                        )
    )
    options = ChromeOptions()
    options.add_argument(fr'--user-data-dir={os.path.dirname(os.path.realpath(__file__))}\cache')

    #Espera logar no whats
    driver.get(fr'https://web.whatsapp.com/')
    driver.maximize_window()
    WebDriverWait(driver,120).until(EC.presence_of_element_located((By.XPATH, "//span[@data-icon='search']")))

    try:
        for linha in ler_base().itertuples():
            try:
                mensagem = urlencode({'text': linha.MENSAGEM})
                driver.get(F'https://web.whatsapp.com/send?phone=+55{linha.NUMERO}&{mensagem}')
                mandar_anexos(driver, glob(fr'{os.path.dirname(os.path.realpath(__file__))}\imagens\*'), glob(fr'{os.path.dirname(os.path.realpath(__file__))}\pdf\*'), glob(fr'{os.path.dirname(os.path.realpath(__file__))}\audio\*'), glob(fr'{os.path.dirname(os.path.realpath(__file__))}\video\*'), grupo = '', texto = mensagem)
                time.sleep(5)
                print(fr'Último Número Enviado:{linha.NUMERO}')
                result = True
            except Exception as ex:
                print(f'Erro ao enviar a mensagem para {linha.NUMERO}: {ex}')
                result = False

            status = 'Enviada' if result else 'Erro'
            message_list.loc[linha.Index, 'Status'] = status
            save_progress(message_list, result_file_path)  # Atualiza o arquivo de resultado
            
            # Interrompe o loop se o navegador for fechado
            if not driver.window_handles:
                print("Navegador foi fechado inesperadamente.")
                break

    except Exception as e:
        print(f'Erro durante o envio das mensagens: {e}')

    finally:
        # Salva qualquer progresso antes de sair em caso de exceção
        save_progress(message_list, result_file_path)
        driver.quit()

print('Terminou')
