# Importa as bibliotecas
import time
import sys
import pandas as pd
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By



# Variáveis
lista = []
i = 1


# Abre a planilha com os dados de login
try:
    df = pd.read_excel('usuarios.xlsx') 
except:
    lista.append('Erro ao acessar a planilha')
    sys.exit()


# Instala a extensão para o chrome
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)
driver.maximize_window()


# Entra no site
driver.get('https://uptodate.ebserh.gov.br/')
time.sleep(30)


# Iterar o dataframe
for index, dado in df.iterrows():
    i += 1
    print(f"Usuário {i} - {dado['usuario']}")

    # Redireciona para o login
    try:
        driver.get('https://www.uptodate.com/login')
        time.sleep(5)
    except:
        lista.append(f"Erro no usuário {i} - {dado['usuario']}")
        continue


    # Aceita os cookies
    try:
        driver.find_element(By.ID, 'onetrust-accept-btn-handler').click()
        time.sleep(2)
    except:
        pass


    # Clica em Aceitar os termos
    try:
        driver.find_element(By.ID, 'accept-license').click()
        time.sleep(2)
    except:
        pass


    # Clica em Ok
    try:
        driver.find_element(By.XPATH, '//*[@id="vue-portal-target-4lVrubU1FpZ-OhkcfVPL9"]/div[2]/div/div/div/div[3]/div/button').click()
        time.sleep(2)
    except:
        pass
    

    # Faz o login
    try:
        driver.find_element(By.ID, 'userName').send_keys(dado['usuario'])
        time.sleep(1)

        driver.find_element(By.ID, 'password').send_keys(dado['senha'])
        time.sleep(1)

        driver.find_element(By.ID, 'btnLoginSubmit').click()
        time.sleep(5)
    except:
        lista.append(f"Erro no usuário {i} - {dado['usuario']}")
        continue


    # Faz o logout
    try:
        sair = driver.find_element(By.CLASS_NAME, 'utd-log-inout')
        sair.find_element(By.TAG_NAME, 'a').click()
        time.sleep(5)
    except:
        lista.append(f"Erro no usuário {i} - {dado['usuario']}")
        continue


# Fecha o driver
driver.quit()


# Gera um txt com os dados do processamento
with open(f"Processamento_{datetime.now().strftime('%d-%m-%Y_%H-%M-%S')}.txt", "w") as arquivo:
    if len(lista) == 0:
        arquivo.write("Nenhum erro encontrado!")
    else:
        for item in lista:
            arquivo.write(item + "\n")
