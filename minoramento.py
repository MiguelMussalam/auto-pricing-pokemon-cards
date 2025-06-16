import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import quote
import time
import re

colunas = ['nome', 'numero', 'condicao', 'preco', 'total', 'data']
album = pd.read_excel('Preço do album.xlsx', skiprows= 4, usecols= 'E, F, G, H, J, K', names= colunas)
total = 0
data = album.data[0]

url_base = 'https://www.ligapokemon.com.br/?view=cards/card&card='
driver = webdriver.Chrome()
driver.maximize_window()
wait = WebDriverWait(driver, 10)

for i, carta in album.iterrows():
    if pd.isna(carta['condicao']):
        continue

    nome_carta = f'{carta.nome.strip()} ({carta.numero.strip()})'
    condicao = carta.condicao
    print(f'Nome: {nome_carta} condição: {condicao}')

    url = f'{url_base}{quote(nome_carta, safe="()/'")}'
    print(url)
    driver.get(url)

    # - Marca o checkbox do estado NM
    map_de_condicao = {
    'M': 'filter_4_0',   # M Nova
    'NM': 'filter_4_1',  # NM Praticamente Nova ou superior
    'SP': 'filter_4_2',  # SP Usada Levemente ou superior
    'MP': 'filter_4_3',  # MP Usada Moderadamente ou superior
    'HP': 'filter_4_4',  # HP Muito Usada ou superior
    'D': 'filter_4_5'    # D Danificada ou superior
    }

    try:
        # Checkbox de condicao
        checkbox_nm = wait.until(EC.element_to_be_clickable((By.ID, map_de_condicao[condicao])))
        driver.execute_script("arguments[0].scrollIntoView(true);", checkbox_nm)
        time.sleep(0.2)
        checkbox_nm.click()
        wait.until(EC.invisibility_of_element_located((By.ID, 'layer-loading')))

        # Botão de comprar 
        marketplace_stores = driver.find_element(By.ID, 'marketplace-stores')
        primeiro_div = marketplace_stores.find_element(By.XPATH, './div[1]')
        botao_container = primeiro_div.find_element(By.CLASS_NAME, "buttons")
        botao = botao_container.find_element(By.XPATH, './div[1]')
        driver.execute_script("arguments[0].click();", botao)
        time.sleep(0.5)

        # Ir pro carrinho e pegar preço
        driver.get('https://www.ligapokemon.com.br/?view=mp/carrinho')
        preco = wait.until(EC.visibility_of_element_located((By.ID, 'sum_preco_0'))).text
        print(preco)

        # Tirar do carrinho e aceitar alerta
        tirar_carrinho = wait.until(EC.element_to_be_clickable((By.CLASS_NAME, 'btn-circle.remove.delete.item-delete')))
        driver.execute_script("arguments[0].scrollIntoView(true);", tirar_carrinho)
        tirar_carrinho.click()
        wait = WebDriverWait(driver, timeout=2)
        alert = wait.until(lambda d : d.switch_to.alert)
        text = alert.text
        alert.accept()
    except:
        print(f'A carta {nome_carta} da condicao {condicao} não existe no site.')
        continue

driver.quit()