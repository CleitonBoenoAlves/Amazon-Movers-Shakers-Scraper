# Import`s
from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
from time import sleep
from datetime import datetime
import pandas as pd
import re

# Armazena os dados de cada carrossel
dados_carrosseis = {}

# Instalação do serviço, caso não houver no path
service = Service(GeckoDriverManager().install())
driver = webdriver.Firefox(service=service)
driver.maximize_window()

# Abrindo o navegador com a URL alvo
url_alvo = "https://www.amazon.com.br/gp/movers-and-shakers/"
driver.get(url_alvo)

# Pegando todos os carrosseis
sleep(5)
xpath = "//div[contains(@id,'anonCarousel')]"
carrosseis = driver.find_elements(By.XPATH, xpath)

# Realizando scroll por conta do lazy loading
i = 0

while i < len(carrosseis):
    carrossel = carrosseis[i]
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", carrossel)
    sleep(1)

    # atualiza a lista porque novos carrosseis podem ter sido injetados
    carrosseis = driver.find_elements(By.XPATH, "//div[contains(@id,'anonCarousel')]")
    i += 1

# Pega os titulos dos carrosseis
titulos_carrosseis = driver.find_elements(By.XPATH, "//h2[@class='a-carousel-heading a-inline-block']")
titulos_carrosseis = [t.text.replace("Produtos em alta em ", "") for t in titulos_carrosseis]


# Pegandos os cards do carrossel
carrosseis = driver.find_elements(By.XPATH, '//div[contains(@id, "anonCarousel")]')
for i in range(len(carrosseis)):
    # remove caracteres inválidos e limita 31 caracteres
    titulo_limpo = "".join(c for c in titulos_carrosseis[i] if c not in '/\\?*[]:')[:31]

    carrossel = carrosseis[i]
    dados = []

    while True:
        xpath = './/li[@class="a-carousel-card"]'
        cards = carrossel.find_elements(By.XPATH, xpath)

        for card in cards:
            try:
                produto = card.find_element(By.XPATH, ".//a[contains(@class,'a-link-normal')]/span/div").text
                preco = card.find_element(By.XPATH, ".//span[contains(@class,'a-size-base')]/span").text
                link = card.find_element(By.XPATH, ".//a[@class='a-link-normal aok-block']").get_attribute("href")
                avaliacoes = card.find_element(By.XPATH, ".//span[@class='a-size-small']").text

                dados.append({
                    "produto": produto,
                    "preco":preco,
                    "link":link,
                    "avaliacoes":avaliacoes
                })
            except NoSuchElementException:
                pass

        posinset = cards[-1].get_attribute('aria-posinset')
        setsize = cards[-1].get_attribute('aria-setsize')

        if posinset == setsize:
            break
        else:
            xpath = './../..//a[@class="a-button a-button-image a-carousel-button a-carousel-goto-nextpage"]'
            botao_next = carrossel.find_element(By.XPATH, xpath)
            botao_next.click()
            sleep(2)
    
    dados_carrosseis[titulo_limpo] = pd.DataFrame(dados)

    carrosseis = driver.find_elements(By.XPATH, '//div[contains(@id, "anonCarousel")]')

# Fecha o navegador
driver.quit()

# Pega a data e hora atual
data_hora = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

# Nome arquivo
nome_arquivo = f"produtos_amazon_{data_hora}.xlsx"

# Criando Excel com várias abas
with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
    for categoria, itens in dados_carrosseis.items():
        df = pd.DataFrame(itens)
        
        # O nome da aba não pode ter mais de 31 caracteres nem caracteres especiais
        aba = categoria[:31].replace("/", "-").replace("\\", "-")
        df.to_excel(writer, sheet_name=aba, index=False)