from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import undetected_chromedriver as uc
import requests
import time
import openpyxl
from openpyxl import load_workbook

excel = load_workbook('C:\\Users\\Terminal_6\\Downloads\\Scraping\\Jusbrasil - Under Armour teste.xlsx')
sheet = excel.active
sheet['A1'] = 'Numero do processo'
sheet['A1'].font = openpyxl.styles.Font(bold=True)
sheet['B1'] = 'Nome do processo'
sheet['B1'].font = openpyxl.styles.Font(bold=True)
sheet['C1'] = 'Tribunal'
sheet['C1'].font = openpyxl.styles.Font(bold=True)
sheet['D1'] = 'Localidade'
sheet['D1'].font = openpyxl.styles.Font(bold=True)
sheet['E1'] = 'Procedimento'
sheet['E1'].font = openpyxl.styles.Font(bold=True)

driver = uc.Chrome()
driver.implicitly_wait(5)
url = 'https://www.jusbrasil.com.br/processos/nome/226787685/under-armour-br-hobby-ua-brasil-comercio-e-distribuicao-de-artigos-esportivos-ltda'
driver.get(url)
driver.implicitly_wait(5)

# SCROLL SCRIPT
SCROLL_PAUSE_TIME = 0.5

# Get scroll height
last_height = driver.execute_script("return document.body.scrollHeight")

while True:
    # Scroll down to bottom
    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    # Wait to load page
    time.sleep(SCROLL_PAUSE_TIME)

    # Calculate new scroll height and compare with last scroll height
    new_height = driver.execute_script("return document.body.scrollHeight")
    if new_height == last_height:
        break
    last_height = new_height
# SCROLL SCRIPT END

soup = BeautifulSoup(driver.page_source, 'html.parser')
table = soup.find('ul', class_="InfiniteList LawsuitList-list")

for i in table:
    numeroprocesso = i.find(
        'span', class_="LawsuitCardPersonPage-header-processNumber")
    if numeroprocesso:
        numeroprocesso = numeroprocesso.text
    else:
        numeroprocesso = "Indefinido"

    nomeprocesso = i.find(
        'strong', class_="LawsuitCardPersonPage-header-processInvolved")
    if nomeprocesso:
        nomeprocesso = nomeprocesso.text
    else:
        nomeprocesso = "Indefinido"

    tribunal = i.find('p', role="body-court")
    if tribunal:
        tribunal = tribunal.text
    else:
        tribunal = "Indefinido"

    localidade = i.find(
        'span', class_="LawsuitCardPersonPage-body-row-item-textSpan")
    if localidade:
        localidade = localidade.text
    else:
        localidade = "Indefinido"

    procedimento = i.find('p', role="body-kind")
    if procedimento:
        procedimento = procedimento.text
    else:
        procedimento = "Indefinido"
    sheet.append([numeroprocesso, nomeprocesso,
                 tribunal, localidade, procedimento])

excel.save('C:\\Users\\Terminal_6\\Downloads\\Scraping\\Jusbrasil - Under Armour teste.xlsx')
print("Finished")
driver.quit()

# https://openpyxl.readthedocs.io/en/stable/defined_names.html
# 1. Nike do Brasil Comércio e Participações Ltda.
# 2. Adidas do Brasil Ltda.
# 3. Puma do Brasil Ltda.
# 4. Reebok Produtos Esportivos Ltda.
# 5. Asics Brasil, Distribuição e Comércio de Artigos Esportivos Ltda.
# 6. Under Armour Brasil Comércio e Distribuição de Artigos Esportivos Ltda
# https://stackoverflow.com/questions/51122855/openpyxl-xlsx-results-are-appending-instead-of-overwriting
# número do processo, partes envolvidas, tribunal, localidade, UF, classe ou procedimento
# With Selenium:
# numeroprocesso = driver.find_elements(
#     By.CLASS_NAME, 'LawsuitCardPersonPage-header-processNumber')
# for i in numeroprocesso:
#     print(i.text)

# nomeprocesso = driver.find_elements(
#     By.CLASS_NAME, 'LawsuitCardPersonPage-header-processInvolved')
# for i in nomeprocesso:
#     print(i.text)

# tribunal = driver.find_elements(
#     By.XPATH, "//p[@role='body-court']")
# for i in tribunal:
#     print(i.text)

# localidade = driver.find_elements(
#     By.XPATH, "//span[@class='LawsuitCardPersonPage-body-row-item-textSpan']")
# for i in localidade:
#     print(i.text)

# procedimento = driver.find_elements(
#     By.XPATH, "//p[@role='body-kind']")
# for i in procedimento:
#     print(i.text)
