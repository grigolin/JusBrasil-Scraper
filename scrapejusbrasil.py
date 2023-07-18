from bs4 import BeautifulSoup
import undetected_chromedriver as uc
import time
import openpyxl

# change the url and the title of the excel file
url = 'https://www.jusbrasil.com.br/processos/nome/226787685/under-armour-br-hobby-ua-brasil-comercio-e-distribuicao-de-artigos-esportivos-ltda'
title = 'Jusbrasil - Under Armour'

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = title
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
driver.get(url)
driver.implicitly_wait(3)

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

# número do processo, partes envolvidas, tribunal, localidade e procedimento
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

excel.save(f"{title}.xlsx")
print("Finished")
driver.quit()