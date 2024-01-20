from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import psutil
from win32com.client import Dispatch
from urllib.parse import quote
import time

var = []

def is_excel_running():
    for process in psutil.process_iter(['pid', 'name']):
        if 'excel.exe' in process.info['name'].lower():
            return True
    return False
num = is_excel_running()

def find_text():
    exel_status = is_excel_running()
    xl = Dispatch('Excel.Application')
    wb = xl.Workbooks.Open(Filename=r'C:\Users\Gustavo Lacerda\Downloads\vba.xlsm', )
    ws = wb.Worksheets(1)
    shapes = ws.Shapes(2)
    text =  shapes.TextFrame.Characters().Text
    if exel_status is False:
        xl.Quit()

    print("Text Box Text:", text)
    return text

def find_contacts():
    # Ler o arquivo Excel
    df = pd.read_excel(r'C:\Users\Gustavo Lacerda\Downloads\vba.xlsm', sheet_name='Planilha1', header=0, names=['Nome', 'Telefone'])

    # Capturar dados da coluna A (nomes) e coluna B (telefones) apenas se as células não estiverem vazias
    nomes = df['Nome'].dropna().tolist()
    telefones = df['Telefone'].dropna().tolist()

    # Exibir os dados capturados
    return zip(nomes, telefones)

def make_url():
    contacts  = find_contacts()
    text = find_text()
    for nome, telefone in contacts :
        client_text = text.replace("$cl", nome)
        url = f"https://web.whatsapp.com/send/?phone={telefone}&text={quote(client_text)}"
        var.append(url)
        print(url)


make_url()

options = webdriver.ChromeOptions()
options.add_argument("--headless")
options.add_argument("--user-data-dir=C:\\Users\\Gustavo Lacerda\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 3")
user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.50 Safari/537.36'
options.add_argument(f'user-agent={user_agent}')

driver = webdriver.Chrome(options=options)

def teste(var):
    driver.get(f"{var}")

    try:
        element = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-icon='send']")))
        print("Elemento span[data-icon='send'] encontrado")
        element.click()

        time.sleep(2)
    except Exception:
        print("Erro ao encontrar elemento span[data-icon='send']:")

for item in var:
    teste(item)
# Fecha navegador 
driver.quit()
