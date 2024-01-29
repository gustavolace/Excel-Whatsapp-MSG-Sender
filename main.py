from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from win32com.client import Dispatch
from urllib.parse import quote
from PIL import ImageGrab
import pandas as pd
import time
import os 
import sys

script_path = sys.argv[0]
dir_atual = os.path.dirname(os.path.abspath(script_path))
xml_file = os.path.join(dir_atual, 'vba.xlsm')

appdata_dir = os.getenv('APPDATA')
chrome_profile_dir = os.path.join(appdata_dir, 'Local', 'Google', 'Chrome', 'User Data', 'Profile 1')

var = []
temp_image_path = ""
have_img = False

def find_sheet_shapes():
    global have_img
    xl = Dispatch('Excel.Application')
    wb = xl.Workbooks.Open(Filename=xml_file)
    ws = wb.Worksheets(1)
    shapes = ws.Shapes(2)
    text =  shapes.TextFrame.Characters().Text
    image_shape = ws.Shapes(3)
    image_shape.Copy()
    image = ImageGrab.grabclipboard()
    if image:
        # Salva a imagem temporariamente em um arquivo
        temp_image_path = os.path.join(os.getcwd(), 'temp_image.png')
        have_img = True
        image.save(temp_image_path)

        print("Imagem salva temporariamente com sucesso:", temp_image_path)

    print(ws.Shapes(3))
    return text

def find_contacts():
    df = pd.read_excel(xml_file, sheet_name='Planilha1', header=0, names=['Nome', 'Telefone'])

    # Capturar dados da coluna A (nomes) e coluna B (telefones) apenas se as células não estiverem vazias
    nomes = df['Nome'].dropna().tolist()
    telefones = df['Telefone'].dropna().tolist()

    # Exibir os dados capturados
    return zip(nomes, telefones)

def make_url():
    contacts  = find_contacts()
    text = find_sheet_shapes()
    for nome, telefone in contacts :
        client_text = text.replace("$cl", nome)
        url = f"https://web.whatsapp.com/send/?phone={telefone}&text={quote(client_text)}"
        var.append(url)
        print(url)

make_url()

options = webdriver.ChromeOptions()
options.add_argument("--headless")
options.add_argument(f"--user-data-dir={chrome_profile_dir}")
user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.50 Safari/537.36'
options.add_argument(f'user-agent={user_agent}')
driver = webdriver.Chrome(options=options)


def navigator_web_driver(var):
    driver.get(f"{var}")

    try:

        if have_img is True:
            attach_icon = WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-icon='attach-menu-plus']")))
            attach_icon.click()
            input_file = WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[accept='image/*,video/mp4,video/3gpp,video/quicktime']")))
            input_file.send_keys(rf"{os.path.abspath(temp_image_path)}\temp_image.png")

            WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-icon='send']")))
            print("Elemento span[data-icon='send'] encontrado")

            time.sleep(1)
            actions = ActionChains(driver)
            actions.send_keys(Keys.ENTER)
            actions.perform()
            time.sleep(2)
        
        else: 
            element = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-icon='send']")))
            print("Elemento span[data-icon='send'] encontrado")
            element.click()
            time.sleep(2)

    except Exception as e:
        print(f"Erro durante a interação: {e}")


for item in var:
    navigator_web_driver(item)
# Fecha navegador 
driver.quit()