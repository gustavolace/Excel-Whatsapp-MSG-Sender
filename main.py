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

log_file_path = os.path.join(dir_atual, "logs/log.log")
sys.stdout = open(log_file_path, 'w')

xml_file = os.path.join(dir_atual, 'vba.xlsm')

appdata_dir = os.getenv('APPDATA')
chrome_profile_dir = os.path.join(appdata_dir, 'Local', 'Google', 'Chrome', 'User Data', 'Profile 1')

var = []
temp_image_path = ""
have_img = False
after_img = False

def find_sheet_shapes():

    # WorkSheet
    global have_img, after_img
    xl = Dispatch('Excel.Application')
    wb = xl.Workbooks.Open(Filename=xml_file)
    ws = wb.Worksheets(1)

    # Text
    text_shapes = ws.Shapes(2)
    text =  text_shapes.TextFrame.Characters().Text
    
    ## Checkbox
    checkbox_afeter_shape = ws.Shapes(3)
    if checkbox_afeter_shape.OLEFormat.Object.Value == 1:
        after_img = True


    ## Get Image
    if 4 <= ws.Shapes.Count:
        image_shape = ws.Shapes(4)
        image_shape.Copy()
        image = ImageGrab.grabclipboard()
        temp_image_path = os.path.join(os.getcwd(), 'assets', 'temp_image.png')
        have_img = True
        image.save(temp_image_path)

        print("Imagem salva temporariamente com sucesso:", temp_image_path)

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

make_url()

options = webdriver.ChromeOptions()
#options.add_argument("--headless")
options.add_argument(f"--user-data-dir={chrome_profile_dir}")
user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.50 Safari/537.36'
options.add_argument(f'user-agent={user_agent}')
driver = webdriver.Chrome(options=options)


def css_selector(selector):
    return WebDriverWait(driver, 40).until(EC.presence_of_element_located((By.CSS_SELECTOR, f"{selector}")))

def send_img(send_log):
    attach_icon = css_selector("span[data-icon='attach-menu-plus']")
    attach_icon.click()
    input_file = css_selector("input[accept='image/*,video/mp4,video/3gpp,video/quicktime']")
    input_file.send_keys(rf"{os.path.abspath(temp_image_path)}\temp_image.png")

    time.sleep(1)
    actions = ActionChains(driver)
    actions.send_keys(Keys.ENTER)
    actions.perform()
    time.sleep(2)
    print(send_log)


def navigator_web_driver(var):
    driver.get(f"{var}")
    send_log = "Mensagem enviada"

    try:
        if have_img is True:
            send_button = css_selector("span[data-icon='send']")
            if after_img:
                send_button.click()
                print(send_log)
                time.sleep(1)
                send_img("Imagem enviada")

            else:
                send_img(send_log)

        else: 
            css_selector("span[data-icon='send']").click()
            print(send_log)
            time.sleep(2)

    except Exception as e:
        print(f"Erro durante a interação: {e}")


for item in var:
    navigator_web_driver(item)
# Fecha navegador 
driver.quit()
sys.stdout.close()