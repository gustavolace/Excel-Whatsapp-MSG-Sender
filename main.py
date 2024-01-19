from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

options = webdriver.ChromeOptions()
#options.add_argument("--headless")
options.add_argument("--user-data-dir=C:\\Users\\Gustavo Lacerda\\AppData\\Local\\Google\\Chrome\\User Data\\Profile 3")

driver = webdriver.Chrome(options=options)

""" var = ["https://api.whatsapp.com/send?phone=5571981854906&text=iai%0Acaara", 
       "https://api.whatsapp.com/send?phone=5571981854906&text=oi%0Atudo"] """

var = [
    "https://web.whatsapp.com/send/?phone=5571981854906&text=oi%250tudo",
    "https://web.whatsapp.com/send/?phone=5571981854906&text=iai%0Acaara"
]

def teste(var):
    driver.execute_script("window.open('', '_blank');")
    driver.switch_to.window(driver.window_handles[1])
    
    driver.get(f"{var}")

    """ try:
        action_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "action-button")))
        print("Elemento 'action-button' encontrado")
        # Encontrando o elemento span dentro do elemento 'action-button'
        span = action_button.find_element(By.CLASS_NAME, '_advp')
        span.click()
    except Exception:
        print("Erro ao encontrar elemento 'action-button':")

    try:
        use_whatsapp_web = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[@class='_advp _aeam' and text()='use WhatsApp Web']")))
        print("Elemento 'use WhatsApp Web' encontrado")

        # Encontrando o link dentro do elemento 'use_whatsapp_web'
        link_element = use_whatsapp_web.find_element(By.XPATH, "./ancestor::a")
        link = link_element.get_attribute("href")

        # Abre o link na página atual
        driver.get(link)
    except Exception as e:
        print(f"Erro ao encontrar elemento 'use WhatsApp Web': {e}") """

    # Aguarda até 10 segundos para a presença do elemento com o seletor CSS "span[data-icon='send']" (ajuste conforme necessário)
    try:
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "span[data-icon='send']")))
        print("Elemento span[data-icon='send'] encontrado")
        element.click()

        time.sleep(2)
        driver.close()
        driver.switch_to.window(driver.window_handles[0])
    except Exception:
        print("Erro ao encontrar elemento span[data-icon='send']:")

for item in var:
    teste(item)
# Fecha navegador 
driver.quit()
