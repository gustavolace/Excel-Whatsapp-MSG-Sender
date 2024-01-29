from selenium import webdriver
import os
import time

appdata_dir = os.getenv('APPDATA')
chrome_profile_dir = os.path.join(appdata_dir, 'Local', 'Google', 'Chrome', 'User Data', 'Profile 1')
print(chrome_profile_dir)

options = webdriver.ChromeOptions()
options.add_argument(f"--user-data-dir={chrome_profile_dir}")
user_agent = 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.50 Safari/537.36'
options.add_argument(f'user-agent={user_agent}')
driver = webdriver.Chrome(options=options)

driver.get("https://web.whatsapp.com/")
time.sleep(30)