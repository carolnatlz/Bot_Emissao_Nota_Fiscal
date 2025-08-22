from selenium import webdriver
import time
from selenium.webdriver.common.by import By

# Inicia o driver do Chrome
driver = webdriver.Chrome()

# Abre uma página de teste
driver.get("https://www.google.com")
driver.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys('Funcionou!!!')

# Espera 15 segundos para que você veja o Chrome em ação
time.sleep(15)

# Fecha o navegador
driver.quit()
