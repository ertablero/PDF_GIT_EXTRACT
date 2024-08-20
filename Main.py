from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# Configurar el WebDriver para Chrome
driver = webdriver.Chrome()

# Abrir la página web de Google
driver.get("https://miportal.dgrcorrientes.gov.ar/")

# Encontrar la barra de búsqueda por su nombre de atributo 'q'
search_box = driver.find_element(By.ID, "username")
# Escribir en la barra de búsqueda
search_box.send_keys("Selenium Python")
search_box = driver.find_element(By.ID, "loginPassword")
search_box.send_keys("Laclave*2652")

# Presionar la tecla Enter para iniciar la búsqueda
#search_box.send_keys(Keys.RETURN)
boton_ingresar = driver.find_element(By.ID, "ingresar")
boton_ingresar.click()

# Esperar unos segundos para ver los resultados
time.sleep(3)

# Cerrar el navegador
driver.quit()
