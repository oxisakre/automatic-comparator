from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Configuración de Selenium para usar Chrome
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Abrir la URL del producto
driver.get("https://www.okapi-online.de/okapi-endoprotect.html")

# Esperar a que el contenido dinámico se cargue
time.sleep(5)  # Ajusta este tiempo según sea necesario

# Extraer la descripción (ajusta el selector según tu caso)
descripcion_element = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.CSS_SELECTOR, "h3.overline-header"))
)
descripcion = descripcion_element.text

print(descripcion)

# Cerrar el navegador
driver.quit()