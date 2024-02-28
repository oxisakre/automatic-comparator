from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import pandas as pd
import urllib.parse
import re
import os
import difflib

chrome_options = Options()
chrome_options = webdriver.ChromeOptions()
service = ChromeService(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)
# Mapeo de nombres de Excel a nombres en la página web
mapeo_nombres = {
    'Fütterungshinweis': 'Fütterungshinweis',
    'Fütterungsempfehlung': 'Fütterungsempfehlung',
    'Zusammensetzung': 'Zusammensetzung',
    'Ernährungsphysiologische Zusatzstofffe je kg': 'Zusatzstoffe',
    'Technologische Zusatzstoffe je kg': 'Zusatzstoffe',  # Ajusta según la variación encontrada
    'Analytische Bestandteile und Gehalte': 'Analytische Bestandteile und Gehalte'
}

excepciones_url = {
    "OKAPI Kernige Cracker": "okapi-kernigecracker",
    "OKAPI Waldweide Kekse": "okapi-wald-weide-kekse",
    "OKAPI Blaubeer Kekse": "okapi-blaubeerkekse",
    "OKAPI Chia Clickerlis": "okapi-chiaclickerlis",
    "OKAPI Knusprige Clickerlis": "okapi-knusprigeclickerlis",
    "OKAPI Leichte Clickerlis": "okapi-leichteclickerlis",
    "OKAPI Cranberry Kekse": "okapi-cranberrykekse",
    "OKAPI Fix & Fertig Esparsette": "okapi-ffesparsette",
    
    # Agrega aquí más excepciones según sea necesario
}

def generar_url(nombre_producto):
    # Convertir a cadena y limpiar espacios
    nombre_producto = str(nombre_producto).strip()
    nombre_url = ""

    # Verificar si el producto está en el diccionario de excepciones
    if nombre_producto in excepciones_url:
        nombre_url = excepciones_url[nombre_producto]
    else:
        # Procesamiento estándar si no es una excepción
        nombre_url = nombre_producto.lower().replace(" ", "-").replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
        nombre_url = nombre_url.rstrip('-')
    
    nombre_url = urllib.parse.quote_plus(nombre_url, safe='-')
    url = f"https://www.okapi-online.de/{nombre_url}.html"
    return url
def extraer_descripciones(url, driver):
    try:
        driver.get(url)
        # Espera hasta que se cargue el contenido dinámico, ajusta los selectores y tiempos según sea necesario
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'h3.overline-header'))
        )
        descripciones = {}
        for h3 in driver.find_elements(By.CSS_SELECTOR, 'h3.overline-header'):
            key = h3.text.strip()
            value = h3.find_element(By.XPATH, 'following-sibling::p').text.strip() if h3.find_element(By.XPATH, 'following-sibling::p') else ''
            descripciones[key] = value
        return descripciones
    except Exception as e:
        print(f"Error al extraer de {url}: {e}")
    return {}

file_path = r"C:\Users\Super Leo\Downloads\Masterliste_Ubersetzungen_19.02.24.xlsx"
df = pd.read_excel(file_path)
productos_no_coincidentes = []
ultimo_producto = None 
columnas_a_verificar = [
    'Fütterungshinweis', 'Fütterungsempfehlung', 'Zusammensetzung',
    'Ernährungsphysiologische Zusatzstofffe je kg', 'Technologische Zusatzstoffe je kg',
    'Analytische Bestandteile und Gehalte', 
]

def normalizar_Analytische(texto):
    # Normalizar el formato de números y porcentajes
    texto = re.sub(r'(\d)[.,](\d)', r'\1.\2', texto)  # Unificar formato decimal
    texto = re.sub(r'(\d)\s*%', r'\1%', texto)  # Eliminar espacios antes de %
    
    # Eliminar puntuación irrelevante
    texto = re.sub(r'[,.]', ' ', texto)
    
    # Convertir múltiples espacios en uno solo
    texto = re.sub(r'\s+', ' ', texto).strip().lower()
    
    return texto
def normalizar_general(texto):
    # Aquí, personaliza la normalización para Fütterungshinweis
    texto = re.sub(r'\s+', ' ', texto).strip().lower()
    texto = re.sub(r'[,.]', ' ', texto)
    return texto
   
def encontrar_diferencias(texto_excel, texto_web):
    diferencias = list(difflib.ndiff([texto_excel], [texto_web]))
    solo_en_excel = ' '.join([dif[2:] for dif in diferencias if dif.startswith('- ')])
    solo_en_web = ' '.join([dif[2:] for dif in diferencias if dif.startswith('+ ')])

    diferencias_str = ""
    if solo_en_excel:
        diferencias_str += f"Excel -->\n{solo_en_excel}\n"
    if solo_en_web:
        diferencias_str += f"Web -->\n{solo_en_web}\n"

    return diferencias_str.strip()
    
# Iterar sobre el DataFrame
for index, row in df.iterrows():
    nombre_producto_actual = str(row['Artikelname Deutsch']).strip()

    # Verificar si el producto actual es el mismo que el último procesado
    if nombre_producto_actual == ultimo_producto:
        continue  # Saltear este producto

    # Actualizar el nombre del último producto procesado
    ultimo_producto = nombre_producto_actual

    # Aquí empiezas a trabajar con el producto actual, ya que es diferente al anterior
    nombre_producto = row['Artikelname Deutsch']
    url_producto = generar_url(nombre_producto)
    descripciones_web = extraer_descripciones(url_producto, driver)
    discrepancias_producto = []

    for columna in columnas_a_verificar:
        if columna in df.columns:
            valor_excel = str(row[columna]).strip() if pd.notnull(row[columna]) else None
            nombre_web_equivalente = mapeo_nombres.get(columna, columna)  # Usar mapeo o el mismo nombre si no hay mapeo
            valor_web = descripciones_web.get(nombre_web_equivalente, '').strip()
            
            if valor_excel is not None and valor_web is not None:
                # Decidir qué normalización aplicar
                if columna == 'Analytische Bestandteile und Gehalte' or columna == ('Ernährungsphysiologische Zusatzstofffe je kg' or 'Technologische Zusatzstoffe je kg'):
                    valor_excel_normalizado = normalizar_Analytische(valor_excel)
                    valor_web_normalizado = normalizar_Analytische(valor_web)
                else:
                    valor_excel_normalizado = normalizar_general(valor_excel)
                    valor_web_normalizado = normalizar_general(valor_web)
                
                # Comparar los textos normalizados
                if valor_excel_normalizado != valor_web_normalizado:
                    resumen_diferencias = encontrar_diferencias(valor_excel_normalizado, valor_web_normalizado)
                    discrepancias_producto.append((columna, valor_excel, valor_web, resumen_diferencias))

    if discrepancias_producto:
        productos_no_coincidentes.append((nombre_producto, discrepancias_producto))

# Preparar las discrepancias para escribir en el archivo
discrepancias_para_archivo = {}
for producto, discrepancias in productos_no_coincidentes:
    discrepancias_detalle = '\n'.join([
        f"{col}:\nExcel -->\n{exc}\nWeb -->\n{web}\nDiferencias:\n{dif}" 
        if dif else f"{col}:\nExcel -->\n{exc}\nWeb -->\n{web}" 
        for col, exc, web, dif in discrepancias
    ])
    discrepancias_para_archivo[producto] = discrepancias_detalle

def escribir_discrepancias_a_archivo(discrepancias, nombre_archivo="Anomalias.txt"):
    ruta_escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
    ruta_completa = os.path.join(ruta_escritorio, nombre_archivo)
    
    with open(ruta_completa, "w", encoding="utf-8") as archivo:
        for producto, detalle in discrepancias.items():
            archivo.write(f"Producto: {producto}\nDiscrepancias:\n{detalle}\n\n")

# Llama a la función para escribir las discrepancias
escribir_discrepancias_a_archivo(discrepancias_para_archivo)

print("Las discrepancias han sido escritas al archivo 'Anomalias.txt' en el escritorio.")
            
driver.quit()
