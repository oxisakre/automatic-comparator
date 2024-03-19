from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import pandas as pd
import urllib.parse
import re
import os
import difflib
import tkinter as tk
from tkinter import filedialog

def seleccionar_archivo_excel():
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal de Tk
    root.update() # Procesar eventos pendientes y asegurarse de que se oculta la ventana
    # Mostrar el cuadro de diálogo para que el usuario elija el archivo
    file_path = filedialog.askopenfilename(
        title="Select the Excel file",
        filetypes=[("Excel files", "*.xlsx")]
    )
    root.destroy()  # Cerrar la ventana de Tkinter
    return file_path
def buscar_un_producto():
    nombre_del_producto = input("Enter the product name: ")
    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    
    # Formatear el nombre del producto para crear la URL
    nombre_producto_formateado = nombre_del_producto.lower().replace(' ', '-')
    url_producto = f"https://www.okapi-online.de/{nombre_producto_formateado}.html"
    # Abrir la URL del producto
    driver.get(url_producto)
    # Esperar a que el contenido dinámico se cargue y extraer la información
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.product.attribute.description"))
        )
        # Encuentra todos los elementos h3 dentro del div con clase 'product attribute description'
        titulos = driver.find_elements(By.CSS_SELECTOR, "div.product.attribute.description h3.overline-header")
        # Crear un diccionario para almacenar la información
        informacion_producto = {}
        # Iterar sobre los elementos h3 y obtener la información
        for titulo in titulos:
            key = titulo.text.strip()
            # Intentar encontrar el siguiente elemento p después del título
            value_element = titulo.find_element(By.XPATH, "./following-sibling::p[1]")
            value = value_element.text.strip() if value_element else ''
            informacion_producto[key] = value
        # Imprimir toda la información
        for key, value in informacion_producto.items():
            print(f"{key}: {value}")
    except TimeoutException as e:
        print("You timed out when searching for the product:", e)
    finally:
        driver.quit()
    return informacion_producto
def leer_todos_los_productos():
    
    chrome_options = Options()
    chrome_options = webdriver.ChromeOptions()
    service = ChromeService(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    # Mapeo de nombres de Excel a nombres en la página web
    mapeo_nombres = {
        'Fütterungshinweis': 'Fütterungshinweis',
        'Fütterungsempfehlung ': 'Fütterungsempfehlung',
        'Zusammensetzung ': 'Zusammensetzung',
        'Ernährungsphysiologische Zusatzstoffe je kg': 'Zusatzstoffe',
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
            print(f"Error extracting from {url}: {e}")
        return {}
    
    file_path = seleccionar_archivo_excel()
    if not file_path:
        print("No file was selected.")
        return
    df = pd.read_excel(file_path)
    productos_no_coincidentes = []
    ultimo_producto = None 
    columnas_a_verificar = [
        'Fütterungshinweis', 'Fütterungsempfehlung ', 'Zusammensetzung ',
        'Ernährungsphysiologische Zusatzstoffe je kg', 'Technologische Zusatzstoffe je kg',
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
        # Convertir a minúsculas y eliminar espacios extra al principio y al final
        texto = texto.strip().lower()
        
        # Unificar formato de números con unidades o porcentajes (eliminando espacio antes de %, g, mg, etc.)
        texto = re.sub(r'(\d)\s*%', r'\1%', texto)  # De "15 %" a "15%"
        texto = re.sub(r'(\d)\s*(g|mg|kg|µg)', r'\1\2', texto)  # De "15 g" a "15g"
        
        # Eliminar puntuación irrelevante y convertir comas a puntos para decimales
        texto = re.sub(r'[,.]', ' ', texto)
        texto = re.sub(r'(\d)\s*\.\s*(\d)', r'\1.\2', texto)  # Corrige espacios alrededor de puntos decimales si existen
        
        # Convertir múltiples espacios a uno solo
        texto = re.sub(r'\s+', ' ', texto)
        
        return texto
    
    
    def encontrar_diferencias(texto_excel, texto_web):
        texto_excel_palabras = texto_excel.split()
        texto_web_palabras = texto_web.split()
        
        d = difflib.Differ()
        diferencias = list(d.compare(texto_excel_palabras, texto_web_palabras))
        
        diferencias_filtradas_excel = []
        diferencias_filtradas_web = []
        
        for line in diferencias:
            if line.startswith('- '):
                diferencias_filtradas_excel.append(line[2:])
            elif line.startswith('+ '):
                diferencias_filtradas_web.append(line[2:])
        
        if not diferencias_filtradas_excel and not diferencias_filtradas_web:
            return "There are no differences"
        else:
            diferencias_texto_excel = ' '.join(diferencias_filtradas_excel)
            diferencias_texto_web = ' '.join(diferencias_filtradas_web)
            
            resultado = ""
            if diferencias_texto_excel:
                resultado += f"En Excel: {diferencias_texto_excel}\n"
            if diferencias_texto_web:
                resultado += f"En Web: {diferencias_texto_web}\n"
            return resultado.strip()

    # Iterar sobre el DataFrame
    productos_no_coincidentes = []
    ultimo_producto = None 
    for index, row in df.iterrows():
        nombre_producto_actual = str(row['Artikelname Deutsch']).strip()

        # Verificar si el producto actual es el mismo que el último procesado
        if nombre_producto_actual == ultimo_producto:
            continue  # Saltear este producto

        # Actualizar el nombre del último producto procesado
            continue
        ultimo_producto = nombre_producto_actual

        # Aquí empiezas a trabajar con el producto actual, ya que es diferente al anterior
        nombre_producto = row['Artikelname Deutsch']
        url_producto = generar_url(nombre_producto)
        descripciones_web = extraer_descripciones(url_producto, driver)
        discrepancias_producto = []
        hay_diferencias = False

        for columna in columnas_a_verificar:
            if columna in df.columns:
                valor_excel = str(row[columna]).strip() if pd.notnull(row[columna]) else None
                nombre_web_equivalente = mapeo_nombres.get(columna, columna)  # Usar mapeo o el mismo nombre si no hay mapeo
                nombre_web_equivalente = mapeo_nombres.get(columna, columna)
                valor_web = descripciones_web.get(nombre_web_equivalente, '').strip()

                if valor_excel is not None and valor_web is not None:
                    # Decidir qué normalización aplicar
                    if columna == 'Analytische Bestandteile und Gehalte' or columna == ('Ernährungsphysiologische Zusatzstoffe je kg' or 'Technologische Zusatzstoffe je kg'):
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
                if columna in ['Analytische Bestandteile und Gehalte', 'Ernährungsphysiologische Zusatzstoffe je kg', 'Technologische Zusatzstoffe je kg']:
                    valor_excel_normalizado = normalizar_Analytische(valor_excel) if valor_excel else ""
                    valor_web_normalizado = normalizar_Analytische(valor_web) if valor_web else ""
                else:
                    valor_excel_normalizado = normalizar_general(valor_excel) if valor_excel else ""
                    valor_web_normalizado = normalizar_general(valor_web) if valor_web else ""

                resumen_diferencias = encontrar_diferencias(valor_excel_normalizado, valor_web_normalizado)
                if resumen_diferencias != "There are no differences":
                    hay_diferencias = True
                    discrepancias_producto.append((columna, resumen_diferencias))

        if not hay_diferencias:
            productos_no_coincidentes.append((nombre_producto, [("General", "There are no differences")]))
        else:
            productos_no_coincidentes.append((nombre_producto, discrepancias_producto))

    # Preparar las discrepancias para escribir en el archivo
    discrepancias_para_archivo = {}
    for producto, discrepancias in productos_no_coincidentes:
        # Omitir productos con nombre no válido (NaN o vacío)
        if pd.isna(producto) or producto.strip() == "":
            continue

        discrepancias_detalle = []
        detalles_procesados = set()

        for dis in discrepancias:
            if len(dis) == 2:
                col, dif = dis
            elif len(dis) == 4:
                col, exc, web, dif = dis
            else:
                print("Error: Tuple of unexpected length.", dis)
                continue

            clave_unica = f"{col}:{dif.lower().replace(' ', '')}"
            if clave_unica not in detalles_procesados:
                detalles_procesados.add(clave_unica)
                if col == "General":
                    discrepancias_detalle.append(dif)
                else:
                    partes_dif = dif.split('\n')
                    dif_formateada = "\n".join(partes_dif)
                    discrepancias_detalle.append(f"{col}:\n{dif_formateada}")

        discrepancias_para_archivo[producto] = '\n'.join(discrepancias_detalle)
   

    def escribir_discrepancias_a_archivo(discrepancias, nombre_archivo="Anomalias.txt"):
        ruta_escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
        ruta_completa = os.path.join(ruta_escritorio, nombre_archivo)
        with open(ruta_completa, "w", encoding="utf-8") as archivo:
            for producto, detalle in discrepancias.items():
                archivo.write(f"Producto: {producto}\n{detalle}\n\n")
    # Llama a la función para escribir las discrepancias
    escribir_discrepancias_a_archivo(discrepancias_para_archivo)
    print("Las discrepancias han sido escritas al archivo 'Anomalias.txt' en el escritorio.")   
def main():
    
    while True:
        print("Elige una opción:")
        print("1. Ver la descripción de un producto")
        print("2. Leer todos los productos")
        print("3. Salir")
        opcion = input("Introduce el número de la opción deseada: ")
        if opcion == '1':
            buscar_un_producto()
        elif opcion == '2':
            print('Cargando... Estaria bueno que esperes un poquito')
            leer_todos_los_productos()
        elif opcion == '3':
            break
        else:
            print("Opción no válida, por favor intenta de nuevo.")
if __name__ == "__main__":
    main()