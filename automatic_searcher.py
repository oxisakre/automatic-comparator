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
import pandas as pd

def seleccionar_archivo_excel():
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal de Tk
    root.update() # Procesar eventos pendientes y asegurarse de que se oculta la ventana
    # Mostrar el cuadro de diálogo para que el usuario elija el archivo
    file_path = filedialog.askopenfilename(
        title="Seleccione el archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx")]
    )

    root.destroy()  # Cerrar la ventana de Tkinter

    return file_path
def buscar_un_producto():
    nombre_del_producto = input("Ingrese el nombre del producto: ")
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
        print("Se agotó el tiempo de espera al buscar el producto:", e)
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
        'Ernährungsphysiologische Zusatzstofffe je kg': 'Zusatzstoffe',
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
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'h3.overline-header'))
            )
            descripciones = {}
            for h3 in driver.find_elements(By.CSS_SELECTOR, 'h3.overline-header'):
                key = h3.text.strip()
                # Busca el párrafo siguiente para cada encabezado h3
                p_element = h3.find_element(By.XPATH, 'following-sibling::p')
                value = p_element.text.strip() if p_element else ''

                # Si el encabezado es 'Zusatzstoffe', debes buscar ambos tipos dentro del texto
                if 'Zusatzstoffe' in key:
                    # Divide el valor por líneas nuevas y procesa cada línea por separado
                    for line in value.split('\n'):
                        if "Ernährungsphysiologische Zusatzstoffe je kg:" in line:
                            # Asigna el valor a la clave correspondiente sin el título
                            descripciones['Ernährungsphysiologische Zusatzstofffe je kg'] = line.split("Ernährungsphysiologische Zusatzstoffe je kg:")[1].strip()
                        elif "Technologische Zusatzstoffe je kg:" in line:
                            # Asigna el valor a la clave correspondiente sin el título
                            descripciones['Technologische Zusatzstoffe je kg'] = line.split("Technologische Zusatzstoffe je kg:")[1].strip()
                else:
                    descripciones[key] = value

            return descripciones
        except Exception as e:
            print(f"Error al extraer de {url}: {e}")
        return {}
    

    file_path = seleccionar_archivo_excel()
    if not file_path:
        print("No se seleccionó ningún archivo.")
        return
    df = pd.read_excel(file_path)
    productos_no_coincidentes = []
    ultimo_producto = None 
    columnas_a_verificar = [
        'Fütterungshinweis', 'Fütterungsempfehlung ', 'Zusammensetzung ',
        'Ernährungsphysiologische Zusatzstofffe je kg', 'Technologische Zusatzstoffe je kg',
        'Analytische Bestandteile und Gehalte', 
    ]

    
    def normalizar_Analytische(texto):
        # Eliminar saltos de línea y retornos de carro
        texto = texto.replace('\n', ' ').replace('\r', ' ')

        # Unificar formato decimal
        texto = re.sub(r'(\d)[.,](\d)', r'\1.\2', texto)

        # Unificar espacios alrededor de las unidades y porcentajes
        texto = re.sub(r'(\d)\s*%', r'\1%', texto)  # "15 %" a "15%"
        texto = re.sub(r'(\d)\s*(mg|g|kg|µg)\b', r'\1\2', texto)  # "100 mg" a "100mg"

        # Convertir a minúsculas y eliminar puntuación irrelevante
        texto = re.sub(r'[,.]', ' ', texto).lower()

        # Eliminar espacios extras
        texto = re.sub(r'\s+', ' ', texto).strip()

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
            return "No hay diferencias"
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
        if nombre_producto_actual == ultimo_producto:
            continue
        ultimo_producto = nombre_producto_actual

        nombre_producto = row['Artikelname Deutsch']
        url_producto = generar_url(nombre_producto)
        descripciones_web = extraer_descripciones(url_producto, driver)
        discrepancias_producto = []
        hay_diferencias = False

        for columna in columnas_a_verificar:
            if columna in df.columns:
                valor_excel = str(row[columna]).strip() if pd.notnull(row[columna]) else ""
                valor_web = descripciones_web.get(mapeo_nombres.get(columna, columna), '').strip()

                # Aplicar la normalización específica según la columna
                valor_excel_normalizado = normalizar_general(valor_excel)
                valor_web_normalizado = normalizar_general(valor_web)

                if columna in ['Analytische Bestandteile und Gehalte', 'Ernährungsphysiologische Zusatzstofffe je kg', 'Technologische Zusatzstoffe je kg', 'Ernährungsphysiologische Zusatzstoffe je kg']:
                    valor_excel_normalizado = normalizar_Analytische(valor_excel)
                    valor_web_normalizado = normalizar_Analytische(valor_web)

                resumen_diferencias = encontrar_diferencias(valor_excel_normalizado, valor_web_normalizado)
                if resumen_diferencias != "No hay diferencias":
                    hay_diferencias = True
                    discrepancias_producto.append((columna, resumen_diferencias))

        if not hay_diferencias:
            # Si no hay diferencias, agregamos una tupla con el nombre del producto y un mensaje general
            productos_no_coincidentes.append((nombre_producto, [("General", "No hay diferencias")]))
        else:
            # Si hay diferencias, agregamos una tupla con el nombre del producto y la lista de discrepancias
            productos_no_coincidentes.append((nombre_producto, discrepancias_producto))

    discrepancias_para_archivo = {}
    for producto, discrepancias in productos_no_coincidentes:
        discrepancias_detalle = []
        for dis in discrepancias:
            # Suponiendo que dis es una tupla de la forma (columna, resumen_diferencias)
            col, dif = dis
            if col == "General":
                discrepancias_detalle.append(dif)  # "No hay diferencias"
            else:
                # Ajusta aquí el formato para incluir un salto de línea antes de "En Excel:" y "En Web:"
                partes_dif = dif.split('\n')  # Esto asume que dif ya incluye "En Excel:" y "En Web:"
                dif_formateada = "\n".join(partes_dif)  # Re-construye la discrepancia con los saltos de línea
                discrepancias_detalle.append(f"{col}:\n{dif_formateada}")
        
        discrepancias_para_archivo[producto] = '\n'.join(discrepancias_detalle)

    def escribir_discrepancias_a_excel(productos_no_coincidentes, nombre_archivo="Anomalias.xlsx"):
    # Lista para almacenar los datos antes de convertirlos a DataFrame
        data = []
        for producto, discrepancias in productos_no_coincidentes:
            datos_producto = {'Producto': producto}
            for columna in columnas_a_verificar:
                datos_producto[columna] = "No hay diferencias"  # Inicializa todas las columnas con 'No hay diferencias'
            for columna, discrepancia in discrepancias:
                if columna == "General" and discrepancia == "No hay diferencias":
                    for col in columnas_a_verificar:
                        datos_producto[col] = "No hay diferencias"  # Establece todas las columnas a 'No hay diferencias'
                else:
                    datos_producto[columna] = discrepancia  # Actualiza la columna específica con su discrepancia
            data.append(datos_producto)

        # Convierte la lista de datos a un DataFrame
        df = pd.DataFrame(data)

        # Escribe el DataFrame a un archivo Excel
        ruta_completa = os.path.join(os.path.expanduser("~"), "Desktop", nombre_archivo)
        df.to_excel(ruta_completa, index=False)
        print(f"El archivo {nombre_archivo} ha sido guardado en el escritorio.")
            # Llama a la función para escribir las discrepancias
    escribir_discrepancias_a_excel(discrepancias_para_archivo)

    

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