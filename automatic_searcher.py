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
        "OKAPI Frühlingsknusper": "okapi-fruehlingsknusper",
        "OKAPI Fix & Fertig Esparsette": "okapi-ffesparsette",
        'OKAPI Junior Mineral G': 'okapi-junior-mineral',
        "OKAPI Heucobs sugar light Futterprobe": "okapi-heucobs-sugar-light",
        "OKAPI Fix & Fertig Luzerne": "okapi-luzerne-fix-fertig",
        "OKAPI Kieselgur Plus": "kieselgur",
        "OKAPI Mineral Pur G Leckschale": "okapi-mineral-pur-leckschale-7-000g",
        "OKAPI Leckschale Junior Mineral": "okapi-junior-mineral-leckschale-7-000g",
        "OKAPI Laub & Rinden": "okapi-laub-und-rinden",
        "OKAPI Pränat Plus Typ Z & K": "pranat-plus-typ-z-k",
        "OKAPI Mineralbricks": "okapi-mineral-bricks",
        "OKAPI Mineralkonzentrat G (S)": "okapi-mineralkonzentrat-g-s",
        "OKAPI Relax": "relax",
        "OKAPI Weidemineral Leckschale": "okapi-weidemineral-leckschale-7-000g",
        "OKAPI Weidemineral G (S)": "okapi-weidemineral-g-s",
        "Teepferdchen Gentle Detox": "gentle-detox",
        "Teepferdchen Happy Belly": "happy-belly",
        "Teepferdchen Relax Me": "relax-me",
        "OKAPI Weidemineral Leckschale": "okapi-weidemineral-leckschale-7-000g",
        "KNÄX Grüne Gemüse": "knaex-gruenes-gemuese",
        "KNÄX Hirschornmehl Pur": "knaex-hirschhornmehl-pur",
        "KNäX Fish 'n' Chips Snacks": "knaex-fish-n-chips-snacks",
        "KNäX Käse & Ei Snacks": "knaex-kaese-ei-snacks",
        "KNÄX GESUND DURCHS JAHR – JANUAR: Aurora": "knaex-aurora",
        "KNÄX GESUND DURCHS JAHR – FEBRUAR: Schneeschmelze": "knaex-Schneeschmelze",
        "KNÄX GESUND DURCHS JAHR – MÄRZ: Apport": "knaex-Apport",
        "KNÄX GESUND DURCHS JAHR – APRIL: Alles im Fluss": "knaex-alles-im-fluss",
        "KNÄX GESUND DURCHS JAHR – MAI: Kleine Riesen": "knaex-kleine-riesen",
        "KNÄX GESUND DURCHS JAHR – JUNI: Heublumen": "knaex-heublumen",
        "KNÄX GESUND DURCHS JAHR – JULI: Schäferfreund": "knaex-Schaeferfreund",
        "KNÄX GESUND DURCHS JAHR – AUGUST: Soleil": "knaex-soleil",
        "KNÄX GESUND DURCHS JAHR – SEPTEMBER: Altweibersommer": "knaex-altweibersommer",
        "KNÄX GESUND DURCHS JAHR – OKTOBER: Starke Abwehr": "knaex-starke-abwehr",
        "KNÄX GESUND DURCHS JAHR – NOVEMBER: Pfützenspaß": "knaex-pfuetzenspass",
        "KNÄX GESUND DURCHS JAHR – DEZEMBER: Kraftquelle": "knaex-kraftquelle",
        "Biostickies Aronia Standard": "biostickies-aronia",
        "Biostickies Fenchel Standard": "biostickies-fenchel",
        "Biostickies Hagebutte Standard": "biostickies-hagebutte",
        "Biostickies Hanf Standard": "biostickies-hanf",
        "Biostickies Mariendistel Standard": "biostickies-mariendistel",
        "Biostickies Natur Pur Standard": "biostickies-natur-pur",
        "Biostickies Ringelblume Standard": "biostickies-ringelblume",
        "Biostickies Schwarzkuemmel Standard": "biostickies-schwarzkummel",
        "Biostickies Suessholz Standard": "biostickies-sussholz",
        "Biostickies Thymian Standard": "biostickies-thymian",
        # Agrega aquí más excepciones según sea necesario
    }
    sinpagina_url = { 
        'OKAPI Entschuldigungspäckchen Wiesenkekse','OKAPI Esparsette Futterprobe', 'OKAPI Fix & Fertig Esparsette Futterprobe', 'OKAPI Fix & Fertig Luzerne Futterprobe',
        'OKAPI Heucobs sugar light', 'OKAPI Pränat Plus Typ K', 'OKAPI Pränat Plus Typ Z', 'OKAPI Ration Balancer', 'OKAPI Vierjahreszeitenfutter Fellwechsel Futterprobe',
        'OKAPI Vierjahreszeitenfutter Frühlingsgefühle Futterprobe', 'OKAPI Vierjahreszeitenfutter Herbsttage Futterprobe', 'OKAPI Vierjahreszeitenfutter Sommerkräuter Futterprobe',
        'OKAPI Vierjahreszeitenfutter Weidestart Futterprobe', 'OKAPI Vierjahreszeitenfutter Winterweide Futterprobe', 'OKAPI Vitalcobs Futterprobe', 'OKAPI Weihnachtskekse','Biostickies Aronia Clickerli', 'Biostickies Fenchel Clickerli',
        'Biostickies Hagebutte Clickerli','Biostickies Hanf Clickerli','Biostickies Mariendistel Clickerli','Biostickies Natur Pur Clickerli','Biostickies Ringelblume Clickerli','Biostickies Schwarzkuemmel Clickerli',
        'Biostickies Suessholz Clickerli','Biostickies Thymian Clickerli','kekse big size','middle size kekse'
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
    def extraer_descripciones_generales(url, driver):
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
    def extraer_descripciones_biostickies(url, driver):
        try:
            from bs4 import BeautifulSoup

            driver.get(url)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "h4.overline-header"))
            )
            descripciones = {}
            html_content = driver.page_source
            soup = BeautifulSoup(html_content, 'html.parser')
            headers = soup.select("h4.overline-header")

            for header in headers:
                key = header.get_text(strip=True)
                value = ''
                # Recorrer los nodos siguientes hasta encontrar el próximo encabezado o llegar al final del contenedor.
                for sibling in header.next_siblings:
                    if sibling.name == 'h4':
                        # Si encontramos otro encabezado, detenemos la búsqueda.
                        break
                    if sibling.name == 'br':
                        # Si encontramos un <br>, continuamos, ya que el texto relevante podría estar después.
                        continue
                    if sibling.name is None:
                        # Esto significa que es un nodo de texto y no una etiqueta.
                        value += sibling.strip()
                descripciones[key] = value

            return descripciones
        except Exception as e:
            print(f"Error extracting from {url}: {e}")
        return {}
    def extraer_descripciones_excepciones(url, driver):
        try:
            from bs4 import BeautifulSoup

            driver.get(url)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "h3.overline-header"))
            )
            descripciones = {}
            html_content = driver.page_source
            soup = BeautifulSoup(html_content, 'html.parser')
            headers = soup.select("h3.overline-header")

            for header in headers:
                key = header.get_text(strip=True)
                value = ''
                # Recorrer los nodos siguientes hasta encontrar el próximo encabezado o llegar al final del contenedor.
                for sibling in header.next_siblings:
                    
                    if sibling.name == 'h3':
                        # Si encontramos otro encabezado, detenemos la búsqueda.
                        break
                    if sibling.name != 'br':
                        # Si encontramos un <br>, continuamos, ya que el texto relevante podría estar después.
                        value = ' '.join([x for x in sibling.stripped_strings])
                        if value != '':
                            break
                        
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
                resultado += f"In Excel: {diferencias_texto_excel}\n"
            if diferencias_texto_web:
                resultado += f"In Web: {diferencias_texto_web}\n"
            return resultado.strip()

    # Iterar sobre el DataFrame
    productos_no_coincidentes = []
    ultimo_producto = None 
    for index, row in df.iterrows():
        nombre_producto_actual = str(row['Artikelname Deutsch']).strip()

        # Verificar si el producto actual es el mismo que el último procesado
        if nombre_producto_actual == ultimo_producto:
            continue  # Saltear este producto
        elif nombre_producto_actual in sinpagina_url:
            continue
        elif nombre_producto_actual == 'nan':
            continue

        ultimo_producto = nombre_producto_actual

        # Aquí empiezas a trabajar con el producto actual, ya que es diferente al anterior
        nombre_producto = row['Artikelname Deutsch']
        url_producto = generar_url(nombre_producto)
        excepciones = ['lapachorinde', 'prodarm','frühlingsknusper', 'tuttifrutti','winterknusper',]
        if any(exc in nombre_producto_actual.lower() for exc in excepciones):
            # Llama a la función especializada para excepciones
            descripciones_web = extraer_descripciones_excepciones(url_producto, driver)
        elif 'biostickies' in nombre_producto_actual.lower():
            # Llama a la función especializada para 'biostickies'
            descripciones_web = extraer_descripciones_biostickies(url_producto, driver)
        else:
            # Llama a la función de extracción original para otros productos
            descripciones_web = extraer_descripciones_generales(url_producto, driver)
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
   

    def escribir_discrepancias_a_archivo(discrepancias, nombre_archivo="Anomalies.txt"):
        ruta_escritorio = os.path.join(os.path.expanduser("~"), "Desktop")
        ruta_completa = os.path.join(ruta_escritorio, nombre_archivo)
        with open(ruta_completa, "w", encoding="utf-8") as archivo:
            for producto, detalle in discrepancias.items():
                archivo.write(f"Producto: {producto}\n{detalle}\n\n")
    # Llama a la función para escribir las discrepancias
    escribir_discrepancias_a_archivo(discrepancias_para_archivo)
    print("The discrepancies have been written to the 'Anomalies.txt' file on the desktop.")
    input("Press any key to continue...")   
def main():
    print('The program is running... Please wait')
    leer_todos_los_productos()
        
main()