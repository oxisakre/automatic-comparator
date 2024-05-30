import os
import urllib.parse
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import csv

def seleccionar_archivo_excel():
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal de Tk
    root.update()  # Procesar eventos pendientes y asegurarse de que se oculta la ventana
    # Mostrar el cuadro de diálogo para que el usuario elija el archivo
    file_path = filedialog.askopenfilename(
        title="Select the Excel file",
        filetypes=[("Excel files", "*.xlsx")]
    )
    root.destroy()  # Cerrar la ventana de Tkinter
    return file_path

def generar_url(nombre_producto):
    nombre_producto = str(nombre_producto).strip()
    nombre_url = ""
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
    }
    
    if nombre_producto in excepciones_url:
        nombre_url = excepciones_url[nombre_producto]
    else:
        nombre_url = nombre_producto.lower().replace(" ", "-").replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
        nombre_url = nombre_url.rstrip('-')
    
    nombre_url = urllib.parse.quote_plus(nombre_url, safe='-')
    url = f"https://www.okapi-online.de/{nombre_url}.html"
    return url

def generar_short_url(nombre_producto):
    nombre_producto = nombre_producto.lower().replace('okapi-', '').strip()
    nombre_producto2 = nombre_producto.split('-')[0] if '-' in nombre_producto else nombre_producto
    nombre_url = nombre_producto2.lower().replace(" ", "-").replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss").replace('okapi-', '')
    return nombre_url

def leer_productos_y_guardar_urls():
    file_path = seleccionar_archivo_excel()
    df = pd.read_excel(file_path)
    
    productos = df['Artikelname Deutsch'].tolist()  # Cambia 'Producto' por el nombre correcto de la columna en tu archivo Excel
    unique_urls = set()
    
    for producto in productos:
        if isinstance(producto, str):  # Asegúrate de procesar solo valores de cadena
            long_url = generar_url(producto)
            short_url = generar_short_url(producto)
            title = producto.lower().replace(" ", "-").replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
            unique_urls.add((long_url, short_url, title))
    
    with open('product_urls.csv', 'w', newline='', encoding='utf-8') as csvfile:
        csvwriter = csv.writer(csvfile, quoting=csv.QUOTE_MINIMAL)
        csvwriter.writerow(['Long URL', 'Short URL', 'Title'])
        csvwriter.writerows(unique_urls)
    
    print(f'{len(unique_urls)} URLs guardados en product_urls.csv')

if __name__ == '__main__':
    leer_productos_y_guardar_urls()
