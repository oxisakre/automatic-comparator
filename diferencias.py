import difflib

def encontrar_diferencias(texto_excel, texto_web):
    # Separar los textos en palabras
    texto_excel_palabras = texto_excel.split()
    texto_web_palabras = texto_web.split()
    
    # Usar difflib para encontrar las diferencias
    d = difflib.Differ()
    diff = list(d.compare(texto_excel_palabras, texto_web_palabras))
    
    # Filtrar las diferencias: las líneas que comienzan con - o + son las que tienen diferencias
    diferencias = [line[2:] for line in diff if line.startswith('- ') or line.startswith('+ ')]
    
    # Unir las diferencias en una cadena de texto para devolverlas
    diferencias_texto = ' '.join(diferencias)
    
    return diferencias_texto

# Ejemplo de uso:
texto_excel = "Este es un ejemplo de texto de Excel con algunas palabras riggea"
texto_web = "Este es un ejemplo de texto de web con algunas palabras diferentes caveradea"

diferencias = encontrar_diferencias(texto_excel, texto_web)
print("Diferencias encontradas:", diferencias)