from openpyxl import Workbook

# Crea un nuevo libro de trabajo y selecciona la hoja activa
workbook = Workbook()
sheet = workbook.active

# Agrega algunos datos a la primera celda
sheet['A1'] = 'Hola, OpenPyXL!'

# Guarda el libro de trabajo en un archivo
workbook.save('test_openpyxl.xlsx')