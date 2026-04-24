import xlwings as xw 
import os


print("Directorio actual:", os.getcwd())
print("Archivos en el directorio:", os.listdir())


# Abrir Excel 
app = xw.App(visible=False) 
wb = app.books.open(r"plantilla.xlsx") 

# Seleccionar hoja origen 
hoja_origen = wb.sheets["BSC Data 02.2026"] 

# Copiar hoja 
hoja_origen.api.Copy(After=wb.sheets[-1].api) 
# Renombrar la nueva hoja 
nueva_hoja = wb.sheets[-1] 
nueva_hoja.name = "BSC Data 03.2026"

# Guardar y cerrar 
wb.save() 
wb.close() 
app.quit()

