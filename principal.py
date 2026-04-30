import xlwings as xw 
from datetime import datetime


def lastMonth():
    hoy = datetime.now()
    if hoy.month == 1:
        return f"12.{hoy.year - 1}"
    else:
        return f"{hoy.month - 1:02d}.{hoy.year}"


def currentMonth():
    return datetime.now().strftime("%m.%Y")


def closeExcel(wb, app):
    wb.save() 
    wb.close() 
    app.quit()

def copyMonth():
    # Abrir Excel 
    app = xw.App(visible=False) 
    wb = app.books.open(r"plantilla.xlsx") 

    # Nombres
    nombre_origen = f"BSC Data {lastMonth()}"
    nombre_destino = f"BSC Data {currentMonth()}"

    nombres_hojas = [s.name for s in wb.sheets]

    # 🔴 Validar hoja origen
    if nombre_origen not in nombres_hojas:
        closeExcel(wb, app)
        raise ValueError(f"La hoja origen {nombre_origen} no existe")

    # 🔴 Validar hoja destino (NO debe existir)
    if nombre_destino in nombres_hojas:
        closeExcel(wb, app)
        raise ValueError(f"La hoja destino {nombre_destino} ya existe")

    # ✅ Copiar
    hoja_origen = wb.sheets[nombre_origen]
    hoja_origen.api.Copy(After=wb.sheets[-1].api)

    # ✅ Renombrar nueva hoja
    nueva_hoja = wb.sheets[-1]
    nueva_hoja.name = nombre_destino

    # Guardar y cerrar
    closeExcel(wb, app)


def menu():
    print("MENU")
    print("1) create BSC Data")

    option = int(input("Elige una opcion: "))

    return option


option = menu()

if option == 1:
    copyMonth()