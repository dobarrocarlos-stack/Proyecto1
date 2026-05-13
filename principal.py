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


    ultima_fila = nueva_hoja.range("R1048576").end("up").row

    nueva_hoja.range(f"FE2:FE{ultima_fila}").formula = '=R2&G2&AE2'

    # Guardar y cerrar
    closeExcel(wb, app)


def createReviewer():

    app = xw.App(visible=False)

    wb_plantilla = app.books.open(r"plantilla.xlsx")
    wb_reviewer = app.books.open(r"reviewer.xlsx")

    hoja_origen = wb_reviewer.sheets[0]

    nombre_hoja = "Reviewer"

    if nombre_hoja in [s.name for s in wb_plantilla.sheets]:
        wb_plantilla.sheets[nombre_hoja].delete()

    hoja_destino = wb_plantilla.sheets.add(nombre_hoja, after=wb_plantilla.sheets[-1])

    rango_usado = hoja_origen.used_range

    hoja_destino.range("A1").value = rango_usado.value

    rango_usado.api.Copy(hoja_destino.range("A1").api)

    wb_plantilla.save()

    wb_reviewer.close()
    wb_plantilla.close()

    app.quit()

def createTb():

    app = xw.App(visible=False)

    wb_plantilla = app.books.open(r"plantilla.xlsx")
    wb_reviewer = app.books.open(r"tb.xlsx")

    hoja_origen = wb_reviewer.sheets[0]

    nombre_hoja = "FDD Data J658"

    if nombre_hoja in [s.name for s in wb_plantilla.sheets]:
        wb_plantilla.sheets[nombre_hoja].delete()

    hoja_destino = wb_plantilla.sheets.add(nombre_hoja, after=wb_plantilla.sheets[-1])

    rango_usado = hoja_origen.used_range

    # Copiar datos
    hoja_destino.range("A1").value = rango_usado.value

    # Copiar formato
    rango_usado.api.Copy(hoja_destino.range("A1").api)

    # Última fila con datos en columna A
    ultima_fila = hoja_destino.range("A1048576").end("up").row

    # Concatenar A + C + B en columna K
    hoja_destino.range(f"K2:K{ultima_fila}").formula = "=A2&C2&B2"

    wb_plantilla.save()

    wb_reviewer.close()
    wb_plantilla.close()

    app.quit()


def menu():
    print("MENU")
    print("1) create BSC Data")
    print("2) create tb")
    print("3) crear Reviewer")

    option = int(input("Elige una opcion: "))

    return option


option = menu()

if option == 1:
    copyMonth()
elif option == 2:
    createTb()
elif option == 3:
    createReviewer()


# columna A hasta la FD, FE cancotenar r2,g2,ae22