import pandas as pd

# codigo para leer el fichero excel
df = pd.read_excel("1.xlsx")

# Nueva fila (como diccionario)
nueva_fila = {"Invoice": 444, "Reference": 555, "Amount": 666}

# Añadirla
df = pd.concat([df, pd.DataFrame([nueva_fila])], ignore_index=True)

# Guardar cambios
df.to_excel("1.xlsx", index=False)

print(df.head())