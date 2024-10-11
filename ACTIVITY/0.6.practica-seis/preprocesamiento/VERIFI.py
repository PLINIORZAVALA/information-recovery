import pandas as pd

# Cargar el archivo Excel
hashs = 'PreprocesamientoSteps/5ListDiccIndex.xlsx'
df = pd.read_excel(hashs)

# Mostrar las columnas del DataFrame
print("Columnas del DataFrame:", df.columns)

# Asegurarse de que la columna 'Hash' se trata como string
df['Hash'] = df['Hash'].astype(str)

# Reemplazar ',' con '.' y convertir a float, manejando posibles errores
df['Hash'] = df['Hash'].str.replace(',', '.', regex=False)
df['Hash'] = pd.to_numeric(df['Hash'], errors='coerce')  # Convertir a float, estableciendo NaN para valores no convertibles

# Verificar si hay términos en la columna 'Steam'
terminos_a_buscar = ['TIT', 'CURIOS', 'OR']

for termino in terminos_a_buscar:
    # Verificar la existencia del término en la columna 'Steam', ignorando mayúsculas
    if df['Steam'].str.contains(termino, case=False, na=False).any():  
        print(f"El término '{termino}' existe en el archivo.")
    else:
        print(f"El término '{termino}' NO existe en el archivo.")
