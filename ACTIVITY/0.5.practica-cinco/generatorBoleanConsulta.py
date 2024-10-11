import os
import pandas as pd
from collections import deque
from nltk.tokenize import word_tokenize
from nltk.stem.snowball import SnowballStemmer

# Inicializa el stemmer para español
stemmer = SnowballStemmer("spanish")

# Función para convertir la consulta a notación postfija
def infix_to_postfix(consulta):
    precedencia = {'not': 3, 'and': 2, 'or': 1}
    output = []
    operadores = []
    
    for token in consulta:
        if token not in precedencia and token != '(' and token != ')':
            output.append(token)
        elif token == '(':
            operadores.append(token)
        elif token == ')':
            while operadores and operadores[-1] != '(':
                output.append(operadores.pop())
            operadores.pop()  # Quitar '(' de la pila
        else:
            while (operadores and precedencia.get(token, 0) <= precedencia.get(operadores[-1], 0)):
                output.append(operadores.pop())
            operadores.append(token)

    while operadores:
        output.append(operadores.pop())
    
    return output

# Función para tokenizar y obtener la raíz de las palabras del contenido
def tokenize_and_stem(content):
    tokens = word_tokenize(content)
    stemmed_tokens = {stemmer.stem(token).upper() for token in tokens if token.isalpha() and len(token) > 1}
    return sorted(stemmed_tokens)

# Función para evaluar la consulta en notación postfija
def evaluate_postfix(postfix, hash_table):
    stack = deque()

    for token in postfix:
        # Aseguramos que los operadores no sean considerados palabras
        if token not in {'and', 'or', 'not'}:
            # Comparamos con las palabras ya procesadas (en mayúsculas)
            if token in hash_table:
                term_vector = hash_table[token]
                stack.append(deque(term_vector))
            else:
                print(f"Advertencia: El término '{token}' no se encuentra en la tabla hash.")
                stack.append(deque([0] * len(next(iter(hash_table.values())))))
        else:
            if token == 'not':
                a = stack.pop()
                print(f"Evaluando 'not': {list(a)}")  # Mostrar vector antes de aplicar 'not'
                stack.append(deque([1 - x for x in a]))
            else:
                b = stack.pop()
                a = stack.pop()
                if token == 'and':
                    print(f"Evaluando 'and': {list(a)} AND {list(b)}")  # Mostrar vectores antes de aplicar 'and'
                    stack.append(deque([x & y for x, y in zip(a, b)]))
                elif token == 'or':
                    print(f"Evaluando 'or': {list(a)} OR {list(b)}")  # Mostrar vectores antes de aplicar 'or'
                    stack.append(deque([x | y for x, y in zip(a, b)]))
    
    return list(stack.pop())

# Cargar el Excel
archivo_excel = 'PreprocesamientoSteps/6MatrizBinariaDeSteams.xlsx'
df = pd.read_excel(archivo_excel, index_col=0)  

hash_table = {col: list(df.loc[col].values) for col in df.index}

print("Matriz binaria de documentos (cargada desde Excel):")
print(df)

try:
    os.makedirs('challengeExample', exist_ok=True)
    ruta_matriz = 'challengeExample/matriz_binaria.xlsx'
    df.to_excel(ruta_matriz)
    print(f"Matriz binaria guardada en el archivo Excel: {ruta_matriz}")
except Exception as e:
    print(f"Ocurrió un error al guardar el archivo: {e}")

# Ingresar consulta
consulta_str = input("Ingrese la consulta en formato infijo con paréntesis: ")

# Tokenizar y obtener los steams de la consulta ingresada por el usuario
tokens = tokenize_and_stem(consulta_str)
print(f"Palabras tokenizadas y en su raíz: {tokens}")

# Convertir la consulta a formato infijo
# Separar los operadores lógicos de los términos
consulta = consulta_str.replace('(', ' ( ').replace(')', ' ) ').split()
# Aseguramos que las palabras se transformen a su raíz en la notación postfija
consulta_postfija = infix_to_postfix(consulta)
# Convertimos los tokens a su raíz también para la consulta en notación postfija
consulta_postfija = [stemmer.stem(token).upper() if token not in {'(', ')', 'and', 'or', 'not'} else token for token in consulta_postfija]
print(f"Consulta en notación postfija: {consulta_postfija}")

try:
    # Evaluar la consulta en notación postfija
    resultado = evaluate_postfix(consulta_postfija, hash_table)
    
    # Crear el DataFrame sin el primer elemento del resultado (resultado[1:]) y sin el índice
    resultado_matriz = pd.DataFrame([resultado[1:]], columns=df.columns[1:])  # Ignorar el primer elemento y ajustar las columnas

    # Mostrar el resultado en forma de matriz
    print("Resultado de la consulta:")
    print(resultado_matriz)

    # Guardar el resultado en Excel
    ruta_resultado = 'challengeExample/resultado_consulta.xlsx'
    resultado_matriz.to_excel(ruta_resultado, index=False)
    
    # Confirmar que el resultado se guardó
    print(f"Resultado guardado en el archivo Excel: {ruta_resultado}")

except ValueError as e:
    print("Error en la evaluación de la consulta:", e)