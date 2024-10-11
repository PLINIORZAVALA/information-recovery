import os
import unicodedata
import pandas as pd
from collections import deque
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
from nltk.stem import SnowballStemmer
from collections import defaultdict
from collections import Counter
from nltk.stem import PorterStemmer, WordNetLemmatizer
import docx
import PyPDF2
import pptx
from openpyxl import load_workbook
from nltk import download
import re 
import math


# Descargar recursos de NLTK
download('punkt')
download('stopwords')
download('wordnet')

# Inicializa el stemmer para español
stemmer = SnowballStemmer("spanish")

# Funciones para leer archivos
def read_word(filepath):
    doc = docx.Document(filepath)
    return "\n".join([para.text for para in doc.paragraphs])

def read_pdf(filepath):
    with open(filepath, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
        return text

def read_powerpoint(filepath):
    ppt = pptx.Presentation(filepath)
    text = ""
    for slide in ppt.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + "\n"
    return text

def read_excel(filepath):
    workbook = load_workbook(filepath)
    text = ""
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows(values_only=True):
            text += " ".join([str(cell) for cell in row if cell]) + "\n"
    return text

def create_folder():
    output_folder = 'PreprocesamientoSteps'
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    return output_folder

# Leer archivo con los steams
def read_steams_from_excel(file_path):
    df = pd.read_excel(file_path)
    steams = df['Termino'].astype(str).tolist()  # Aseguramos que todo sea cadena
    return steams

# Función hash personalizada para evitar colisiones
def custom_hash(steam, table_size):
    return hash(steam) % table_size

# Función para indexar (hashing) los steams con manejo de colisiones usando sondeo lineal
def index_steams():
    # Leer los steams del archivo generado en el proceso anterior (4DiccDeSteams.xlsx)
    file_path = "PreprocesamientoSteps/4DiccDeSteams.xlsx"
    if not os.path.exists(file_path):
        print(f"El archivo {file_path} no existe.")
        return

    steams = read_steams_from_excel(file_path)

    # Definir el tamaño de la tabla hash (puede ajustarse según la cantidad de datos)
    table_size = len(steams)

    # Crear una tabla hash vacía (None significa que la posición está libre)
    steam_hash_table = [None] * table_size

    # Función para insertar steams usando sondeo lineal
    def insert_with_linear_probing(steam):
        # Calcular el hash inicial
        index = custom_hash(steam, table_size)
        
        # Sondeo lineal: buscar el siguiente espacio disponible
        original_index = index
        while steam_hash_table[index] is not None:
            index = (index + 1) % table_size  # Avanzar al siguiente índice (circular)
            if index == original_index:  # Si volvemos al punto de inicio, la tabla está llena
                raise Exception("Tabla hash llena, no se puede insertar más elementos.")
        
        # Insertar el steam en el índice disponible
        steam_hash_table[index] = steam
        return index

    # Insertar cada steam en la tabla hash usando sondeo lineal
    steam_to_hash_mapping = []
    for steam in steams:
        hash_index = insert_with_linear_probing(steam)
        steam_to_hash_mapping.append((steam, hash_index))

    # Guardar la tabla en un archivo Excel
    output_folder = create_folder()
    df = pd.DataFrame(steam_to_hash_mapping, columns=['Steam', 'Hash'])
    output_file = os.path.join(output_folder, '5ListDiccIndex.xlsx')
    df.to_excel(output_file, index=False)
    print(f"Tabla hash guardada en {output_file}.")
    
    return steam_hash_table  # Devolvemos la tabla hash para su uso posterior

# Función para procesar un solo documento (sin mayúsculas ni stemming aún)
def process_document_basic(content, stop_words_upper):
    tokens = word_tokenize(content)
    tokens = [word for word in tokens if word.isalnum() and not word.isnumeric()]
    tokens = [unicodedata.normalize('NFKD', word).encode('ascii', 'ignore').decode('utf-8') for word in tokens]
    
    # Eliminar stopwords (pero sin mayúsculas aún)
    tokens = [token for token in tokens if token.upper() not in stop_words_upper]
    
    return tokens

def count_documents(corpus_dir):
    document_count = 0

    # Recorrer todos los archivos en el directorio
    for filename in os.listdir(corpus_dir):
        file_path = os.path.join(corpus_dir, filename)
        
        # Verificar si el archivo tiene una extensión que corresponde a documentos soportados
        if filename.endswith(('.txt', '.pdf', '.docx', '.xlsx', '.pptx')):
            document_count += 1

    return document_count

def tokenize_and_process_documents(corpus_dir):
    initial_dictionary = {}
    stemmer = SnowballStemmer("spanish")
    stop_words = set(word.upper() for word in stopwords.words('spanish'))  # Stopwords en mayúsculas
    output_folder = create_folder()  # Crear la carpeta para los archivos de salida

    document_word_counts = {}

    # Proceso 1: Crear el diccionario inicial
    for filename in os.listdir(corpus_dir):
        file_path = os.path.join(corpus_dir, filename)
        content = ""

        if filename.endswith(".txt"):
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
        elif filename.endswith(".pdf"):
            content = read_pdf(file_path)
        elif filename.endswith(".docx"):
            content = read_word(file_path)
        elif filename.endswith(".xlsx"):
            content = read_excel(file_path)
        elif filename.endswith(".pptx"):
            content = read_powerpoint(file_path)
        else:
            print(f"Ignorando archivo no soportado: {filename}")
            continue

        # Aplicar el preprocesamiento básico (sin mayúsculas ni stemming)
        processed_tokens = process_document_basic(content, stop_words)
        tokens = word_tokenize(content)
        tokens = [word for word in tokens if word.isalnum() and not word.isnumeric()]
        tokens = [unicodedata.normalize('NFKD', word).encode('ascii', 'ignore').decode('utf-8') for word in tokens]

        # Contar la frecuencia de palabras preprocesadas en cada documento
        doc_word_counts = {}
        for token in tokens:
            token_stemmed = stemmer.stem(token).upper()  # Aplicar stemming y convertir a mayúsculas
            if token_stemmed not in doc_word_counts:
                doc_word_counts[token_stemmed] = 1
            else:
                doc_word_counts[token_stemmed] += 1

        document_word_counts[filename] = doc_word_counts

        # Actualizar el diccionario inicial con el conteo de todas las palabras
        for token in processed_tokens:
            if token not in initial_dictionary:
                initial_dictionary[token] = 1
            else:
                initial_dictionary[token] += 1

    # Guardar el diccionario inicial en "1Diccionario.xlsx" (sin convertir a mayúsculas ni aplicar stemming)
    df_initial = pd.DataFrame(list(initial_dictionary.items()), columns=['Termino', 'Frecuencia'])
    df_initial.to_excel(os.path.join(output_folder, '1Diccionario.xlsx'), index=False)
    print('Diccionario inicial guardado en "1Diccionario.xlsx".')

    # Proceso 2: Convertir el diccionario inicial a mayúsculas
    upper_dictionary = {}
    for token in initial_dictionary:
        token_upper = token.upper()
        if token_upper not in upper_dictionary:
            upper_dictionary[token_upper] = initial_dictionary[token]
        else:
            upper_dictionary[token_upper] += initial_dictionary[token]

    # Guardar el diccionario con mayúsculas en "2DiccMayus.xlsx"
    df_upper = pd.DataFrame(list(upper_dictionary.items()), columns=['Termino', 'Frecuencia'])
    df_upper.to_excel(os.path.join(output_folder, '2DiccMayus.xlsx'), index=False)
    print('Diccionario con mayúsculas guardado en "2DiccMayus.xlsx".')

    # Proceso 3: Eliminar stopwords del diccionario con mayúsculas
    stop_words_upper = {word.upper() for word in stop_words}
    filtered_dictionary = {word: count for word, count in upper_dictionary.items() if word not in stop_words_upper}

    # Guardar el diccionario sin stopwords en "3DiccMSinStopWords.xlsx"
    df_filtered = pd.DataFrame(list(filtered_dictionary.items()), columns=['Termino', 'Frecuencia'])
    df_filtered.to_excel(os.path.join(output_folder, '3DiccMSinStopWords.xlsx'), index=False)
    print('Diccionario sin stop words guardado en "3DiccMSinStopWords.xlsx".')

    # Proceso 4: Aplicar stemming al diccionario
    stemmed_dictionary = {}
    for word, count in filtered_dictionary.items():
        stemmed_word = stemmer.stem(word).upper()
        if stemmed_word not in stemmed_dictionary:
            stemmed_dictionary[stemmed_word] = count
        else:
            stemmed_dictionary[stemmed_word] += count

    # Guardar el diccionario con stemming en "4DiccDeSteams.xlsx"
    df_stemmed = pd.DataFrame(list(stemmed_dictionary.items()), columns=['Termino', 'Frecuencia'])
    df_stemmed.to_excel(os.path.join(output_folder, '4DiccDeSteams.xlsx'), index=False)
    print('Diccionario con stemming guardado en "4DiccDeSteams.xlsx".')

    # Proceso 5: Tbla de hash
    index_steams()
    
    # Llamar a la función con el argumento correcto
    save_document_word_counts(document_word_counts, output_folder, stemmed_dictionary)
    print("Frecuencia de términos por documento guardada correctamente.")

    total_documents = count_documents(corpus_dir)
    calculate_tf(Frecuencia_Terminos, total_documents, output_folder)

    # Leer el archivo de conteo de documentos por término
    document_term_count_file = os.path.join(output_folder, 'Conteo_Documentos_Termino.xlsx')
    # Calcular el IDF
    calculate_idf(total_documents, document_term_count_file, output_folder)

    # Al final del procesamiento, después de calcular TF e IDF
    calculate_tf_idf(Frecuencia_Terminos, os.path.join(output_folder, 'IDF.xlsx'), output_folder)
    

def save_document_word_counts(document_word_counts, output_folder, stemmed_dictionary):
    # Obtener todas las palabras únicas del diccionario con stemming
    all_words = set(stemmed_dictionary.keys())

    # Filtrar palabras que contienen números
    all_words = {word for word in all_words if not any(char.isdigit() for char in word)}

    all_words = sorted(list(all_words))  # Ordenar todas las palabras alfabéticamente
    
    # Crear un DataFrame con los nombres de los documentos en las filas y los términos en las columnas
    df = pd.DataFrame(index=list(document_word_counts.keys()), columns=all_words)

    # Llenar el DataFrame con las frecuencias
    for doc, counts in document_word_counts.items():
        for word in all_words:
            df.at[doc, word] = counts.get(word, 0)  # Poner el conteo o 0 si no aparece la palabra

    # Convertir todos los términos a mayúsculas en el DataFrame
    df.columns = [word.upper() for word in df.columns]

    # Guardar el DataFrame en un archivo Excel
    output_file = os.path.join(output_folder, 'Frecuencia_Terminos_Por_Documento.xlsx')
    df.to_excel(output_file, index=True)
    print(f"Frecuencia de términos por documento guardada en: {output_file}")

    # Crear un diccionario para contar en cuántos documentos aparece cada término
    term_document_count = {}

    for word in all_words:
        term_count_in_docs = (df[word] > 0).sum()  # Contar cuántos documentos tienen al menos una aparición del término
        term_document_count[word] = term_count_in_docs

    # Guardar el conteo de documentos por término en un archivo Excel
    df_doc_count = pd.DataFrame(list(term_document_count.items()), columns=['Termino', 'Documentos_Contienen_Termino'])
    output_doc_count_file = os.path.join(output_folder, 'Conteo_Documentos_Termino.xlsx')
    df_doc_count.to_excel(output_doc_count_file, index=False)
    print(f"Conteo de documentos por término guardado en: {output_doc_count_file}")

def calculate_tf(Frecuencia_Terminos, total_documents, output_folder):
    # Leer el archivo Excel existente
    df = pd.read_excel(Frecuencia_Terminos, index_col=0)  # Usa la primera columna como índice (nombres de documentos)
    
    # Dividir cada valor por el total de documentos para obtener TF
    df_tf = df.div(total_documents)

    # Guardar el DataFrame en un nuevo archivo Excel
    output_file = os.path.join(output_folder, 'TF.xlsx')
    df_tf.to_excel(output_file, index=True)
    print(f"TF guardado en: {output_file}")

Frecuencia_Terminos = 'PreprocesamientoSteps/Frecuencia_Terminos_Por_Documento.xlsx'  # Asegúrate de dar la ruta correcta

def calculate_idf(total_documents, document_term_count_file, output_folder):
    # Leer el archivo de conteo de documentos por término
    df_term_doc_count = pd.read_excel(document_term_count_file)

    # Calcular el IDF
    df_term_doc_count['IDF'] = df_term_doc_count['Documentos_Contienen_Termino'].apply(
        lambda doc_count: math.log2(total_documents / doc_count) if doc_count > 0 else 0
    )
    # Crear un DataFrame con los términos como columnas
    df_idf = pd.DataFrame(df_term_doc_count['IDF']).T  # Transponer para que los términos sean columnas
    df_idf.columns = df_term_doc_count['Termino']  # Asignar los términos como nombres de las columnas

    # Guardar el DataFrame transpuesto en un archivo Excel
    output_file = os.path.join(output_folder, 'IDF.xlsx')
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    # Escribir el DataFrame en el archivo, comenzando desde la celda B1
        df_idf.to_excel(writer, startrow=0, startcol=1, index=False, header=True)

    
    print(f"IDF guardado en: {output_file}")

def calculate_tf_idf(tf_file, idf_file, output_folder):
    # Leer los archivos de TF y IDF
    df_tf = pd.read_excel(tf_file, index_col=0)  # Cargar TF y usar el nombre de documentos como índice
    df_idf = pd.read_excel(idf_file, header=0)  # Cargar IDF y usar la primera fila como encabezado

    # Obtener los valores de IDF (la única fila de valores)
    idf_values = df_idf.iloc[0].values  # Obtener la fila de IDF
    idf_terms = df_idf.columns  # Obtener los términos de IDF

    # Asegurarse de que los términos de TF y IDF coincidan
    common_terms = df_tf.columns.intersection(idf_terms)  # Términos comunes entre TF e IDF

    if len(common_terms) == 0:
        raise ValueError("No hay términos en común entre TF e IDF.")

    # Filtrar DF TF para mantener solo los términos comunes
    df_tf_common = df_tf[common_terms]

    # Filtrar los valores IDF para que coincidan con los términos comunes
    idf_values_common = [idf_values[df_idf.columns.get_loc(term)] for term in common_terms]

    # Multiplicar cada valor en df_tf_common por su correspondiente en idf_values_common
    df_tf_idf = df_tf_common.multiply(idf_values_common, axis=1)

    # Guardar el DataFrame de TF-IDF en un archivo Excel
    output_file = os.path.join(output_folder, ' 8MatrizTF-IDFdeSteams.xlsx')
    df_tf_idf.to_excel(output_file, index=True)
    print(f"TF-IDF guardado en: {output_file}")


# FUCIONES PARA EL PROCESO DE CONSULTA


# Inicializar el lemmatizer
lemmatizer = WordNetLemmatizer()

# Cargar el Excel
archivo_excel = 'PreprocesamientoSteps/8MatrizTF-IDFdeSteams.xlsx'
hashs = 'PreprocesamientoSteps/5ListDiccIndex.xlsx'
corpus_dir = 'corpus'


# Ruta del archivo de frecuencias de términos
Frecuencia_Terminos = 'PreprocesamientoSteps/Frecuencia_Terminos_Por_Documento.xlsx'

# 1. Cargar el archivo Excel
df_frecuencia = pd.read_excel(Frecuencia_Terminos)

# 2. Leer la consulta en lenguaje natural
consulta_q = input("Ingrese la consulta en lenguaje natural: ")

# 3. Procesar la consulta (sin aplicar stemming)
def procesar_consulta(consulta_q):
    stop_words = set(stopwords.words('spanish'))
    tokens = consulta_q.split()
    
    # Filtrar stopwords (opcional, según tu preferencia)
    consulta_procesada = [
        token for token in tokens if token not in stop_words
    ]

    return consulta_procesada

consulta_procesada = procesar_consulta(consulta_q)
print("Consulta procesada:", consulta_procesada)

# 4. Verificar si los términos procesados están en las columnas del DataFrame
terminos_presentes = {}
for term in consulta_procesada:
    terminos_presentes[term] = term in df_frecuencia.columns


# Crear un diccionario para almacenar los términos y sus valores como columnas
data_dict = {}

# 5. Recopilar términos y valores en el diccionario
for term, existe in terminos_presentes.items():
    if existe:
        # Extraer los valores de la columna correspondiente
        valores = df_frecuencia[term]
        data_dict[term] = valores.tolist()  # Guardar los valores como una lista en el diccionario
    else:
        print(f"El término '{term}' NO está presente en el archivo Excel.")

# Crear un DataFrame a partir del diccionario
df_resultados = pd.DataFrame(data_dict)

# Guardar el DataFrame en un archivo Excel
df_resultados.to_excel('PreprocesamientoSteps/resultados_consult.xlsx', index=False)

print("El archivo Excel ha sido creado exitosamente.")


# Función para iniciar el proceso completo
def preprocess_documents():
    
    # Contar los documentos en el corpus
    total_documents = count_documents(corpus_dir)
    print(f"Número total de documentos en el corpus: {total_documents}")

    if os.path.exists(corpus_dir):
        tokenize_and_process_documents(corpus_dir)
    else:
        print("El directorio 'Corpus' no existe.")

# Ejecutar el procesamiento
preprocess_documents()
