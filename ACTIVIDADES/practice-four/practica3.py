import tkinter as tk
from tkinter import ttk
import os
import pandas as pd
import random

# Función para abrir documentos (para la interfaz gráfica)
def opendocument(link):
    try:
        if os.name == 'nt':  # Para Windows
            os.startfile(link)
    except Exception as e:
        print(f"Error al intentar abrir el documento: {e}")

# Crear carpeta si no existe
def create_folder():
    folder_name = "PreprocesamientoSteps"
    if not os.path.exists(folder_name):
        os.makedirs(folder_name)
        print(f'Carpeta "{folder_name}" creada con éxito.')
    else:
        print(f'La carpeta "{folder_name}" ya existe.')
    return folder_name

# Leer archivo con los steams
def read_steams_from_excel(file_path):
    df = pd.read_excel(file_path)
    steams = df['Termino'].tolist()
    return steams

# Función para indexar (hashing) los steams
def index_steams():
    # Leer los steams del archivo generado en el proceso anterior (4DiccDeSteams.xlsx)
    file_path = "PreprocesamientoSteps/4DiccDeSteams.xlsx"
    if not os.path.exists(file_path):
        print(f"El archivo {file_path} no existe.")
        return

    steams = read_steams_from_excel(file_path)

    # Revolver los steams (mezclar aleatoriamente)
    random.shuffle(steams)

    # Crear una tabla hash donde cada steam tendrá un valor hash
    steam_hash_table = {steam: hash(steam) for steam in steams}

    # Guardar la tabla en un archivo Excel
    output_folder = create_folder()
    df = pd.DataFrame(list(steam_hash_table.items()), columns=['Steam', 'Hash'])
    output_file = os.path.join(output_folder, '5ListDiccIndex.xlsx')
    df.to_excel(output_file, index=False)
    print(f"Tabla hash guardada en {output_file}.")
    return steam_hash_table  # Devolvemos la tabla hash para su uso posterior

# Función para consultar si un steam existe en la tabla hash
def search_steam(steam, hash_table):
    if steam in hash_table:
        return f"El steam '{steam}' existe con el valor hash: {hash_table[steam]}"
    else:
        return f"El steam '{steam}' no existe en la tabla hash."

# Función principal del preprocesado de documentos
def preprocess_documents():
    corpus_dir = 'Corpus'
    if os.path.exists(corpus_dir):
        tokenize_and_process_documents(corpus_dir)
    else:
        print("El directorio 'Corpus' no existe.")

# Función para manejar la búsqueda de steams en la interfaz
def handle_search(search_entry, result_label, hash_table):
    steam = search_entry.get()
    result = search_steam(steam, hash_table)
    result_label.config(text=result)  # Mostramos el resultado en la etiqueta

# Crear la aplicación gráfica con Tkinter
def createapp():
    root = tk.Tk()
    root.title("Cronus")
    root.configure(bg='black')
    root.attributes('-fullscreen', True)

    style = ttk.Style()
    style.configure("TButton", font=("arial", 16))
    style.configure("TEntry", font=("arial", 16))

    search_frame = tk.Frame(root, bg='black')
    search_frame.pack(pady=20)

    search_entry = ttk.Entry(search_frame, width=30, style="TEntry")
    search_entry.pack(side=tk.LEFT, padx=5)

    search_button = ttk.Button(search_frame, text="🔍", style="TButton")
    search_button.pack(side=tk.LEFT)

    preprocess_button = ttk.Button(root, text="Preprocesado", command=preprocess_documents, style="TButton")
    preprocess_button.pack(pady=10)

    # Botón para indexar los steams (agregar paso de indexación)
    hash_table = {}  # Creamos una tabla hash vacía

    def index_and_update_table():
        nonlocal hash_table
        hash_table = index_steams()

    index_button = ttk.Button(root, text="Indexación", command=index_and_update_table, style="TButton")
    index_button.pack(pady=10)

    # Etiqueta para mostrar el resultado de la búsqueda
    result_label = tk.Label(root, text="", fg="white", bg="black", font=("arial", 16))
    result_label.pack(pady=10)

    # Actualizar el botón de búsqueda para que llame a handle_search
    search_button.config(command=lambda: handle_search(search_entry, result_label, hash_table))

    canvas = tk.Canvas(root, bg='gray')
    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollbar.pack(side=tk.RIGHT, fill='y')

    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    frame = tk.Frame(canvas, bg='gray')
    canvas.create_window((0, 0), window=frame, anchor='nw')

    corpus_dir = 'Corpus'
    if not os.path.exists(corpus_dir):
        os.makedirs(corpus_dir)

    documents = [f for f in os.listdir(corpus_dir) if os.path.isfile(os.path.join(corpus_dir, f))]

    for doc in documents:
        doc_path = os.path.join(corpus_dir, doc)
        link = tk.Label(frame, text=doc, fg="white", cursor="hand2", bg='gray', font=("arial", 16))
        link.pack(anchor='w', pady=2)
        link.bind("<Button-1>", lambda e, url=doc_path: opendocument(url))

    root.bind("<Escape>", lambda e: root.attributes("-fullscreen", False))
    root.mainloop()

if __name__ == "__main__":
    createapp()
