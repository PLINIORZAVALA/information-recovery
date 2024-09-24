import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import random
from arbol import BTree  # Importamos el árbol B para su uso

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
    steams = df['Termino'].astype(str).tolist()  # Aseguramos que todo sea cadena
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

# Clase principal para integrar el Árbol B y la interfaz
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Cronus con Árbol B")
        self.root.configure(bg='black')
        self.root.attributes('-fullscreen', True)
        
        self.b_tree = BTree(t=4)  # Crear un árbol B de grado 4
        self.hash_table = {}  # Tabla hash vacía
        self.create_widgets()

    def create_widgets(self):
        # Estilo
        style = ttk.Style()
        style.configure("TButton", font=("arial", 16))
        style.configure("TEntry", font=("arial", 16))

        search_frame = tk.Frame(self.root, bg='black')
        search_frame.pack(pady=20)

        self.search_entry = ttk.Entry(search_frame, width=30, style="TEntry")
        self.search_entry.pack(side=tk.LEFT, padx=5)

        search_button = ttk.Button(search_frame, text="🔍", command=self.search_word, style="TButton")
        search_button.pack(side=tk.LEFT)

        # Botón para preprocesar documentos
        preprocess_button = ttk.Button(self.root, text="Preprocesado", command=preprocess_documents, style="TButton")
        preprocess_button.pack(pady=10)

        # Botón para indexar los steams
        index_button = ttk.Button(self.root, text="Indexación", command=self.index_steams, style="TButton")
        index_button.pack(pady=10)

        # Botón para cargar steams en el Árbol B
        load_button = ttk.Button(self.root, text="Cargar STEAMS", command=self.load_steams, style="TButton")
        load_button.pack(pady=10)

        # Etiqueta para mostrar el resultado de la búsqueda
        self.result_label = tk.Label(self.root, text="", fg="white", bg="black", font=("arial", 16))
        self.result_label.pack(pady=10)

        # Canvas para mostrar documentos del corpus
        canvas = tk.Canvas(self.root, bg='gray')
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
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

        self.root.bind("<Escape>", lambda e: self.root.attributes("-fullscreen", False))

    def load_steams(self):
        # Leer el archivo de Excel
        file_path = "PreprocesamientoSteps/4DiccDeSteams.xlsx"
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"El archivo {file_path} no existe.")
            return
        
        steams = read_steams_from_excel(file_path)
        for steam in steams:
            try:
                self.b_tree.insert(steam)
            except Exception as e:
                messagebox.showerror("Error de inserción", f"Error al insertar '{steam}': {str(e)}")

        self.b_tree.display()  # Mostrar el árbol en consola
        messagebox.showinfo("Éxito", "STEAMS cargados en el árbol.")

    def search_word(self):
        word = self.search_entry.get()
        if word:
            found = self.b_tree.search(word)
            if found:
                self.result_label.config(text=f"La palabra '{word}' fue encontrada en el árbol.")
            else:
                self.result_label.config(text=f"La palabra '{word}' NO fue encontrada en el árbol.")
        else:
            messagebox.showwarning("Entrada vacía", "Por favor, ingresa una palabra para buscar.")

    def index_steams(self):
        self.hash_table = index_steams()
        if self.hash_table:
            messagebox.showinfo("Éxito", "Indexación completada.")

def preprocess_documents():
    corpus_dir = 'Corpus'
    if os.path.exists(corpus_dir):
        print("Preprocesando documentos...")
        # Aquí puedes implementar las funciones de preprocesamiento si es necesario.
    else:
        print("El directorio 'Corpus' no existe.")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
