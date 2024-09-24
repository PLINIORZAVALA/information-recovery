import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from arbol import BTree  # Importamos el árbol B

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("B-Tree App")
        self.b_tree = BTree(t=4)  # Crear un árbol B de grado 4
        self.create_widgets()

    def create_widgets(self):
        load_button = tk.Button(self.root, text="Cargar STEAMS", command=self.load_steams)
        load_button.pack(pady=10)

        search_label = tk.Label(self.root, text="Buscar palabra:")
        search_label.pack(pady=5)

        self.search_entry = tk.Entry(self.root)
        self.search_entry.pack(pady=5)

        search_button = tk.Button(self.root, text="Buscar", command=self.search_word)
        search_button.pack(pady=10)

    def read_steams_from_excel(self, file_path):
        df = pd.read_excel(file_path)
        # Convierte todos los términos a cadena
        return df['Termino'].astype(str).tolist()

    def load_steams(self):
        # Leer el archivo de Excel
        file_path = "PreprocesamientoSteps/4DiccDeSteams.xlsx"
        if not os.path.exists(file_path):
            messagebox.showerror("Error", f"El archivo {file_path} no existe.")
            return
        
        steams = self.read_steams_from_excel(file_path)
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
                messagebox.showinfo("Resultado", f"La palabra '{word}' fue encontrada en el árbol.")
            else:
                messagebox.showinfo("Resultado", f"La palabra '{word}' NO fue encontrada en el árbol.")
        else:
            messagebox.showwarning("Entrada vacía", "Por favor, ingresa una palabra para buscar.")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
