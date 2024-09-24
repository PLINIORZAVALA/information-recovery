class BTreeNode:
    def __init__(self, t, leaf=False):
        self.t = t  # Grado mínimo
        self.leaf = leaf  # Si es un nodo hoja
        self.keys = []  # Lista de claves en el nodo
        self.children = []  # Lista de hijos

    def insert_non_full(self, key):
        i = len(self.keys) - 1
        if self.leaf:
            # Insertar clave en el nodo hoja
            self.keys.append(None)  # Añadir un espacio vacío
            while i >= 0 and key < self.keys[i]:
                self.keys[i + 1] = self.keys[i]  # Desplazar claves hacia la derecha
                i -= 1
            self.keys[i + 1] = key  # Insertar la nueva clave
        else:
            # Encontrar el hijo donde insertar
            while i >= 0 and key < self.keys[i]:
                i -= 1
            i += 1
            if len(self.children[i].keys) == 2 * self.t - 1:
                self.split_child(i)  # Dividir el hijo
                if key > self.keys[i]:
                    i += 1
            self.children[i].insert_non_full(key)

    def split_child(self, i):
        t = self.t
        y = self.children[i]  # Niño que será dividido
        z = BTreeNode(t, y.leaf)  # Nuevo nodo
        self.children.insert(i + 1, z)  # Añadir nuevo hijo al nodo actual
        self.keys.insert(i, y.keys[t - 1])  # Subir la clave media

        # Copiar las últimas t-1 claves de y a z
        z.keys = y.keys[t:]  
        y.keys = y.keys[:t - 1]  # Mantener las primeras t-1 claves en y

        # Si no es hoja, mover sus hijos
        if not y.leaf:
            z.children = y.children[t:]  
            y.children = y.children[:t]

class BTree:
    def __init__(self, t):
        self.root = BTreeNode(t, True)  # Crear nodo raíz

    def insert(self, key):
        root = self.root
        if len(root.keys) == 2 * root.t - 1:  # Si está lleno
            new_root = BTreeNode(root.t)  # Crear nueva raíz
            new_root.children.append(root)  # Hacer raíz hijo
            new_root.split_child(0)  # Dividir la antigua raíz
            i = 0
            if new_root.keys[0] < key:
                i += 1
            new_root.children[i].insert_non_full(key)
            self.root = new_root  # Cambiar raíz
        else:
            root.insert_non_full(key)

    def display(self, node=None, level=0):
        if node is None:
            node = self.root
        print("Level", level, "Keys:", node.keys)
        level += 1
        for child in node.children:
            self.display(child, level)

