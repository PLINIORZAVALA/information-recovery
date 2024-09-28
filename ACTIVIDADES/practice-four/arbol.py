class BTreeNode:
    def __init__(self, leaf=False):
        self.leaf = leaf  # Indica si el nodo es una hoja
        self.keys = []  # Claves almacenadas en el nodo
        self.children = []  # Hijos del nodo (solo si no es hoja)

class BTree:
    def __init__(self, t):
        if t < 3:
            raise ValueError("El valor de t debe ser al menos 3 para que los nodos tengan al menos 2 claves.")
        self.root = BTreeNode(True)  # Crear un nodo raíz vacío
        self.t = t  # El número máximo de claves por nodo será t

    def insert(self, key):
        root = self.root
        # Si la raíz está llena, necesitamos dividirla
        if len(root.keys) == self.t:  # Cambiado de (2 * t - 1) a t
            new_node = BTreeNode()  # Nuevo nodo raíz
            self.root = new_node
            new_node.children.append(root)  # Hacer que la raíz anterior sea hijo del nuevo nodo
            self.split_child(new_node, 0)  # Dividir el primer hijo
            self.insert_non_full(new_node, key)  # Insertar la clave en el nuevo nodo
        else:
            self.insert_non_full(root, key)  # Insertar normalmente si la raíz no está llena

    def split_child(self, parent, i):
        t = self.t
        y = parent.children[i]  # Nodo a dividir
        z = BTreeNode(y.leaf)  # Nuevo nodo que almacenará la mitad de las claves de y
        parent.children.insert(i + 1, z)  # Insertar el nuevo nodo en los hijos del padre
        parent.keys.insert(i, y.keys[t // 2])  # Subir la clave mediana al padre

        # Dividir las claves entre y y z
        z.keys = y.keys[(t // 2) + 1:]  # Asignar las claves superiores a z
        y.keys = y.keys[:t // 2]  # Mantener las claves inferiores en y

        # Si y no es hoja, dividir también los hijos
        if not y.leaf:
            z.children = y.children[(t // 2) + 1:]
            y.children = y.children[:t // 2 + 1]

    def insert_non_full(self, node, key):
        i = len(node.keys) - 1
        if node.leaf:
            node.keys.append(None)  # Añadir espacio para una nueva clave
            while i >= 0 and key < node.keys[i]:
                node.keys[i + 1] = node.keys[i]  # Desplazar claves hacia la derecha
                i -= 1
            node.keys[i + 1] = key  # Insertar la nueva clave en su lugar
        else:
            while i >= 0 and key < node.keys[i]:
                i -= 1
            i += 1
            if len(node.children[i].keys) == self.t:  # Si el hijo está lleno (máximo t claves)
                self.split_child(node, i)  # Dividir el hijo
                if key > node.keys[i]:
                    i += 1
            self.insert_non_full(node.children[i], key)  # Insertar en el hijo adecuado

    def search(self, key, node=None):
        if node is None:
            node = self.root
        i = 0
        while i < len(node.keys) and key > node.keys[i]:
            i += 1
        if i < len(node.keys) and key == node.keys[i]:
            return True  # La clave fue encontrada
        if node.leaf:
            return False  # Si llegamos a una hoja y no encontramos la clave
        return self.search(key, node.children[i])  # Buscar en el hijo adecuado

    def display(self, node=None, level=0):
        if node is None:
            node = self.root
        print("Level", level, "Keys:", node.keys)
        if not node.leaf:
            for child in node.children:
                self.display(child, level + 1)