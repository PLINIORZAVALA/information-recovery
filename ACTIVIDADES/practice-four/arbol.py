class BTreeNode:
    def __init__(self, leaf=False):
        self.leaf = leaf
        self.keys = []
        self.children = []

class BTree:
    def __init__(self, t):
        self.root = BTreeNode(True)
        self.t = t  # El grado mínimo del árbol B

    def insert(self, key):
        root = self.root
        if len(root.keys) == (2 * self.t) - 1:  # Si la raíz está llena
            new_node = BTreeNode()
            self.root = new_node
            new_node.children.append(root)
            self.split_child(new_node, 0)
            self.insert_non_full(new_node, key)
        else:
            self.insert_non_full(root, key)

    def split_child(self, parent, i):
        t = self.t
        y = parent.children[i]
        z = BTreeNode(y.leaf)
        parent.children.insert(i + 1, z)
        parent.keys.insert(i, y.keys[t - 1])
        z.keys = y.keys[t:(2 * t) - 1]
        y.keys = y.keys[0:t - 1]
        if not y.leaf:
            z.children = y.children[t:(2 * t)]
            y.children = y.children[0:t]

    def insert_non_full(self, node, key):
        i = len(node.keys) - 1
        if node.leaf:
            node.keys.append(None)
            while i >= 0 and key < node.keys[i]:
                node.keys[i + 1] = node.keys[i]
                i -= 1
            node.keys[i + 1] = key
        else:
            while i >= 0 and key < node.keys[i]:
                i -= 1
            i += 1
            if len(node.children[i].keys) == (2 * self.t) - 1:
                self.split_child(node, i)
                if key > node.keys[i]:
                    i += 1
            self.insert_non_full(node.children[i], key)

    def search(self, key, node=None):
        if node is None:
            node = self.root

        i = 0
        # Buscar la primera clave mayor o igual que la buscada
        while i < len(node.keys) and key > node.keys[i]:
            i += 1

        # Si la clave encontrada es igual a la clave buscada, retornamos True
        if i < len(node.keys) and key == node.keys[i]:
            return True  # La clave existe en el nodo

        # Si el nodo es una hoja, la clave no está en el árbol
        if node.leaf:
            return False

        # Si no es una hoja, bajamos al hijo adecuado
        return self.search(key, node.children[i])

    def display(self, node=None, level=0):
        if node is None:
            node = self.root
        print("Level", level, "Keys:", node.keys)
        if not node.leaf:
            for child in node.children:
                self.display(child, level + 1)
