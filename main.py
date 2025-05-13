class NumberItem:
    def __init__(self, index, values):
        self.index = index  # Порядковый номер элемента
        self.values = values  # Массив из 4 чисел: [длина, ширина, длина, ширина]

    def __repr__(self):
        return f"NumberItem(index={self.index}, values={self.values})"