from random import random

from Item import NumberItem


class ItemManager:
    def __init__(self, num_items, value_range=(30, 150)):
        self.num_items = num_items  # Количество элементов
        self.value_range = value_range  # Диапазон значений для чисел

    def create_items(self):
        items = []
        for i in range(1, self.num_items + 1):
            values = [random.randint(*self.value_range) for _ in range(4)]  # Генерируем 4 значения
            item = NumberItem(i, values)  # Создаем объект с 4 значениями
            items.append(item)
        return items