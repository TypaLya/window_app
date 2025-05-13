class GlassCutter:
    def __init__(self, items):
        self.sheet_width = 6000
        self.sheet_height = 6000
        self.items = sorted(items, key=lambda x: max(x[1], x[2]), reverse=True)  # сортировка по макс. размеру
        self.groups = []  # список листов (каждый — словарь с размещёнными заготовками)

    def pack_items(self):
        for order_id, width, height in self.items:
            placed = False
            for group in self.groups:
                if self.try_place(group, order_id, width, height):
                    placed = True
                    break
            if not placed:
                # создаем новый лист
                new_group = {"items": [], "positions": []}
                self.try_place(new_group, order_id, width, height)
                self.groups.append(new_group)

    def try_place(self, group, order_id, width, height):
        scale = 1  # работа в натуральных единицах
        if not group["positions"]:
            # первая заготовка — в левый верхний угол
            group["positions"].append((0, 0, width, height))
            group["items"].append((order_id, width, height))
            return True

        # проходим по всей сетке листа с шагом 10мм
        for y in range(0, self.sheet_height - height + 1, 10):
            for x in range(0, self.sheet_width - width + 1, 10):
                if self.fits(group["positions"], x, y, width, height):
                    group["positions"].append((x, y, width, height))
                    group["items"].append((order_id, width, height))
                    return True
        return False

    def fits(self, placed, x, y, width, height):
        for px, py, pw, ph in placed:
            if not (x + width <= px or x >= px + pw or y + height <= py or y >= py + ph):
                return False
        return True
