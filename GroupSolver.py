from itertools import combinations


class GroupSolver:
    def __init__(self, items, target_sum=6000, min_sum_ratio=0.9):
        self.items = items  # Список объектов NumberItem
        self.target_sum = target_sum  # Целевая сумма для группы
        self.min_sum = target_sum * min_sum_ratio  # Минимальная допустимая сумма для группы

    def group_numbers(self):
        groups = []
        unused_values = {}

        # Преобразуем список объектов NumberItem в словарь {index: values}
        numbers_dict = {item.index: item.values[:] for item in self.items}

        while numbers_dict:
            best_combination = []
            best_sum = 0

            # Преобразуем словарь в список кортежей, где каждый элемент - (ключ, значение-массив)
            items = list(numbers_dict.items())

            # Перебираем комбинации элементов массивов из разных ключей
            for r in range(1, len(items) + 1):
                for comb in combinations(items, r):
                    # Подготовим список всех значений из массивов текущей комбинации
                    all_values = []
                    for key, values in comb:
                        all_values.extend([(key, value) for value in values])

                    # Перебираем все подмножества значений, которые можно взять для суммы
                    for k in range(1, len(all_values) + 1):
                        for subset in combinations(all_values, k):
                            # Считаем сумму выбранного подмножества
                            current_sum = sum(value for _, value in subset)

                            # Проверяем, подходит ли комбинация для текущей группы
                            if best_sum < current_sum <= self.target_sum:
                                best_combination = subset
                                best_sum = current_sum
                                # Если сумма достигла целевой, прекращаем дальнейший перебор
                                if best_sum == self.target_sum:
                                    break
                        if best_sum == self.target_sum:
                            break
                    if best_sum == self.target_sum:
                        break

            # Если не удалось собрать группу с минимальной суммой, выходим
            if best_sum < self.min_sum:
                for key, values in numbers_dict.items():
                    if key not in unused_values:
                        unused_values[key] = values
                    else:
                        unused_values[key].extend(values)
                break

            # Формируем новую группу из лучшей комбинации
            group_dict = {}
            used_values = {}  # Для хранения использованных значений по ключам

            for key, value in best_combination:
                if key not in group_dict:
                    group_dict[key] = []
                group_dict[key].append(value)

                # Отслеживаем использованные значения
                if key not in used_values:
                    used_values[key] = []
                used_values[key].append(value)

            groups.append(group_dict)

            # Удаляем использованные значения из словаря
            for key, values in used_values.items():
                if key in numbers_dict:
                    for value in values:
                        if value in numbers_dict[key]:
                            numbers_dict[key].remove(value)
                    if not numbers_dict[key]:  # Если больше значений не осталось
                        del numbers_dict[key]

        return groups, unused_values

