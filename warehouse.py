import os
import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox

import xlrd
from customtkinter import CTkButton, CTkFrame

from database import get_all_film_materials, add_film_material, add_component_material, get_all_component_materials, \
    get_all_window_materials, add_window_material, get_all_triplex_materials, add_triplex_material, \
    get_all_main_glass_materials, add_main_glass_material, get_all_cutting_materials, add_cutting_material


def safe_float(value):
    try:
        return float(value)
    except (ValueError, TypeError):
        return None


def parse_excel_data(file_path, flag_cutting = False):
    """Парсинг данных заказа из Excel файла (поддержка .xls и .xlsx)"""
    try:
        # Определяем расширение файла
        ext = file_path.split('.')[-1].lower()
        if flag_cutting:
            warehouse_data = {
                'num': [],
                'name': [],
                'unit': [],
                'photo': [],
                'start_balance': [],
                'income': [],
                'outcome': [],
                'end_balance': [],
                'reserved': [],
                'in_transit': [],
                'price': [],
                'total_sum': [],
                'description': []
            }
        else:
            warehouse_data = {
                'num': [],
                'name': [],
                'unit': [],
                'photo': [],
                'start_balance': [],
                'income': [],
                'outcome': [],
                'end_balance': [],
                'reserved': [],
                'in_transit': [],
                'price': [],
                'total_sum': [],
                'last_income_date': [],
                'last_move_date': [],
                'description': []
            }

        if ext == 'xls':
            book = xlrd.open_workbook(file_path, encoding_override='cp1251')
            sheet = book.sheet_by_index(0)

            # Парсим список материалов
            start_parsing = False
            skip1 = skip2 = False
            for row_idx in range(sheet.nrows):
                if skip1:
                    skip1 = False
                    skip2 = True
                    continue
                elif skip2:
                    skip2 = False
                    continue
                row = sheet.row_values(row_idx)
                # Ищем начало таблицы со стеклопакетами
                if row and len(row) > 1:
                    first_cell = str(row[0]).strip()
                    if first_cell in ["№"] and "Склад" in str(row[1]):
                        start_parsing = True
                        skip1 = True
                        # print(row)
                        continue

                if start_parsing and row:
                    try:
                        if isinstance(safe_float(row[0]), float) and row[1].isdigit():
                            if flag_cutting:
                                item_data = {
                                    'num': str(row[0]).strip(),
                                    'name': f"{str(row[1]).strip()} {str(row[2]).strip()}",
                                    'unit': str(row[4]).strip(),
                                    'photo': str(row[5]).strip(),
                                    'start_balance': str(row[6]).strip(),
                                    'income': str(row[7]).strip(),
                                    'outcome': str(row[8]).strip(),
                                    'end_balance': str(row[9]).strip(),
                                    'reserved': str(row[10]).strip(),
                                    'in_transit': str(row[11]).strip(),
                                    'price': str(row[12]).strip(),
                                    'total_sum': str(row[13]).strip(),
                                    'description': str(row[14]).strip()
                                }
                            else:
                                item_data = {
                                    'num': str(row[0]).strip(),
                                    'name': f"{str(row[1]).strip()} {str(row[2]).strip()}",
                                    'unit': str(row[4]).strip(),
                                    'photo': str(row[5]).strip(),
                                    'start_balance': str(row[6]).strip(),
                                    'income': str(row[7]).strip(),
                                    'outcome': str(row[8]).strip(),
                                    'end_balance': str(row[9]).strip(),
                                    'reserved': str(row[10]).strip(),
                                    'in_transit': str(row[11]).strip(),
                                    'price': str(row[12]).strip(),
                                    'total_sum': str(row[13]).strip(),
                                    'last_income_date': str(row[14]).strip(),
                                    'last_move_date': str(row[15]).strip(),
                                    'description': str(row[16]).strip()
                                }

                            # Добавляем данные в warehouse_data
                            for key, value in item_data.items():
                                warehouse_data[key].append(value)
                                # print(f"{key} = {value}")

                            # for key, value in window_data.items():
                            #     print(f"{key}: {value}")
                        else:
                            start_parsing = False
                    except (ValueError, IndexError, AttributeError) as e:
                        print(f"Ошибка при парсинге строки {row_idx}: {e}")
        else:
            raise Exception("Неподдерживаемый формат файла. Используйте .xls или .xlsx")

        return warehouse_data

    except Exception as e:
        raise Exception(f"Ошибка при чтении файла: {str(e)}")




class WarehouseTab(CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.main_glass_tree = None
        self.window_tree = None
        self.components_tree = None
        self.main_glass_tab = None
        self.triplex_tab = None
        self.glass_units_tree = None
        self.glass_units_tab = None
        self.film_tab = None
        self.components_tab = None
        self.parent = parent

        # Создаем Notebook для вкладок
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Создаем вкладки
        self.create_components_tab()
        self.create_film_tab()
        self.create_windows_tab()
        self.create_triplex_tab()
        self.create_main_glass_tab()
        self.create_cutting_tab()

        # Загружаем данные
        self.load_all_data()

    def create_components_tab(self):
        """Вкладка материалов участка наклейки пленки с расширенной таблицей"""
        self.components_tab = CTkFrame(self.notebook)
        self.notebook.add(self.components_tab, text="Материалы cклада комплектующих")

        # Панель управления
        control_frame = CTkFrame(self.components_tab)
        control_frame.pack(fill=tk.X, pady=5)

        CTkButton(control_frame, text="Обновить", command=self.load_components_data).pack(side=tk.LEFT, padx=5)

        # Таблица с горизонтальной прокруткой
        container = CTkFrame(self.components_tab)
        container.pack(fill=tk.BOTH, expand=True)

        # Создаем Treeview с горизонтальной прокруткой
        self.components_tree = ttk.Treeview(container,
                                      columns=("num", "name", "unit", "photo",
                                               "start_balance", "income", "outcome", "end_balance",
                                               "reserved", "in_transit", "price", "total_sum",
                                               "last_income_date", "last_move_date", "description"),
                                      show="headings",
                                      height=20)

        # Настройка столбцов
        columns = [
            ("№", 30), ("Склад/ТМЦ/Документы движения", 250), ("Ед. изм.", 80),
            ("Фото", 50), ("Начальный остаток", 100), ("Приход", 80), ("Расход", 80),
            ("Конечный остаток", 100), ("в т.ч. резерв", 80), ("в пути", 80),
            ("Цена остатка", 100), ("Сумма остатка", 100), ("Дата последнего прихода", 120),
            ("Дата последнего перемещения", 120), ("Описание [назначение]", 100)
        ]

        for idx, (text, width) in enumerate(columns):
            self.components_tree.heading(f"#{idx + 1}", text=text)
            self.components_tree.column(f"#{idx + 1}", width=width, anchor='center')

        # Вертикальная прокрутка
        y_scroll = ttk.Scrollbar(container, orient="vertical", command=self.components_tree.yview)
        self.components_tree.configure(yscrollcommand=y_scroll.set)

        # Горизонтальная прокрутка
        x_scroll = ttk.Scrollbar(container, orient="horizontal", command=self.components_tree.xview)
        self.components_tree.configure(xscrollcommand=x_scroll.set)

        # Размещение элементов
        self.components_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

    def load_components_data(self):
        """Загрузка данных склада комплектующих"""
        # Очищаем таблицу
        for item in self.components_tree.get_children():
            self.components_tree.delete(item)

        try:
            file_path = "warehouse_table/Ведомость материалы Склад комплектующих.xls"
            if not os.path.exists(file_path):
                messagebox.showwarning("Внимание", "Файл с данными склада не найден. Загружаем из базы данных.")
                materials = get_all_film_materials()
            else:
                # Парсим данные из Excel
                warehouse_data = parse_excel_data(file_path)

                # Сохраняем данные в БД
                conn = sqlite3.connect('orders.db')
                cursor = conn.cursor()

                # Очищаем старые данные
                cursor.execute("DELETE FROM components_warehouse")
                conn.commit()

                # Вставляем новые данные
                num_records = len(warehouse_data['num'])
                for i in range(num_records):
                    data = {
                        'num': warehouse_data['num'][i],
                        'name': warehouse_data['name'][i],
                        'unit': warehouse_data['unit'][i],
                        'photo': warehouse_data['photo'][i],
                        'start_balance': warehouse_data['start_balance'][i],
                        'income': warehouse_data['income'][i],
                        'outcome': warehouse_data['outcome'][i],
                        'end_balance': warehouse_data['end_balance'][i],
                        'reserved': warehouse_data['reserved'][i],
                        'in_transit': warehouse_data['in_transit'][i],
                        'price': warehouse_data['price'][i],
                        'total_sum': warehouse_data['total_sum'][i],
                        'last_income_date': warehouse_data['last_income_date'][i],
                        'last_move_date': warehouse_data['last_move_date'][i],
                        'description': warehouse_data['description'][i]
                    }
                    add_component_material(data)

                conn.close()
                materials = get_all_component_materials()

            # Добавляем данные в таблицу
            for material in materials:
                # Пропускаем id (material[0]) и updated_at (последний элемент)
                self.components_tree.insert("", tk.END, values=material[1:])

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {str(e)}")

    def create_film_tab(self):
        """Вкладка материалов участка наклейки пленки с расширенной таблицей"""
        self.film_tab = CTkFrame(self.notebook)
        self.notebook.add(self.film_tab, text="Материалы участка наклейки пленки")

        # Панель управления
        control_frame = CTkFrame(self.film_tab)
        control_frame.pack(fill=tk.X, pady=5)

        CTkButton(control_frame, text="Обновить", command=self.load_film_data).pack(side=tk.LEFT, padx=5)

        # Таблица с горизонтальной прокруткой
        container = CTkFrame(self.film_tab)
        container.pack(fill=tk.BOTH, expand=True)

        # Создаем Treeview с горизонтальной прокруткой
        self.film_tree = ttk.Treeview(container,
                                      columns=("num", "name", "unit", "photo",
                                               "start_balance", "income", "outcome", "end_balance",
                                               "reserved", "in_transit", "price", "total_sum",
                                               "last_income_date", "last_move_date", "description"),
                                      show="headings",
                                      height=20)

        # Настройка столбцов
        columns = [
            ("№", 30), ("Склад/ТМЦ/Документы движения", 250), ("Ед. изм.", 80),
            ("Фото", 50), ("Начальный остаток", 100), ("Приход", 80), ("Расход", 80),
            ("Конечный остаток", 100), ("в т.ч. резерв", 80), ("в пути", 80),
            ("Цена остатка", 100), ("Сумма остатка", 100), ("Дата последнего прихода", 120),
            ("Дата последнего перемещения", 120), ("Описание [назначение]", 100)
        ]

        for idx, (text, width) in enumerate(columns):
            self.film_tree.heading(f"#{idx + 1}", text=text)
            self.film_tree.column(f"#{idx + 1}", width=width, anchor='center')

        # Вертикальная прокрутка
        y_scroll = ttk.Scrollbar(container, orient="vertical", command=self.film_tree.yview)
        self.film_tree.configure(yscrollcommand=y_scroll.set)

        # Горизонтальная прокрутка
        x_scroll = ttk.Scrollbar(container, orient="horizontal", command=self.film_tree.xview)
        self.film_tree.configure(xscrollcommand=x_scroll.set)

        # Размещение элементов
        self.film_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

    def load_film_data(self):
        """Загрузка данных для участка наклейки пленки (только исходные столбцы)"""
        # Очищаем таблицу
        for item in self.film_tree.get_children():
            self.film_tree.delete(item)

        try:
            file_path = "warehouse_table/Ведомость материалы участок наклейки пленки.xls"
            if not os.path.exists(file_path):
                messagebox.showwarning("Внимание", "Файл с данными склада не найден. Загружаем из базы данных.")
                materials = get_all_film_materials()
            else:
                # Парсим данные из Excel
                warehouse_data = parse_excel_data(file_path)

                # Сохраняем данные в БД
                conn = sqlite3.connect('orders.db')
                cursor = conn.cursor()

                # Очищаем старые данные
                cursor.execute("DELETE FROM film_warehouse")
                conn.commit()

                # Вставляем новые данные
                num_records = len(warehouse_data['num'])
                for i in range(num_records):
                    data = {
                        'num': warehouse_data['num'][i],
                        'name': warehouse_data['name'][i],
                        'unit': warehouse_data['unit'][i],
                        'photo': warehouse_data['photo'][i],
                        'start_balance': warehouse_data['start_balance'][i],
                        'income': warehouse_data['income'][i],
                        'outcome': warehouse_data['outcome'][i],
                        'end_balance': warehouse_data['end_balance'][i],
                        'reserved': warehouse_data['reserved'][i],
                        'in_transit': warehouse_data['in_transit'][i],
                        'price': warehouse_data['price'][i],
                        'total_sum': warehouse_data['total_sum'][i],
                        'last_income_date': warehouse_data['last_income_date'][i],
                        'last_move_date': warehouse_data['last_move_date'][i],
                        'description': warehouse_data['description'][i]
                    }
                    add_film_material(data)

                conn.close()
                materials = get_all_film_materials()

            # Добавляем данные в таблицу
            for material in materials:
                self.film_tree.insert("", tk.END, values=material[1:])

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {str(e)}")


    def create_windows_tab(self):
        """Вкладка материалов участка стеклопакетов"""
        self.window_tab = CTkFrame(self.notebook)
        self.notebook.add(self.window_tab, text="Материалы участка стеклопакетов")

        # Панель управления
        control_frame = CTkFrame(self.window_tab)
        control_frame.pack(fill=tk.X, pady=5)

        CTkButton(control_frame, text="Обновить", command=self.load_windows_data).pack(side=tk.LEFT, padx=5)

        # Таблица с горизонтальной прокруткой
        container = CTkFrame(self.window_tab)
        container.pack(fill=tk.BOTH, expand=True)

        # Создаем Treeview с горизонтальной прокруткой
        self.window_tree = ttk.Treeview(container,
                                      columns=("num", "name", "unit", "photo",
                                               "start_balance", "income", "outcome", "end_balance",
                                               "reserved", "in_transit", "price", "total_sum",
                                               "last_income_date", "last_move_date", "description"),
                                      show="headings",
                                      height=20)

        # Настройка столбцов
        columns = [
            ("№", 30), ("Склад/ТМЦ/Документы движения", 250), ("Ед. изм.", 80),
            ("Фото", 50), ("Начальный остаток", 100), ("Приход", 80), ("Расход", 80),
            ("Конечный остаток", 100), ("в т.ч. резерв", 80), ("в пути", 80),
            ("Цена остатка", 100), ("Сумма остатка", 100), ("Дата последнего прихода", 120),
            ("Дата последнего перемещения", 120), ("Описание [назначение]", 100)
        ]

        for idx, (text, width) in enumerate(columns):
            self.window_tree.heading(f"#{idx + 1}", text=text)
            self.window_tree.column(f"#{idx + 1}", width=width, anchor='center')

        # Вертикальная прокрутка
        y_scroll = ttk.Scrollbar(container, orient="vertical", command=self.window_tree.yview)
        self.window_tree.configure(yscrollcommand=y_scroll.set)

        # Горизонтальная прокрутка
        x_scroll = ttk.Scrollbar(container, orient="horizontal", command=self.window_tree.xview)
        self.window_tree.configure(xscrollcommand=x_scroll.set)

        # Размещение элементов
        self.window_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

    def load_windows_data(self):
        """Загрузка данных склада комплектующих"""
        # Очищаем таблицу
        for item in self.window_tree.get_children():
            self.window_tree.delete(item)

        try:
            file_path = "warehouse_table/Ведомость материалы участок стеклопакетов.xls"
            if not os.path.exists(file_path):
                messagebox.showwarning("Внимание", "Файл с данными склада не найден. Загружаем из базы данных.")
                materials = get_all_window_materials()
            else:
                # Парсим данные из Excel
                warehouse_data = parse_excel_data(file_path)

                # Сохраняем данные в БД
                conn = sqlite3.connect('orders.db')
                cursor = conn.cursor()

                # Очищаем старые данные
                cursor.execute("DELETE FROM windows_warehouse")
                conn.commit()

                # Вставляем новые данные
                num_records = len(warehouse_data['num'])
                for i in range(num_records):
                    data = {
                        'num': warehouse_data['num'][i],
                        'name': warehouse_data['name'][i],
                        'unit': warehouse_data['unit'][i],
                        'photo': warehouse_data['photo'][i],
                        'start_balance': warehouse_data['start_balance'][i],
                        'income': warehouse_data['income'][i],
                        'outcome': warehouse_data['outcome'][i],
                        'end_balance': warehouse_data['end_balance'][i],
                        'reserved': warehouse_data['reserved'][i],
                        'in_transit': warehouse_data['in_transit'][i],
                        'price': warehouse_data['price'][i],
                        'total_sum': warehouse_data['total_sum'][i],
                        'last_income_date': warehouse_data['last_income_date'][i],
                        'last_move_date': warehouse_data['last_move_date'][i],
                        'description': warehouse_data['description'][i]
                    }
                    add_window_material(data)

                conn.close()
                materials = get_all_window_materials()

            # Добавляем данные в таблицу
            for material in materials:
                # Пропускаем id (material[0]) и updated_at (последний элемент)
                self.window_tree.insert("", tk.END, values=material[1:])

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {str(e)}")

    def create_triplex_tab(self):
        """Вкладка материалов участка триплекса"""
        self.triplex_tab = CTkFrame(self.notebook)
        self.notebook.add(self.triplex_tab, text="Материалы участка триплекса")

        # Панель управления
        control_frame = CTkFrame(self.triplex_tab)
        control_frame.pack(fill=tk.X, pady=5)

        CTkButton(control_frame, text="Обновить", command=self.load_triplex_data).pack(side=tk.LEFT, padx=5)

        # Таблица с горизонтальной прокруткой
        container = CTkFrame(self.triplex_tab)
        container.pack(fill=tk.BOTH, expand=True)

        # Создаем Treeview с горизонтальной прокруткой
        self.triplex_tree = ttk.Treeview(container,
                                        columns=("num", "name", "unit", "photo",
                                                 "start_balance", "income", "outcome", "end_balance",
                                                 "reserved", "in_transit", "price", "total_sum",
                                                 "last_income_date", "last_move_date", "description"),
                                        show="headings",
                                        height=20)

        # Настройка столбцов
        columns = [
            ("№", 30), ("Склад/ТМЦ/Документы движения", 250), ("Ед. изм.", 80),
            ("Фото", 50), ("Начальный остаток", 100), ("Приход", 80), ("Расход", 80),
            ("Конечный остаток", 100), ("в т.ч. резерв", 80), ("в пути", 80),
            ("Цена остатка", 100), ("Сумма остатка", 100), ("Дата последнего прихода", 120),
            ("Дата последнего перемещения", 120), ("Описание [назначение]", 100)
        ]

        for idx, (text, width) in enumerate(columns):
            self.triplex_tree.heading(f"#{idx + 1}", text=text)
            self.triplex_tree.column(f"#{idx + 1}", width=width, anchor='center')

        # Вертикальная прокрутка
        y_scroll = ttk.Scrollbar(container, orient="vertical", command=self.triplex_tree.yview)
        self.triplex_tree.configure(yscrollcommand=y_scroll.set)

        # Горизонтальная прокрутка
        x_scroll = ttk.Scrollbar(container, orient="horizontal", command=self.triplex_tree.xview)
        self.triplex_tree.configure(xscrollcommand=x_scroll.set)

        # Размещение элементов
        self.triplex_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

    def load_triplex_data(self):
        """Загрузка данных склада триплекса"""
        # Очищаем таблицу
        for item in self.triplex_tree.get_children():
            self.triplex_tree.delete(item)

        try:
            file_path = "warehouse_table/Ведомость материалы участок триплекса.xls"
            if not os.path.exists(file_path):
                messagebox.showwarning("Внимание", "Файл с данными склада не найден. Загружаем из базы данных.")
                materials = get_all_triplex_materials()
            else:
                # Парсим данные из Excel
                warehouse_data = parse_excel_data(file_path)

                # Сохраняем данные в БД
                conn = sqlite3.connect('orders.db')
                cursor = conn.cursor()

                # Очищаем старые данные
                cursor.execute("DELETE FROM triplex_warehouse")
                conn.commit()

                # Вставляем новые данные
                num_records = len(warehouse_data['num'])
                for i in range(num_records):
                    data = {
                        'num': warehouse_data['num'][i],
                        'name': warehouse_data['name'][i],
                        'unit': warehouse_data['unit'][i],
                        'photo': warehouse_data['photo'][i],
                        'start_balance': warehouse_data['start_balance'][i],
                        'income': warehouse_data['income'][i],
                        'outcome': warehouse_data['outcome'][i],
                        'end_balance': warehouse_data['end_balance'][i],
                        'reserved': warehouse_data['reserved'][i],
                        'in_transit': warehouse_data['in_transit'][i],
                        'price': warehouse_data['price'][i],
                        'total_sum': warehouse_data['total_sum'][i],
                        'last_income_date': warehouse_data['last_income_date'][i],
                        'last_move_date': warehouse_data['last_move_date'][i],
                        'description': warehouse_data['description'][i]
                    }
                    add_triplex_material(data)

                conn.close()
                materials = get_all_triplex_materials()

            # Добавляем данные в таблицу
            for material in materials:
                # Пропускаем id (material[0]) и updated_at (последний элемент)
                self.triplex_tree.insert("", tk.END, values=material[1:])

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {str(e)}")


    def create_main_glass_tab(self):
        """Вкладка основного склада стекла"""
        self.main_glass_tab = CTkFrame(self.notebook)
        self.notebook.add(self.main_glass_tab, text="Материалы основного склада стекла")

        # Панель управления
        control_frame = CTkFrame(self.main_glass_tab)
        control_frame.pack(fill=tk.X, pady=5)

        CTkButton(control_frame, text="Обновить", command=self.load_main_glass_data).pack(side=tk.LEFT, padx=5)

        # Таблица с горизонтальной прокруткой
        container = CTkFrame(self.main_glass_tab)
        container.pack(fill=tk.BOTH, expand=True)

        # Создаем Treeview с горизонтальной прокруткой
        self.main_glass_tree = ttk.Treeview(container,
                                         columns=("num", "name", "unit", "photo",
                                                  "start_balance", "income", "outcome", "end_balance",
                                                  "reserved", "in_transit", "price", "total_sum",
                                                  "last_income_date", "last_move_date", "description"),
                                         show="headings",
                                         height=20)

        # Настройка столбцов
        columns = [
            ("№", 30), ("Склад/ТМЦ/Документы движения", 250), ("Ед. изм.", 80),
            ("Фото", 50), ("Начальный остаток", 100), ("Приход", 80), ("Расход", 80),
            ("Конечный остаток", 100), ("в т.ч. резерв", 80), ("в пути", 80),
            ("Цена остатка", 100), ("Сумма остатка", 100), ("Дата последнего прихода", 120),
            ("Дата последнего перемещения", 120), ("Описание [назначение]", 100)
        ]

        for idx, (text, width) in enumerate(columns):
            self.main_glass_tree.heading(f"#{idx + 1}", text=text)
            self.main_glass_tree.column(f"#{idx + 1}", width=width, anchor='center')

        # Вертикальная прокрутка
        y_scroll = ttk.Scrollbar(container, orient="vertical", command=self.main_glass_tree.yview)
        self.main_glass_tree.configure(yscrollcommand=y_scroll.set)

        # Горизонтальная прокрутка
        x_scroll = ttk.Scrollbar(container, orient="horizontal", command=self.main_glass_tree.xview)
        self.main_glass_tree.configure(xscrollcommand=x_scroll.set)

        # Размещение элементов
        self.main_glass_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

    def load_main_glass_data(self):
        """Загрузка данных основного склада стекла"""
        # Очищаем таблицу
        for item in self.main_glass_tree.get_children():
            self.main_glass_tree.delete(item)

        try:
            file_path = "warehouse_table/Ведомость стекло основной склад.xls"
            if not os.path.exists(file_path):
                messagebox.showwarning("Внимание", "Файл с данными склада не найден. Загружаем из базы данных.")
                materials = get_all_main_glass_materials()
            else:
                # Парсим данные из Excel
                warehouse_data = parse_excel_data(file_path)

                # Сохраняем данные в БД
                conn = sqlite3.connect('orders.db')
                cursor = conn.cursor()

                # Очищаем старые данные
                cursor.execute("DELETE FROM main_glass_warehouse")
                conn.commit()

                # Вставляем новые данные
                num_records = len(warehouse_data['num'])
                for i in range(num_records):
                    data = {
                        'num': warehouse_data['num'][i],
                        'name': warehouse_data['name'][i],
                        'unit': warehouse_data['unit'][i],
                        'photo': warehouse_data['photo'][i],
                        'start_balance': warehouse_data['start_balance'][i],
                        'income': warehouse_data['income'][i],
                        'outcome': warehouse_data['outcome'][i],
                        'end_balance': warehouse_data['end_balance'][i],
                        'reserved': warehouse_data['reserved'][i],
                        'in_transit': warehouse_data['in_transit'][i],
                        'price': warehouse_data['price'][i],
                        'total_sum': warehouse_data['total_sum'][i],
                        'last_income_date': warehouse_data['last_income_date'][i],
                        'last_move_date': warehouse_data['last_move_date'][i],
                        'description': warehouse_data['description'][i]
                    }
                    add_main_glass_material(data)

                conn.close()
                materials = get_all_main_glass_materials()

            # Добавляем данные в таблицу
            for material in materials:
                self.main_glass_tree.insert("", tk.END, values=material[1:])

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {str(e)}")



    def create_cutting_tab(self):
        """Вкладка материалов участка резки"""
        self.cutting_tab = CTkFrame(self.notebook)
        self.notebook.add(self.cutting_tab, text="Стекло участок резки")

        # Панель управления
        control_frame = CTkFrame(self.cutting_tab)
        control_frame.pack(fill=tk.X, pady=5)

        CTkButton(control_frame, text="Обновить", command=self.load_cutting_data).pack(side=tk.LEFT, padx=5)

        # Таблица с горизонтальной прокруткой
        container = CTkFrame(self.cutting_tab)
        container.pack(fill=tk.BOTH, expand=True)

        # Создаем Treeview с горизонтальной прокруткой
        self.cutting_tree = ttk.Treeview(container,
                                            columns=("num", "name", "unit", "photo",
                                                     "start_balance", "income", "outcome", "end_balance",
                                                     "reserved", "in_transit", "price", "total_sum", "description"),
                                            show="headings",
                                            height=20)

        # Настройка столбцов
        columns = [
            ("№", 30), ("Склад/ТМЦ/Документы движения", 250), ("Ед. изм.", 80),
            ("Фото", 50), ("Начальный остаток", 100), ("Приход", 80), ("Расход", 80),
            ("Конечный остаток", 100), ("в т.ч. резерв", 80), ("в пути", 80),
            ("Цена остатка", 100), ("Сумма остатка", 100), ("Описание [назначение]", 100)
        ]

        for idx, (text, width) in enumerate(columns):
            self.cutting_tree.heading(f"#{idx + 1}", text=text)
            self.cutting_tree.column(f"#{idx + 1}", width=width, anchor='center')

        # Вертикальная прокрутка
        y_scroll = ttk.Scrollbar(container, orient="vertical", command=self.cutting_tree.yview)
        self.cutting_tree.configure(yscrollcommand=y_scroll.set)

        # Горизонтальная прокрутка
        x_scroll = ttk.Scrollbar(container, orient="horizontal", command=self.cutting_tree.xview)
        self.cutting_tree.configure(xscrollcommand=x_scroll.set)

        # Размещение элементов
        self.cutting_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

    def load_cutting_data(self):
        """Загрузка данных склада резки стекла"""
        # Очищаем таблицу
        for item in self.cutting_tree.get_children():
            self.cuttings_tree.delete(item)

        try:
            file_path = "warehouse_table/Ведомость стекло участок резки.xls"
            if not os.path.exists(file_path):
                messagebox.showwarning("Внимание", "Файл с данными склада не найден. Загружаем из базы данных.")
                materials = get_all_cutting_materials()
            else:
                # Парсим данные из Excel
                warehouse_data = parse_excel_data(file_path, True)

                # Сохраняем данные в БД
                conn = sqlite3.connect('orders.db')
                cursor = conn.cursor()

                # Очищаем старые данные
                cursor.execute("DELETE FROM cutting_warehouse")
                conn.commit()

                # Вставляем новые данные
                num_records = len(warehouse_data['num'])
                for i in range(num_records):
                    data = {
                        'num': warehouse_data['num'][i],
                        'name': warehouse_data['name'][i],
                        'unit': warehouse_data['unit'][i],
                        'photo': warehouse_data['photo'][i],
                        'start_balance': warehouse_data['start_balance'][i],
                        'income': warehouse_data['income'][i],
                        'outcome': warehouse_data['outcome'][i],
                        'end_balance': warehouse_data['end_balance'][i],
                        'reserved': warehouse_data['reserved'][i],
                        'in_transit': warehouse_data['in_transit'][i],
                        'price': warehouse_data['price'][i],
                        'total_sum': warehouse_data['total_sum'][i],
                        'description': warehouse_data['description'][i]
                    }
                    add_cutting_material(data)

                conn.close()
                materials = get_all_cutting_materials()

            # Добавляем данные в таблицу
            for material in materials:
                self.cutting_tree.insert("", tk.END, values=material[1:])

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить данные: {str(e)}")


    def load_all_data(self):
        """Загрузка данных для всех вкладок"""
        self.load_components_data()
        self.load_film_data()
        self.load_windows_data()
        self.load_triplex_data()
        self.load_main_glass_data()
        self.load_cutting_data()

