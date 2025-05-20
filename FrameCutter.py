import tkinter as tk

from customtkinter import CTkLabel, CTkEntry, CTkButton, CTkFrame, CTkScrollbar, CTkRadioButton
from tkinter import messagebox

from Item import NumberItem
from GroupSolver import GroupSolver

from database import (add_order_to_db, delete_order_from_db, update_order_in_db,
                      get_all_orders_from_db)

class FrameCuttingTab(CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.groups = []  # Список для хранения карт раскроя

        # Левый фрейм для ввода заказов
        self.left_frame = CTkFrame(self, width=200)
        self.left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10, expand=False)

        self.label_width = CTkLabel(self.left_frame, text="Ширина стеклопакета (мм):")
        self.label_width.pack(pady=5)
        self.entry_width = CTkEntry(self.left_frame, width=100)
        self.entry_width.pack(pady=5)

        self.label_height = CTkLabel(self.left_frame, text="Высота стеклопакета (мм):")
        self.label_height.pack(pady=5)
        self.entry_height = CTkEntry(self.left_frame, width=100)
        self.entry_height.pack(pady=5)

        self.label_type = CTkLabel(self.left_frame, text="Тип стеклопакета:")
        self.label_type.pack(pady=5)

        self.package_type = tk.StringVar(value="Однокамерный")

        self.radio_single = CTkRadioButton(self.left_frame, text="Однокамерный", variable=self.package_type,
                                           value="Однокамерный")
        self.radio_single.pack(anchor='w', padx=10)
        self.radio_double = CTkRadioButton(self.left_frame, text="Двухкамерный", variable=self.package_type,
                                           value="Двухкамерный")
        self.radio_double.pack(anchor='w', padx=10)

        self.add_order_button = CTkButton(self.left_frame, text="Добавить заказ", command=self.add_order)
        self.add_order_button.pack(pady=10)

        self.delete_order_button = CTkButton(self.left_frame, text="Удалить заказ", command=self.delete_order)
        self.delete_order_button.pack(pady=10)

        self.update_order_button = CTkButton(self.left_frame, text="Изменить заказ", command=self.update_order)
        self.update_order_button.pack(pady=10)

        self.order_list_frame = CTkFrame(self.left_frame)
        self.order_list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.scrollbar = CTkScrollbar(self.order_list_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.order_listbox = tk.Listbox(self.order_list_frame, yscrollcommand=self.scrollbar.set, height=15,
                                        bg="#333333", fg="white", font=("Arial", 12))
        self.order_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.configure(command=self.order_listbox.yview)

        # Правый фрейм для результатов оптимизации
        self.right_frame = CTkFrame(self, width=200)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10, expand=False)

        self.card_list_frame = CTkFrame(self.right_frame)
        self.card_list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.card_listbox = tk.Listbox(self.card_list_frame, height=15, bg="#333333", fg="white", font=("Arial", 12))
        self.card_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.scrollbar_cards = CTkScrollbar(self.card_list_frame)
        self.scrollbar_cards.pack(side=tk.RIGHT, fill=tk.Y)
        self.scrollbar_cards.configure(command=self.card_listbox.yview)

        # Центральная область
        self.center_frame = CTkFrame(self)
        self.center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.optimize_button = CTkButton(self.center_frame, text="Запустить оптимизацию", command=self.optimize_cutting,
                                         width=200)
        self.optimize_button.pack(pady=5, anchor='n')

        self.card_canvas = tk.Canvas(self.center_frame, width=800, height=100, bg="white")
        self.card_canvas.pack(pady=10)

        self.unused_label = CTkLabel(self.center_frame, text=f"")
        self.unused_label.pack(anchor='w', padx=10, pady=5)

        # Загружаем заказы из базы данных
        self.load_orders_from_db()

    def add_order(self):
        width = self.entry_width.get()
        height = self.entry_height.get()
        if width and height:
            try:
                width = int(width)
                height = int(height)
                if width <= 0 or height <= 0:
                    messagebox.showerror("Ошибка", "Ширина и высота должны быть положительными числами.")
                    return

                # Добавляем заказ в базу данных
                add_order_to_db(width, height, self.package_type.get())

                # Обновляем список заказов
                self.load_orders_from_db()

                # Оповещаем родительское окно об изменении данных
                self.parent.on_orders_updated()

            except ValueError:
                messagebox.showerror("Ошибка", "Неверный ввод. Пожалуйста, введите целые числа.")

    def delete_order(self):
        selected_order_index = self.order_listbox.curselection()
        if selected_order_index:
            # Получаем ID выбранного заказа
            orders = get_all_orders_from_db()
            order_id = orders[selected_order_index[0]][0]

            delete_order_from_db(order_id)
            self.load_orders_from_db()
            self.parent.on_orders_updated()
        else:
            messagebox.showwarning("Предупреждение", "Пожалуйста, выберите заказ для удаления.")

    def update_order(self):
        selected_order_index = self.order_listbox.curselection()
        if selected_order_index:
            # Получаем выбранный заказ
            orders = get_all_orders_from_db()
            order_id = orders[selected_order_index[0]][0]

            new_width = self.entry_width.get()
            new_height = self.entry_height.get()

            if new_width and new_height:
                try:
                    new_width = int(new_width)
                    new_height = int(new_height)

                    update_order_in_db(order_id, new_width, new_height)
                    self.load_orders_from_db()
                    self.parent.on_orders_updated()

                    self.entry_width.delete(0, tk.END)
                    self.entry_height.delete(0, tk.END)

                except ValueError:
                    messagebox.showerror("Ошибка", "Неверный ввод. Пожалуйста, введите целые числа.")
            else:
                messagebox.showwarning("Предупреждение", "Пожалуйста, введите новые значения.")
        else:
            messagebox.showwarning("Предупреждение", "Пожалуйста, выберите заказ для изменения.")

    def load_orders_from_db(self):
        orders = get_all_orders_from_db()
        self.order_listbox.delete(0, tk.END)
        for order in orders:
            order_id, width, height, package_type = order
            order_text = f"Заказ {order_id}: {width}x{height} ({package_type})"
            self.order_listbox.insert(tk.END, order_text)

    def optimize_cutting(self):
        orders = get_all_orders_from_db()
        items = []

        for order in orders:
            order_id, width, height, package_type = order
            if package_type == "Однокамерный":
                values = [width, height, width, height]
            else:
                values = [width, height, width, height, width, height, width, height]

            items.append(NumberItem(index=order_id, values=values))

        solver = GroupSolver(items)
        groups, unused_values = solver.group_numbers()

        self.groups = groups
        self.card_listbox.delete(0, tk.END)

        for idx, group in enumerate(groups):
            group_sum = sum(value for values in group.values() for value in values)
            group_text = f"Карта {idx + 1}: Сумма = {group_sum}"
            self.card_listbox.insert(tk.END, group_text)

        if unused_values:
            self.unused_label.configure(text=f"Не задействованные значения: {unused_values}")

    def draw_horizontal_cutting_plan(self, group):
        canvas_width = 800
        canvas_height = 100
        self.card_canvas.delete("all")

        total_length = 6000
        x_start = 0
        used_length = 0

        for order_id, lengths in group.items():
            for length in lengths:
                rect_width = (length / total_length) * canvas_width
                x_end = x_start + rect_width

                y_top = (canvas_height - 30) / 2
                y_bottom = (canvas_height + 30) / 2
                self.card_canvas.create_rectangle(
                    x_start, y_top, x_end, y_bottom, fill="green", outline="black"
                )

                self.card_canvas.create_text(
                    (x_start + x_end) / 2, (y_top + y_bottom) / 2,
                    text=f"{length}", fill="black", font=("Arial", 10)
                )

                self.card_canvas.create_text(
                    (x_start + x_end) / 2, canvas_height - 10,
                    text=f"{order_id}", fill="blue", font=("Arial", 10)
                )

                x_start = x_end
                used_length += length

        remaining_length = total_length - used_length
        if remaining_length > 0:
            rect_width = max((remaining_length / total_length) * canvas_width, 1)
            x_end = x_start + rect_width

            y_top = (canvas_height - 30) / 2
            y_bottom = (canvas_height + 30) / 2

            self.card_canvas.create_rectangle(
                x_start, y_top, x_end, y_bottom, fill="red", outline="black"
            )

            self.card_canvas.create_text(
                (x_start + x_end) / 2, (y_top + y_bottom) / 2,
                text=f"{remaining_length} мм", fill="white", font=("Arial", 10)
            )

    def display_card_details(self, event):
        selected_index = self.card_listbox.curselection()
        if selected_index:
            group = self.groups[selected_index[0]]
            self.draw_horizontal_cutting_plan(group)
