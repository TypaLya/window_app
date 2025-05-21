import tkinter as tk
from customtkinter import CTkLabel, CTkEntry, CTkButton, CTkFrame, CTkScrollbar
from tkinter import messagebox

from database import (
    get_production_orders,
    get_windows_for_production_order
)


class FrameCuttingTab(CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.groups = []

        # Настроим 3 строки:
        # 0 - шапка (фиксированная высота 80px)
        # 1 - холст карты раскроя (занимает максимум доступного места)
        # 2 - нижняя часть с двумя списками (примерно 40% от высоты окна)
        self.grid_rowconfigure(0, minsize=80)
        self.grid_rowconfigure(1, weight=6)
        self.grid_rowconfigure(2, weight=4)
        self.grid_columnconfigure(0, weight=1)

        # --- Верхняя шапка ---
        header_frame = CTkFrame(self)
        header_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10,5))

        # Делим шапку на 2 колонки (по 50%)
        header_frame.grid_columnconfigure(0, weight=1)
        header_frame.grid_columnconfigure(1, weight=1)

        # Ввод длины профиля слева, по центру
        self.length_entry = CTkEntry(header_frame, placeholder_text="2000 <= Длина профиля (мм) <= 6000")
        self.length_entry.grid(row=0, column=0, sticky="ew", padx=(0, 10), pady=10)
        header_frame.grid_rowconfigure(0, weight=1)

        # Кнопка запуска оптимизации справа, по центру
        self.optimize_button = CTkButton(header_frame, text="Запустить оптимизацию", command=self.optimize_cutting)
        self.optimize_button.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=10)

        # --- Холст для визуализации карты раскроя ---
        self.card_canvas = tk.Canvas(self, bg="white")
        self.card_canvas.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        # --- Нижняя часть с двумя списками ---
        bottom_frame = CTkFrame(self)
        bottom_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=(5, 10))

        # Две колонки по 50% ширины
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_columnconfigure(1, weight=1)
        bottom_frame.grid_rowconfigure(0, weight=1)

        # Список заказов слева
        orders_frame = CTkFrame(bottom_frame)
        orders_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        orders_frame.grid_rowconfigure(1, weight=1)
        orders_frame.grid_columnconfigure(0, weight=1)

        orders_label = CTkLabel(orders_frame, text="Список заказов")
        orders_label.grid(row=0, column=0, sticky="w")

        self.order_listbox = tk.Listbox(orders_frame, bg="#333333", fg="white", font=("Arial", 12))
        self.order_listbox.grid(row=1, column=0, sticky="nsew")

        scrollbar_orders = CTkScrollbar(orders_frame, command=self.order_listbox.yview)
        scrollbar_orders.grid(row=1, column=1, sticky="ns")
        self.order_listbox.config(yscrollcommand=scrollbar_orders.set)

        # Список карт раскроя справа
        cards_frame = CTkFrame(bottom_frame)
        cards_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        cards_frame.grid_rowconfigure(1, weight=1)
        cards_frame.grid_columnconfigure(0, weight=1)

        cards_label = CTkLabel(cards_frame, text="Карты раскроя")
        cards_label.grid(row=0, column=0, sticky="w")

        self.card_listbox = tk.Listbox(cards_frame, bg="#333333", fg="white", font=("Arial", 12))
        self.card_listbox.grid(row=1, column=0, sticky="nsew")

        scrollbar_cards = CTkScrollbar(cards_frame, command=self.card_listbox.yview)
        scrollbar_cards.grid(row=1, column=1, sticky="ns")
        self.card_listbox.config(yscrollcommand=scrollbar_cards.set)

        self.card_listbox.bind("<<ListboxSelect>>", self.display_card_details)

        self.load_orders_from_db()


    # Далее идут ваши методы (без изменений)...
    def update_orders_list(self):
        self.order_listbox.delete(0, tk.END)
        orders = get_production_orders()
        if not orders:
            return
        for order in orders:
            order_id = order[0]
            status = order[-1]
            if status.lower() == "завершен":
                continue
            windows = get_windows_for_production_order(order_id)
            for window in windows:
                window_id, window_type, width, height, quantity = window
                for _ in range(quantity):
                    order_text = f"Заказ {order_id}: {width}x{height} ({window_type})"
                    self.order_listbox.insert(tk.END, order_text)

    def load_orders_from_db(self):
        self.update_orders_list()

    def optimize_cutting(self):
        total_length_str = self.length_entry.get().strip()
        if total_length_str == "":
            total_length = 6000  # значение по умолчанию
        else:
            try:
                total_length = int(total_length_str)
                if not (2000 <= total_length <= 6000):
                    raise ValueError()
            except Exception:
                messagebox.showerror("Ошибка", "Введите корректную длину профиля от 2000 до 6000 мм (целое число).")
                return

        reserve = 21

        # Далее используем total_length в расчетах
        pieces = []

        for i in range(self.order_listbox.size()):
            order_text = self.order_listbox.get(i)
            try:
                order_id_str = order_text.split(" ")[1].rstrip(":")
                order_id = int(order_id_str)
                dims_part = order_text.split(": ")[1].split(" ")[0]
                width_str, height_str = dims_part.split("x")
                width = int(width_str)
                height = int(height_str)

                horizontal_length = width - reserve
                vertical_length = height - reserve

                if horizontal_length > 0:
                    pieces.append((order_id, horizontal_length))
                    pieces.append((order_id, horizontal_length))
                if vertical_length > 0:
                    pieces.append((order_id, vertical_length))
                    pieces.append((order_id, vertical_length))

            except Exception:
                continue

        pieces.sort(key=lambda x: x[1], reverse=True)

        groups = []
        groups_sum = []

        for piece in pieces:
            placed = False
            for i, used in enumerate(groups_sum):
                if used + piece[1] <= total_length:
                    groups[i].append(piece)
                    groups_sum[i] += piece[1]
                    placed = True
                    break
            if not placed:
                groups.append([piece])
                groups_sum.append(piece[1])

        self.groups = []
        for group in groups:
            d = {}
            for order_id, length in group:
                d.setdefault(order_id, []).append(length)
            self.groups.append(d)

        self.card_listbox.delete(0, tk.END)
        for i, group in enumerate(self.groups):
            summary = []
            for oid, lengths in group.items():
                counts = {}
                for l in lengths:
                    counts[l] = counts.get(l, 0) + 1
                parts_desc = [f"{cnt}×{lng}мм" for lng, cnt in counts.items()]
                summary.append(f"Заказ {oid}: " + ", ".join(parts_desc))
            text = f"Профиль {i + 1}: " + "; ".join(summary)
            self.card_listbox.insert(tk.END, text)

        if self.groups:
            self.draw_horizontal_cutting_plan(self.groups[0])

    def draw_horizontal_cutting_plan(self, group):
        canvas_width = self.card_canvas.winfo_width() or 800
        canvas_height = 100
        self.card_canvas.delete("all")

        total_length = int(self.length_entry.get()) if self.length_entry.get().isdigit() else 6000
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
                text=f"{remaining_length}", fill="black", font=("Arial", 10)
            )

    def display_card_details(self, event):
        selection = event.widget.curselection()
        if not selection:
            return
        index = selection[0]
        group = self.groups[index]
        self.draw_horizontal_cutting_plan(group)
