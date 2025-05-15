import tkinter as tk
from tkinter import ttk, filedialog
from datetime import datetime, timedelta

import openpyxl
import xlrd
from customtkinter import CTk, CTkLabel, CTkEntry, CTkButton, CTkFrame, CTkScrollbar, CTkRadioButton, CTkComboBox
from tkinter import messagebox
from database import (add_production_order, get_production_orders,
                     update_production_order_status, delete_production_order,
                     add_window_to_production_order, get_windows_for_production_order,
                     delete_window_from_production_order)


class ProductionPlanningTab(CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.current_order_id = None
        self._first_draw = True  # Флаг первого отображения

        # Основные фреймы
        self.left_frame = CTkFrame(self, width=300)
        self.left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10)

        self.right_frame = CTkFrame(self)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Форма добавления заказа
        self.add_order_frame = CTkFrame(self.left_frame)
        self.add_order_frame.pack(fill=tk.X, pady=5)

        CTkLabel(self.add_order_frame, text="Новый производственный заказ").pack(pady=5)

        # Поля ввода
        self.order_name_entry = CTkEntry(self.add_order_frame, placeholder_text="Название заказа")
        self.order_name_entry.pack(fill=tk.X, pady=5)

        self.customer_entry = CTkEntry(self.add_order_frame, placeholder_text="Заказчик")
        self.customer_entry.pack(fill=tk.X, pady=5)

        self.deadline_entry = CTkEntry(self.add_order_frame, placeholder_text="Срок (ДД.ММ.ГГГГ)")
        self.deadline_entry.pack(fill=tk.X, pady=5)

        self.priority_var = tk.StringVar(value="Средний")
        self.priority_combo = CTkComboBox(self.add_order_frame,
                                          values=["Низкий", "Средний", "Высокий", "Критичный"],
                                          variable=self.priority_var)
        self.priority_combo.pack(fill=tk.X, pady=5)

        self.add_button = CTkButton(self.add_order_frame, text="Добавить заказ", command=self.add_production_order)
        self.add_button.pack(fill=tk.X, pady=5)

        self.import_button = CTkButton(self.add_order_frame, text="Импорт заказа",
                                       command=self.import_order_from_excel)
        self.import_button.pack(fill=tk.X, pady=5)

        # Список заказов
        self.orders_frame = CTkFrame(self.left_frame)
        self.orders_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.orders_listbox = tk.Listbox(self.orders_frame, bg="#333333", fg="white", font=("Arial", 12))
        self.orders_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.orders_listbox.bind('<<ListboxSelect>>', self.show_order_details)

        self.scrollbar = CTkScrollbar(self.orders_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.orders_listbox.config(yscrollcommand=self.scrollbar.set)
        self.scrollbar.configure(command=self.orders_listbox.yview)
        self.load_production_orders()

        # Кнопки управления
        self.control_frame = CTkFrame(self.left_frame)
        self.control_frame.pack(fill=tk.X, pady=5)

        self.start_button = CTkButton(self.control_frame, text="В работу",
                                      command=lambda: self.change_order_status("В работе"))
        self.start_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        self.complete_button = CTkButton(self.control_frame, text="Завершить",
                                         command=lambda: self.change_order_status("Завершен"))
        self.complete_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        self.delete_button = CTkButton(self.control_frame, text="Удалить", command=self.delete_order)
        self.delete_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        # Детали заказа
        self.details_frame = CTkFrame(self.right_frame)
        self.details_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # Вкладки для деталей заказа
        self.details_notebook = ttk.Notebook(self.details_frame)
        self.details_notebook.pack(fill=tk.BOTH, expand=True)

        # Вкладка с информацией о заказе
        self.info_tab = CTkFrame(self.details_notebook)
        self.details_notebook.add(self.info_tab, text="Информация")

        self.order_details_text = tk.Text(self.info_tab, bg="#333333", fg="white", font=("Arial", 12))
        self.order_details_text.pack(fill=tk.BOTH, expand=True)

        # Вкладка со стеклопакетами
        self.windows_tab = CTkFrame(self.details_notebook)
        self.details_notebook.add(self.windows_tab, text="Стеклопакеты")

        # Кнопки управления стеклопакетами
        self.windows_control_frame = CTkFrame(self.windows_tab)
        self.windows_control_frame.pack(fill=tk.X, pady=5)

        self.add_window_button = CTkButton(self.windows_control_frame, text="Добавить стеклопакет",
                                           command=self.add_window_to_order)
        self.add_window_button.pack(side=tk.LEFT, padx=5)

        self.delete_window_button = CTkButton(self.windows_control_frame, text="Удалить стеклопакет",
                                              command=self.delete_window_from_order)
        self.delete_window_button.pack(side=tk.LEFT, padx=5)

        # Таблица стеклопакетов
        self.windows_tree = ttk.Treeview(self.windows_tab, columns=("id", "type", "width", "height", "quantity"), show="headings")
        self.windows_tree.heading("id", text="ID")
        self.windows_tree.heading("type", text="Тип")
        self.windows_tree.heading("width", text="Ширина (мм)")
        self.windows_tree.heading("height", text="Высота (мм)")
        self.windows_tree.heading("quantity", text="Количество")
        self.windows_tree.column("id", width=50)
        self.windows_tree.column("type", width=150)
        self.windows_tree.column("width", width=75)
        self.windows_tree.column("height", width=75)
        self.windows_tree.column("quantity", width=50)

        scrollbar = ttk.Scrollbar(self.windows_tab, orient="vertical", command=self.windows_tree.yview)
        self.windows_tree.configure(yscrollcommand=scrollbar.set)

        self.windows_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Календарь производства
        self.calendar_frame = CTkFrame(self.right_frame)
        self.calendar_frame.pack(fill=tk.BOTH, expand=True, pady=5)

       # Загрузка данных
        self.load_production_orders()

        # Загрузка данных и отрисовка календаря
        self.after(100, self.update_calendar)  # Запускаем через небольшой таймаут
        self.bind("<Visibility>", self.on_visibility_changed)
        self.bind("<Configure>", self.on_resize)  # Обработчик изменения размеров

        # Панель навигации по месяцам
        self.calendar_nav_frame = CTkFrame(self.calendar_frame)
        self.calendar_nav_frame.pack(fill=tk.X, pady=5)

        self.prev_month_btn = CTkButton(self.calendar_nav_frame, text="<", width=30,
                                        command=self.show_prev_month)
        self.prev_month_btn.pack(side=tk.LEFT, padx=5)

        self.month_label = CTkLabel(self.calendar_nav_frame, text="", font=("Arial", 12))
        self.month_label.pack(side=tk.LEFT, expand=True)

        self.next_month_btn = CTkButton(self.calendar_nav_frame, text=">", width=30,
                                        command=self.show_next_month)
        self.next_month_btn.pack(side=tk.RIGHT, padx=5)

        # Холст календаря
        self.calendar_canvas = tk.Canvas(self.calendar_frame, bg="white")
        self.calendar_canvas.pack(fill=tk.BOTH, expand=True)
        self.calendar_canvas.bind("<Button-1>", self.on_calendar_click)

        # Текущий месяц и год для отображения
        self.current_date = datetime.now().date()
        self.current_month = self.current_date.month
        self.current_year = self.current_date.year

        # Загрузка данных и отрисовка календаря
        self.update_calendar()

    def on_visibility_changed(self, event):
        """Обработчик изменения видимости вкладки"""
        if self._first_draw and self.winfo_ismapped():
            self._first_draw = False
            self.after(100, self.update_calendar)

    def on_resize(self, event):
        """Обработчик изменения размеров"""
        if self.winfo_ismapped():  # Проверяем, что вкладка видима
            self.update_calendar()

    def update_calendar(self):
        """Обновление отображения календаря"""
        if not self.winfo_ismapped():  # Не обновляем, если вкладка не видна
            return

        if not self.calendar_canvas.winfo_exists():
            return

        self.calendar_canvas.delete("all")
        self.month_label.configure(text=f"{self.current_month}.{self.current_year}")

        # Получаем все заказы
        orders = get_production_orders()

        # Определяем первый и последний день месяца
        first_day = datetime(self.current_year, self.current_month, 1).date()
        last_day = datetime(self.current_year, self.current_month + 1, 1).date() - timedelta(days=1)

        # Определяем дни недели для первого и последнего дня месяца
        start_weekday = first_day.weekday()  # 0-пн, 6-вс
        total_days = last_day.day

        # Настройки отображения
        canvas_width = self.calendar_canvas.winfo_width()
        canvas_height = self.calendar_canvas.winfo_height()
        if canvas_width < 10 or canvas_height < 10:  # Минимальные размеры
            return
        day_width = canvas_width / 7
        day_height = canvas_height / 6  # Максимум 6 недель в месяце

        # Рисуем заголовок с днями недели
        days = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"]
        for i, day in enumerate(days):
            x = i * day_width
            self.calendar_canvas.create_rectangle(x, 0, x + day_width, day_height, fill="#444444", outline="black")
            self.calendar_canvas.create_text(x + day_width / 2, day_height / 2, text=day, fill="white",
                                             font=("Arial", 10))

        # Рисуем сетку календаря
        current_day = 1
        for week in range(6):
            for day in range(7):
                if (week == 0 and day < start_weekday) or current_day > total_days:
                    continue

                x = day * day_width
                y = (week + 1) * day_height

                # Рисуем ячейку дня
                self.calendar_canvas.create_rectangle(
                    x, y, x + day_width, y + day_height,
                    fill="#555555", outline="black"
                )

                # Подпись числа
                self.calendar_canvas.create_text(
                    x + 5, y + 5,
                    text=str(current_day),
                    anchor="nw", fill="white", font=("Arial", 10)
                )

                # Проверяем заказы на эту дату
                current_date = datetime(self.current_year, self.current_month, current_day).date()
                day_orders = [o for o in orders
                              if datetime.strptime(o[3], "%Y-%m-%d").date() == current_date]

                if day_orders:
                    # Разделяем заказы на выполненные и невыполненные
                    completed_orders = [o for o in day_orders if o[5] == "Завершен"]
                    active_orders = [o for o in day_orders if o[5] != "Завершен"]

                    # Рисуем индикаторы
                    if completed_orders:
                        # Зеленый для выполненных
                        self.calendar_canvas.create_oval(
                            x + day_width - 15, y + 5,
                            x + day_width - 5, y + 15,
                            fill="#4CAF50", outline="black"
                        )
                        self.calendar_canvas.create_text(
                            x + day_width - 10, y + 10,
                            text=str(len(completed_orders)),
                            fill="white", font=("Arial", 8)
                        )

                    if active_orders:
                        # Красный для активных
                        self.calendar_canvas.create_oval(
                            x + day_width - 35, y + 5,
                            x + day_width - 25, y + 15,
                            fill="#F44336", outline="black"
                        )
                        self.calendar_canvas.create_text(
                            x + day_width - 30, y + 10,
                            text=str(len(active_orders)),
                            fill="white", font=("Arial", 8)
                        )

                current_day += 1

    def on_calendar_click(self, event):
        """Обработчик клика по календарю"""
        canvas_width = self.calendar_canvas.winfo_width()
        canvas_height = self.calendar_canvas.winfo_height()
        day_width = canvas_width / 7
        day_height = canvas_height / 6

        # Определяем координаты клика
        day_col = int(event.x // day_width)
        day_row = int(event.y // day_height)

        # Если клик в заголовке или вне дней месяца
        if day_row == 0 or day_col < 0 or day_col > 6:
            return

        # Получаем первый день месяца
        first_day = datetime(self.current_year, self.current_month, 1).date()
        start_weekday = first_day.weekday()  # 0-пн, 6-вс

        # Вычисляем день месяца
        day_num = (day_row - 1) * 7 + day_col - start_weekday + 1
        last_day = (datetime(self.current_year, self.current_month + 1, 1) - timedelta(days=1)).day

        if day_num < 1 or day_num > last_day:
            return

        # Получаем заказы на эту дату
        current_date = datetime(self.current_year, self.current_month, day_num).date()
        orders = get_production_orders()
        day_orders = [o for o in orders
                      if datetime.strptime(o[3], "%Y-%m-%d").date() == current_date]

        if not day_orders:
            messagebox.showinfo("Информация", f"На {current_date.strftime('%d.%m.%Y')} нет заказов")
            return

        # Создаем окно с информацией о заказах
        self.show_day_orders(day_orders, current_date)

    def show_day_orders(self, orders, date):
        """Показывает заказы на выбранную дату"""
        dialog = tk.Toplevel(self)
        dialog.title(f"Заказы на {date.strftime('%d.%m.%Y')}")
        dialog.geometry("500x400")

        # Список заказов
        orders_frame = CTkFrame(dialog)
        orders_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = CTkScrollbar(orders_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        orders_listbox = tk.Listbox(
            orders_frame,
            bg="#333333",
            fg="white",
            font=("Arial", 12),
            yscrollcommand=scrollbar.set
        )
        orders_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.configure(command=orders_listbox.yview)

        # Заполняем список
        for order in orders:
            order_id, name, customer, deadline, priority, status = order
            orders_listbox.insert(tk.END, f"{name} (Приоритет: {priority}, Статус: {status})")

        # Кнопка просмотра
        def view_order():
            selection = orders_listbox.curselection()
            if not selection:
                return

            selected_order = orders[selection[0]]
            self.show_order_details_by_id(selected_order[0])
            dialog.destroy()

        button_frame = CTkFrame(dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=5)

        CTkButton(button_frame, text="Просмотреть", command=view_order).pack(side=tk.LEFT, padx=5)
        CTkButton(button_frame, text="Закрыть", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)

    def show_order_details_by_id(self, order_id):
        """Показывает детали заказа по ID"""
        self.current_order_id = order_id
        orders = get_production_orders()

        for order in orders:
            if order[0] == order_id:
                order_id, name, customer, deadline, priority, status = order
                deadline_date = datetime.strptime(deadline, "%Y-%m-%d").strftime("%d.%m.%Y")

                details = f"Заказ №{order_id}\n\n"
                details += f"Название: {name}\n"
                details += f"Заказчик: {customer}\n"
                details += f"Срок выполнения: {deadline_date}\n"
                details += f"Приоритет: {priority}\n"
                details += f"Статус: {status}\n"

                self.order_details_text.config(state=tk.NORMAL)
                self.order_details_text.delete(1.0, tk.END)
                self.order_details_text.insert(tk.END, details)
                self.order_details_text.config(state=tk.DISABLED)

                # Загружаем стеклопакеты для этого заказа
                self.load_windows_for_order(order_id)
                break

    def show_prev_month(self):
        """Показывает предыдущий месяц"""
        self.current_month -= 1
        if self.current_month < 1:
            self.current_month = 12
            self.current_year -= 1
        self.update_calendar()

    def show_next_month(self):
        """Показывает следующий месяц"""
        self.current_month += 1
        if self.current_month > 12:
            self.current_month = 1
            self.current_year += 1
        self.update_calendar()

    def add_production_order(self):
        """Добавление нового производственного заказа"""
        name = self.order_name_entry.get()
        customer = self.customer_entry.get()
        deadline = self.deadline_entry.get()
        priority = self.priority_var.get()

        if not name or not deadline:
            messagebox.showerror("Ошибка", "Название и срок обязательны для заполнения")
            return

        try:
            deadline_date = datetime.strptime(deadline, "%d.%m.%y").date()
        except ValueError:
            try:
                deadline_date = datetime.strptime(deadline, "%d.%m.%Y").date()
            except ValueError:
                messagebox.showerror("Ошибка", "Неверный формат даты. Используйте ДД.ММ.ГГ или ДД.ММ.ГГГГ")
                return

        # Добавляем заказ в БД
        order_id = add_production_order(name, customer, deadline_date.strftime("%Y-%m-%d"), priority, "В ожидании")

        # Загружаем список заказов
        self.load_production_orders()

        # Устанавливаем текущий заказ
        self.current_order_id = order_id
        self.show_order_details_by_id(order_id)

        # Очищаем поля
        self.order_name_entry.delete(0, tk.END)
        self.customer_entry.delete(0, tk.END)
        self.deadline_entry.delete(0, tk.END)

        # Обновляем календарь
        self.update_calendar()

    def load_production_orders(self):
        """Загрузка списка производственных заказов"""
        orders = get_production_orders()
        self.orders_listbox.delete(0, tk.END)

        for order in orders:
            order_id, name, customer, deadline, priority, status = order
            deadline_date = datetime.strptime(deadline, "%Y-%m-%d").strftime("%d.%m.%Y")
            self.orders_listbox.insert(tk.END, f"{order_id}: {name} ({status})")

    def show_order_details(self, event=None):
        """Отображение деталей выбранного заказа"""
        selection = self.orders_listbox.curselection()
        if not selection:
            return

        # Получаем ID заказа из строки (формат "ID: Название")
        order_text = self.orders_listbox.get(selection[0])
        order_id = order_text.split(":")[0].strip()

        self.current_order_id = order_id
        orders = get_production_orders()

        for order in orders:
            if str(order[0]) == order_id:
                order_id, name, customer, deadline, priority, status = order
                deadline_date = datetime.strptime(deadline, "%Y-%m-%d").strftime("%d.%m.%Y")

                details = f"Заказ №{order_id}\n\n"
                details += f"Название: {name}\n"
                details += f"Заказчик: {customer}\n"
                details += f"Срок выполнения: {deadline_date}\n"
                details += f"Приоритет: {priority}\n"
                details += f"Статус: {status}\n"

                # Обновляем текстовое поле
                self.order_details_text.config(state=tk.NORMAL)
                self.order_details_text.delete(1.0, tk.END)
                self.order_details_text.insert(tk.END, details)
                self.order_details_text.config(state=tk.DISABLED)

                # Загружаем связанные стеклопакеты
                self.load_windows_for_order(order_id)
                break

    def load_windows_for_order(self, order_id):
        """Загрузка списка стеклопакетов для заказа"""
        # Очищаем таблицу
        for item in self.windows_tree.get_children():
            self.windows_tree.delete(item)

        # Загружаем данные из БД
        windows = get_windows_for_production_order(order_id)

        # Добавляем с порядковым номером
        for i, window in enumerate(windows, 1):
            self.windows_tree.insert("", tk.END, values=(i, *window[1:]))  # Пропускаем id и order_id

    def add_window_to_order(self):
        """Добавление стеклопакета к заказу"""
        if not self.current_order_id:
            messagebox.showwarning("Предупреждение", "Сначала выберите заказ")
            return

        # Создаем диалоговое окно
        self.window_dialog = tk.Toplevel(self)
        self.window_dialog.title("Добавить стеклопакет")
        self.window_dialog.geometry("300x250")
        self.window_dialog.protocol("WM_DELETE_WINDOW", self.close_window_dialog)

        # Переменные для хранения значений
        self.width_var = tk.StringVar()
        self.height_var = tk.StringVar()
        self.quantity_var = tk.StringVar()
        self.type_var = tk.StringVar()

        # Элементы управления
        CTkLabel(self.window_dialog, text="Ширина (мм):").pack(pady=5)
        width_entry = CTkEntry(self.window_dialog, textvariable=self.width_var)
        width_entry.pack(pady=5)
        width_entry.focus_set()

        CTkLabel(self.window_dialog, text="Высота (мм):").pack(pady=5)
        height_entry = CTkEntry(self.window_dialog, textvariable=self.height_var)
        height_entry.pack(pady=5)

        CTkLabel(self.window_dialog, text="Количество:").pack(pady=5)
        quantity_entry = CTkEntry(self.window_dialog, textvariable=self.quantity_var)
        quantity_entry.pack(pady=5)

        CTkLabel(self.window_dialog, text="Тип стеклопакета:").pack(pady=5)
        type_entry = CTkEntry(self.window_dialog, textvariable=self.type_var)
        type_entry.pack(pady=5)


        # Кнопки
        button_frame = CTkFrame(self.window_dialog)
        button_frame.pack(pady=10)

        CTkButton(button_frame, text="Сохранить и добавить еще",
                  command=self.save_window_and_continue).pack(side=tk.LEFT, padx=5)
        CTkButton(button_frame, text="Сохранить и закрыть",
                  command=self.save_window_and_close).pack(side=tk.LEFT, padx=5)
        CTkButton(button_frame, text="Отмена",
                  command=self.close_window_dialog).pack(side=tk.LEFT, padx=5)

    def save_window_and_continue(self):
        """Сохраняет текущий стеклопакет и очищает поля для ввода следующего"""
        try:
            width = int(self.width_var.get())
            height = int(self.height_var.get())
            quantity = int(self.quantity_var.get())
            window_type = self.type_var.get()

            if width <= 0 or height <= 0:
                messagebox.showerror("Ошибка", "Размеры должны быть положительными числами")
                return

            if quantity <= 0:
                messagebox.showerror("Ошибка", "Количество должно быть положительным")
                return

            add_window_to_production_order(self.current_order_id, window_type, width, height, quantity)
            self.load_windows_for_order(self.current_order_id)

            # Очищаем поля для ввода следующего значения
            self.width_var.set("")
            self.height_var.set("")
            self.quantity_var.set("")
            self.window_dialog.focus_set()  # Возвращаем фокус в окно

        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректные числовые значения")

    def save_window_and_close(self):
        """Сохраняет стеклопакет и закрывает окно"""
        self.save_window_and_continue()
        self.close_window_dialog()

    def close_window_dialog(self):
        """Закрывает окно добавления стеклопакетов"""
        if hasattr(self, 'window_dialog') and self.window_dialog:
            self.window_dialog.destroy()

    def delete_window_from_order(self):
        """Удаление стеклопакета из заказа"""
        if not self.current_order_id:
            messagebox.showwarning("Предупреждение", "Сначала выберите заказ")
            return

        selection = self.windows_tree.selection()
        if not selection:
            messagebox.showwarning("Предупреждение", "Выберите стеклопакет для удаления")
            return

        if messagebox.askyesno("Подтверждение", "Удалить выбранный стеклопакет?"):
            window_id = self.windows_tree.item(selection[0], "values")[0]
            delete_window_from_production_order(window_id)
            self.load_windows_for_order(self.current_order_id)

    def change_order_status(self, new_status):
        """Изменение статуса заказа"""
        if not self.current_order_id:
            messagebox.showwarning("Предупреждение", "Выберите заказ")
            return

        update_production_order_status(self.current_order_id, new_status)
        self.load_production_orders()
        self.show_order_details(None)
        self.update_calendar()

    def delete_order(self):
        """Удаление производственного заказа"""
        if not self.current_order_id:
            messagebox.showwarning("Предупреждение", "Выберите заказ")
            return

        if messagebox.askyesno("Подтверждение", "Удалить выбранный заказ?"):
            delete_production_order(self.current_order_id)
            self.load_production_orders()
            self.order_details_text.config(state=tk.NORMAL)
            self.order_details_text.delete(1.0, tk.END)
            self.order_details_text.config(state=tk.DISABLED)
            self.current_order_id = None

            # Очищаем таблицу стеклопакетов
            for item in self.windows_tree.get_children():
                self.windows_tree.delete(item)
        self.update_calendar()

    def import_order_from_excel(self):
        """Импорт заказа из Excel (всегда создаёт новый заказ)"""
        file_path = filedialog.askopenfilename(
            title="Выберите файл заказа",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            # 1. Парсим Excel (получаем данные заказа и окон)
            order_data = self.parse_excel_order(file_path)

            # 2. Создаём НОВЫЙ заказ (игнорируем self.current_order_id)
            deadline = datetime.strptime(order_data['deadline'], "%d.%m.%Y").strftime("%Y-%m-%d")
            new_order_id = add_production_order(
                name=order_data['order_name'],
                customer=order_data['customer'],
                deadline=deadline,
                priority="Средний",
                status="В ожидании"
            )

            # 3. Добавляем все стеклопакеты из файла в новый заказ
            for window in order_data['windows']:
                add_window_to_production_order(
                    order_id=new_order_id,  # Привязываем к новому заказу
                    window_type=window['type'],
                    width=window['width'],
                    height=window['height'],
                    quantity=window['quantity']
                )

            # 4. Обновляем интерфейс
            self.load_production_orders()
            self.current_order_id = new_order_id  # Выбираем новый заказ
            self.show_order_details_by_id(new_order_id)
            messagebox.showinfo("Успех", "Новый заказ создан и данные импортированы!")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка импорта:\n{str(e)}")
            print(f"DEBUG: {e}")  # Для диагностики

    def parse_excel_order(self, file_path):
        """Парсинг данных заказа из Excel файла (поддержка .xls и .xlsx)"""
        try:
            # Определяем расширение файла
            ext = file_path.split('.')[-1].lower()

            order_data = {
                'order_name': '',
                'customer': '',
                'deadline': '',
                'windows': []
            }

            if ext == 'xls':
                # Используем xlrd для старых .xls файлов
                import xlrd
                book = xlrd.open_workbook(file_path)
                sheet = book.sheet_by_index(0)

                # Парсим основные данные заказа
                for row_idx in range(sheet.nrows):
                    row = sheet.row_values(row_idx)
                    # print("row ", row)

                    # Номер заказа (ищем строку, содержащую "Заказ №")
                    if not order_data['order_name'] and row and len(row) > 1:
                        cell_value = str(row[1]).strip()
                        if "Заказ №" in cell_value:
                            order_data['order_name'] = cell_value
                            print(cell_value)

                    # Заказчик (ищем строку, где первый столбец содержит "Заказчик:" или "Çàêàç÷èê:")
                    if not order_data['customer'] and row and len(row) > 1:
                        cell_value = str(row[1]).strip()
                        if "Заказчик:" in cell_value:
                            customer = str(row[8]).strip() if len(row) > 8 and row[8] else ""
                            # Очищаем от лишней информации (телефонов и т.д.)
                            customer = customer.split(',')[0].split('тел.:')[0].strip()
                            order_data['customer'] = customer
                            print(customer)

                    # Дата доставки (ищем строку, где первый столбец содержит "Дата доставки:" или "Äàòà äîñòàâêè:")
                    if not order_data['deadline'] and row and len(row) > 1:
                        cell_value = str(row[17]).strip()
                        if "Дата изготовления:" in cell_value:
                            if len(row) > 17 and row[24]:
                                date_str = str(row[24]).split('\\')[0].strip()
                                # Пропускаем некорректные даты (например, ". .")
                                if date_str and date_str != '. .':
                                    # Преобразуем дату из формата "09.04.25" в "09.04.2025"
                                    if len(date_str.split('.')) == 3 and len(date_str.split('.')[2]) == 2:
                                        day, month, year = date_str.split('.')
                                        date_str = f"{day}.{month}.20{year}"
                                    order_data['deadline'] = date_str
                                    print(date_str)

                # Парсим список стеклопакетов (начинаем с строки, где есть заголовок "№")
                start_parsing = False
                one_skip = False
                for row_idx in range(sheet.nrows):
                    if one_skip:
                        one_skip = False
                        continue
                    row = sheet.row_values(row_idx)
                    # print("row steklo ", row)
                    # Ищем начало таблицы со стеклопакетами
                    if row and len(row) > 1:
                        first_cell = str(row[1]).strip()
                        if first_cell in ["№"] and "Поз" in str(row[3]):
                            start_parsing = True
                            one_skip = True
                            continue

                    if start_parsing and row and len(row) > 14:
                        try:
                            # Проверяем, что первая ячейка содержит число (номер строки)
                            if row[1] and row[5]:
                                # Обрабатываем размеры (удаляем пробелы и разделяем на ширину/высоту)
                                size_str = str(row[15]).replace(' ', '') if len(row) > 15 and row[15] else "0x0"
                                size_parts = size_str.split('x')

                                # Получаем количество (может быть в разных столбцах)
                                quantity = 1
                                if len(row) > 20 and row[20]:
                                    try:
                                        quantity = int(float(row[20]))
                                    except:
                                        quantity = 1

                                window_data = {
                                    'type': str(row[5]).strip(),
                                    'width': int(float(size_parts[0])) if size_parts[0] else 0,
                                    'height': int(float(size_parts[1])) if len(size_parts) > 1 and size_parts[1] else 0,
                                    'quantity': quantity
                                }
                                order_data['windows'].append(window_data)
                                # print(window_data, "     dsfsd")
                                # for key, value in window_data.items():
                                #     print(f"{key}: {value}")
                            else:
                                start_parsing = False
                        except (ValueError, IndexError, AttributeError) as e:
                            print(f"Ошибка при парсинге строки {row_idx}: {e}")
                            continue
                print(order_data['windows'])


            elif ext == 'xlsx':
                # Используем openpyxl для новых .xlsx файлов
                import openpyxl
                wb = openpyxl.load_workbook(file_path)
                sheet = wb.active

                # Парсим основные данные заказа
                for row in sheet.iter_rows(values_only=True):
                    # Номер заказа
                    if not order_data['order_name'] and row and len(row) > 1:
                        cell_value = str(row[1]).strip()
                        if "Заказ №" in cell_value:
                            order_data['order_name'] = cell_value

                    # Заказчик
                    if not order_data['customer'] and row and len(row) > 1:
                        cell_value = str(row[1]).strip()
                        if "Заказчик:" in cell_value:
                            customer = str(row[7]).strip() if len(row) > 7 and row[7] else ""
                            customer = customer.split(',')[0].split('тел.:')[0].strip()
                            order_data['customer'] = customer

                    # Дата доставки
                    if not order_data['deadline'] and row and len(row) > 1:
                        cell_value = str(row[1]).strip()
                        if "Дата доставки:" in cell_value:
                            if len(row) > 7 and row[7]:
                                date_str = str(row[7]).split('\\')[0].strip()
                                if date_str and date_str != '. .':
                                    if len(date_str.split('.')) == 3 and len(date_str.split('.')[2]) == 2:
                                        day, month, year = date_str.split('.')
                                        date_str = f"{day}.{month}.20{year}"
                                    order_data['deadline'] = date_str

                # Парсим список стеклопакетов
                start_parsing = False
                for row in sheet.iter_rows(values_only=True):
                    # Ищем начало таблицы со стеклопакетами
                    if row and len(row) > 1:
                        first_cell = str(row[1]).strip()
                        if first_cell in ["№", "¹"] and "Поз" in str(row[3]):
                            start_parsing = True
                            continue

                    if start_parsing and row and len(row) > 14:
                        try:
                            if isinstance(row[0], (int, float)) and row[4] and any(
                                    x in str(row[4]) for x in ["СПД", "СПО", "ÑÏÄ", "ÑÏÎ"]):
                                size_str = str(row[14]).replace(' ', '') if len(row) > 14 and row[14] else "0x0"
                                size_parts = size_str.split('x')

                                quantity = 1
                                if len(row) > 19 and row[19]:
                                    try:
                                        quantity = int(float(row[19]))
                                    except:
                                        quantity = 1

                                window_data = {
                                    'type': str(row[4]).strip(),
                                    'width': int(float(size_parts[0])) if size_parts[0] else 0,
                                    'height': int(float(size_parts[1])) if len(size_parts) > 1 and size_parts[1] else 0,
                                    'quantity': quantity
                                }
                                order_data['windows'].append(window_data)
                        except (ValueError, IndexError, AttributeError) as e:
                            print(f"Ошибка при парсинге строки: {e}")
                            continue
            else:
                raise Exception("Неподдерживаемый формат файла. Используйте .xls или .xlsx")

            return order_data

        except Exception as e:
            raise Exception(f"Ошибка при чтении файла: {str(e)}")