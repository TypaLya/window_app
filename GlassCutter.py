import hashlib
import time
import tkinter as tk
import rectpack
from tkinter import ttk
from typing import Dict, List, Tuple

from functools import lru_cache
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import itertools
from collections import defaultdict

from customtkinter import CTkLabel, CTkEntry, CTkButton, CTkFrame, CTkScrollbar, CTkProgressBar, CTkComboBox
from tkinter import messagebox

from database import (get_windows_for_production_order, get_production_orders)


class GlassCuttingTab(CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.bind("<Visibility>", lambda e: self.update_orders_list())
        self.min_rotation_angle = 90
        self.parent = parent
        self.groups = []
        self.sheet_width = 6000
        self.sheet_height = 6000

        # Добавляем минимальный и максимальный масштаб
        self.zoom_level_prev = 1.0  # Предыдущий уровень масштабирования
        self.current_zoom = 1.0  # Текущий уровень масштабирования (1.0 = 100%)
        self.zoom_factor = 1.1  # Коэффициент масштабирования при прокрутке
        self.min_zoom = 0.5  # Минимальный масштаб
        self.max_zoom = 3.0  # Максимальный масштаб
        self.canvas_offset_x = 0
        self.canvas_offset_y = 0
        self.last_mouse_pos = (0, 0)  # Для запоминания позиции курсора

        self.optimization_mode = "normal"  # "normal" или "deep"
        self._is_running = True
        self._gui_update_queue = []
        self.after(100, self._process_gui_updates)  # Запускаем обработчик обновлений GUI

        self.packing_cache = {}
        self.combination_cache = {}

        self.selected_item = None
        self.hover_item = None
        self.selection_rect = None
        self.hover_rect = None
        self.tooltip = None
        self.side_panels_width = 200

        # Левый фрейм для работы с заказами
        self.left_frame = CTkFrame(self, width=self.side_panels_width)
        self.left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10, expand=False)

        # Разделитель для левой панели
        self.left_separator = ttk.Separator(self, orient="vertical")
        self.left_separator.pack(side=tk.LEFT, fill="y", padx=2)
        self.left_separator.bind("<B1-Motion>", self.resize_left_panel)
        self.left_separator.bind("<Button-3>", self.show_panel_context_menu)

        # Добавляем поля для ввода размеров листа стекла
        self.label_sheet_width = CTkLabel(self.left_frame, text="Ширина листа стекла (мм):")
        self.label_sheet_width.pack(pady=5)
        self.entry_sheet_width = CTkEntry(self.left_frame, width=100)
        self.entry_sheet_width.insert(0, "6000")
        self.entry_sheet_width.pack(pady=5)

        self.label_sheet_height = CTkLabel(self.left_frame, text="Высота листа стекла (мм):")
        self.label_sheet_height.pack(pady=5)
        self.entry_sheet_height = CTkEntry(self.left_frame, width=100)
        self.entry_sheet_height.insert(0, "6000")
        self.entry_sheet_height.pack(pady=5)

        # Правый фрейм для результатов
        self.right_frame = CTkFrame(self, width=self.side_panels_width)
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10, expand=False)

        # Разделитель для правой панели
        self.right_separator = ttk.Separator(self, orient="vertical")
        self.right_separator.pack(side=tk.RIGHT, fill="y", padx=2)
        self.right_separator.bind("<B1-Motion>", self.resize_right_panel)
        self.right_separator.bind("<Button-3>", self.show_panel_context_menu)

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

        # Создаем фрейм для элементов управления оптимизацией
        self.optimization_control_frame = CTkFrame(self.center_frame)
        self.optimization_control_frame.pack(pady=5, anchor='n', fill=tk.X)

        # Кнопка оптимизации
        self.optimize_button = CTkButton(
            self.optimization_control_frame,
            text="Запустить оптимизацию",
            command=self.optimize_cutting,
            width=200
        )
        self.optimize_button.pack(side=tk.LEFT, padx=(0, 10))

        # Выбор режима оптимизации
        self.optimization_combobox = CTkComboBox(
            self.optimization_control_frame,
            values=["Быстрый (по умолчанию)", "Глубокий перебор"],
            command=self.set_optimization_mode,
            width=180
        )
        self.optimization_combobox.pack(side=tk.LEFT)
        self.optimization_combobox.set("Быстрый (по умолчанию)")

        # Добавляем поле для порога заполненности
        self.threshold_frame = CTkFrame(self.optimization_control_frame)
        self.threshold_frame.pack(side=tk.LEFT, padx=(10, 0))

        self.threshold_label = CTkLabel(self.threshold_frame, text="Порог заполнения (%):")
        self.threshold_label.pack(side=tk.LEFT, padx=(0, 5))

        self.threshold_entry = CTkEntry(self.threshold_frame, width=50)
        self.threshold_entry.insert(0, "90")  # Значение по умолчанию
        self.threshold_entry.pack(side=tk.LEFT)

        # Прогресс-бар и статус
        self.progress_bar = CTkProgressBar(self.center_frame, mode='determinate')
        self.progress_bar.pack(pady=5, fill=tk.X, padx=20)
        self.progress_bar.pack_forget()

        self.status_label = CTkLabel(self.center_frame, text="")
        self.status_label.pack(pady=5)
        self.status_label.pack_forget()

        # Холст для отображения
        self.card_canvas = tk.Canvas(
            self.center_frame,
            width=600,
            height=600,
            bg="white",
            highlightthickness=0
        )
        self.card_canvas.pack(pady=10, fill=tk.BOTH, expand=True)

        # Метка для неиспользованной области
        self.unused_label = CTkLabel(self.center_frame, text="")
        self.unused_label.pack(anchor='w', padx=10, pady=5)

        # Список заказов (стеклопакетов)
        self.order_list_frame = CTkFrame(self.left_frame)
        self.order_list_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        self.scrollbar = CTkScrollbar(self.order_list_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.order_listbox = tk.Listbox(self.order_list_frame, yscrollcommand=self.scrollbar.set, height=15,
                                        bg="#333333", fg="white", font=("Arial", 12))
        self.order_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.configure(command=self.order_listbox.yview)

        # Привязка событий
        self.card_canvas.bind("<Control-MouseWheel>", self.on_mousewheel_zoom)
        self.card_canvas.bind("<ButtonPress-2>", self.reset_view)  # Для сброса масштаба
        self.card_canvas.bind("<Motion>", self.store_mouse_position)
        self.card_canvas.bind("<Motion>", self.on_canvas_hover)
        self.card_canvas.bind("<Button-1>", self.on_canvas_click)
        self.card_canvas.bind("<Leave>", self.hide_tooltip)
        # Привязываем обработчик изменения размера
        self.card_canvas.bind("<Configure>", self._on_canvas_resize)
        self.card_listbox.bind('<<ListboxSelect>>', self.on_card_select)

        self.select_default_card()
        self.update_orders_list()

    def store_mouse_position(self, event):
        """Запоминаем текущую позицию курсора"""
        self.last_mouse_pos = (event.x, event.y)

    def reset_view(self, event=None):
        """Сброс масштаба и положения"""
        self.current_zoom = 1.0
        self.canvas_offset_x = 0
        self.canvas_offset_y = 0
        if self.groups and self.card_listbox.curselection():
            self.display_cutting_plan(self.card_listbox.curselection()[0])

    def on_mousewheel_zoom(self, event):
        """Масштабирование с учетом позиции курсора"""
        if not self.groups or not self.card_listbox.curselection():
            return

        # Определяем направление масштабирования
        zoom_factor = 1.1 if event.delta > 0 else 1 / 1.1
        new_zoom = max(self.min_zoom, min(self.max_zoom, self.current_zoom * zoom_factor))

        if new_zoom == self.current_zoom:
            return

        # Сохраняем позицию курсора в реальных координатах (до масштабирования)
        group = self.groups[self.card_listbox.curselection()[0]]
        scale = self.get_current_scale(group)
        x = self.card_canvas.canvasx(event.x) / scale
        y = self.card_canvas.canvasy(event.y) / scale

        # Обновляем масштаб
        self.current_zoom = new_zoom

        # Перерисовываем план раскроя
        self.display_cutting_plan(self.card_listbox.curselection()[0])

        # Вычисляем новую позицию курсора после масштабирования
        new_scale = self.get_current_scale(group)
        new_x = x * new_scale
        new_y = y * new_scale

        # Прокручиваем холст так, чтобы точка под курсором осталась на месте
        self.card_canvas.scan_mark(event.x, event.y)
        self.card_canvas.scan_dragto(
            int(event.x - (new_x - x * scale)),
            int(event.y - (new_y - y * scale)),
            gain=1
        )

    def display_cutting_plan(self, index):
        """Отображение с учетом текущего масштаба (без автоскейла)"""
        if not self.groups or index >= len(self.groups):
            return

        group = self.groups[index]
        self.card_canvas.delete("all")

        # Рассчитываем базовый масштаб для полного отображения
        canvas_width = self.card_canvas.winfo_width()
        canvas_height = self.card_canvas.winfo_height()
        base_scale = min(canvas_width / group['width'],
                         canvas_height / group['height']) * 0.95  # Небольшой отступ

        # Применяем текущий масштаб
        scale = base_scale * self.current_zoom

        # Рисуем все элементы
        self.draw_background(group, scale)
        self.draw_grid(group, scale)

        # Границы листа
        self.card_canvas.create_rectangle(
            0, 0,
            group['width'] * scale,
            group['height'] * scale,
            outline="black", width=3
        )

        # Рисуем элементы
        for item in group['items']:
            self.draw_glass_item(item, scale)

        # Обновляем инфопанель
        self.update_info_panel(group, index, scale)

        # Восстанавливаем выделение
        if hasattr(self, 'selected_item') and self.selected_item:
            for item in group['items']:
                if item['id'] == self.selected_item['id']:
                    self.select_item(item, scale)
                    break

    def _on_canvas_resize(self, event):
        """Обновляет отображение при изменении размера холста"""
        if (hasattr(self, 'groups') and self.groups and
                hasattr(self, 'card_listbox') and self.card_listbox.curselection()):
            selected = self.card_listbox.curselection()
            if selected:
                self.display_cutting_plan(selected[0])

    def update_orders_list(self):
        """Обновляет список заказов из базы данных, исключая завершенные"""
        self.order_listbox.delete(0, tk.END)  # Очищаем текущий список

        # Получаем заказы, исключая завершенные
        orders = get_production_orders()

        if not orders:
            return

        for order in orders:
            order_id = order[0]
            status = order[-1]

            # Пропускаем завершенные заказы
            if status.lower() == "завершен":
                continue

            windows = get_windows_for_production_order(order_id)

            for window in windows:
                window_id, window_type, width, height, quantity = window
                for _ in range(quantity):
                    order_text = f"Заказ {order_id}: {width}x{height} ({window_type})"
                    self.order_listbox.insert(tk.END, order_text)

    def load_orders_from_db(self):
        """Алиас для update_orders_list для обратной совместимости"""
        self.update_orders_list()

    def on_close(self):
        """Обработчик закрытия окна"""
        self._is_running = False
        try:
            # Очищаем очередь, чтобы избежать попыток обновления после закрытия
            self._gui_update_queue.clear()

            # Даем время на завершение операций
            self.after(100, self.destroy)
        except Exception as e:
            print(f"Ошибка при закрытии: {e}")

    def _process_gui_updates(self):
        """Обрабатывает все ожидающие обновления GUI из главного потока"""
        while self._gui_update_queue and self._is_running:
            try:
                func, args, kwargs = self._gui_update_queue.pop(0)
                func(*args, **kwargs)
            except Exception as e:
                print(f"Ошибка при обработке обновления GUI: {e}")

        if self._is_running:
            self.after(100, self._process_gui_updates)

    def _safe_gui_update(self, func, *args, **kwargs):
        """Добавляет обновление GUI в очередь (можно вызывать из любого потока)"""
        if self._is_running:
            self._gui_update_queue.append((func, args, kwargs))

    def set_optimization_mode(self, choice):
        """Устанавливает режим оптимизации"""
        self.optimization_mode = "deep" if "Глубокий" in choice else "normal"

    def resize_left_panel(self, event):
        """Изменяет ширину левой панели"""
        new_width = event.x_root - self.left_frame.winfo_rootx()
        if 150 <= new_width <= 400:  # Ограничиваем минимальную и максимальную ширину
            self.side_panels_width = new_width
            self.left_frame.configure(width=new_width)
            self.left_frame.pack_configure(padx=10 if new_width > 160 else 5)

    def resize_right_panel(self, event):
        """Изменяет ширину правой панели"""
        new_width = self.right_frame.winfo_rootx() + self.right_frame.winfo_width() - event.x_root
        if 150 <= new_width <= 400:  # Ограничиваем минимальную и максимальную ширину
            self.side_panels_width = new_width
            self.right_frame.configure(width=new_width)
            self.right_frame.pack_configure(padx=10 if new_width > 160 else 5)

    def show_panel_context_menu(self, event):
        """Показывает контекстное меню для управления панелями"""
        # Определяем, какая панель была кликнута
        if event.widget == self.left_separator:
            panel = "left"
        elif event.widget == self.right_separator:
            panel = "right"
        else:
            return

        # Создаем меню
        menu = tk.Menu(self, tearoff=0)

        # Добавляем команды
        menu.add_command(
            label="Увеличить ширину",
            command=lambda: self.adjust_panel_width(panel, "increase")
        )
        menu.add_command(
            label="Уменьшить ширину",
            command=lambda: self.adjust_panel_width(panel, "decrease")
        )
        menu.add_separator()
        menu.add_command(
            label="Сбросить ширину",
            command=lambda: self.reset_panel_width(panel)
        )
        menu.add_separator()
        menu.add_command(
            label="Открыть в отдельном окне",
            command=lambda: self.open_panel_in_window(panel)
        )

        # Показываем меню
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def adjust_panel_width(self, panel, action):
        """Регулирует ширину панели"""
        step = 20  # Шаг изменения
        current_width = self.side_panels_width

        if action == "increase" and current_width < 400:
            new_width = current_width + step
        elif action == "decrease" and current_width > 150:
            new_width = current_width - step
        else:
            return

        self.side_panels_width = new_width

        if panel == "left":
            self.left_frame.configure(width=new_width)
        else:
            self.right_frame.configure(width=new_width)

    def reset_panel_width(self, panel):
        """Сбрасывает ширину панели к значению по умолчанию"""
        self.side_panels_width = 200
        if panel == "left":
            self.left_frame.configure(width=200)
        else:
            self.right_frame.configure(width=200)

    def open_panel_in_window(self, panel):
        """Открывает панель в отдельном окне"""
        if panel == "left":
            content = self.left_frame
            title = "Список заказов"
        else:
            content = self.right_frame
            title = "Карты раскроя"

        # Создаем новое окно
        new_window = tk.Toplevel(self)
        new_window.title(title)
        new_window.geometry(f"{self.side_panels_width}x600")

        # Перемещаем содержимое в новое окно
        content.pack_forget()
        content.pack(in_=new_window, fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Настраиваем закрытие окна
        def on_close():
            content.pack_forget()
            if panel == "left":
                content.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=10, expand=False)
            else:
                content.pack(side=tk.RIGHT, fill=tk.Y, padx=10, pady=10, expand=False)
            new_window.destroy()

        new_window.protocol("WM_DELETE_WINDOW", on_close)

        # Делаем окно изменяемым по размеру
        new_window.resizable(True, True)

    def select_default_card(self):
        """Выбирает первую карту раскроя по умолчанию"""
        if self.card_listbox.size() > 0:
            self.card_listbox.selection_set(0)
            self.card_listbox.see(0)
            self.display_cutting_plan(0)

    def on_canvas_hover(self, event):
        """Всплывающая подсказка при наведении с учетом масштаба"""
        if not self.groups or not self.card_listbox.curselection():
            self.hide_tooltip()
            return

        group = self.groups[self.card_listbox.curselection()[0]]
        scale = self.get_current_scale(group)

        # Преобразуем координаты курсора с учетом текущего масштаба
        real_x = self.card_canvas.canvasx(event.x) / scale
        real_y = self.card_canvas.canvasy(event.y) / scale

        # Ищем заготовку под курсором
        new_hover_item = None
        for item in group['items']:
            x1 = item['x']
            y1 = item['y']
            x2 = x1 + item['width']
            y2 = y1 + item['height']

            if x1 <= real_x <= x2 and y1 <= real_y <= y2:
                new_hover_item = item
                break

        # Обновляем hover-эффект
        if new_hover_item != self.hover_item:
            self.update_hover_effect(new_hover_item, scale)
            self.hover_item = new_hover_item

            if new_hover_item:
                self.show_tooltip(event.x, event.y, new_hover_item)
            else:
                self.hide_tooltip()

    def update_hover_effect(self, item, scale):
        """Обновляет подсветку при наведении"""
        if self.hover_rect:
            self.card_canvas.delete(self.hover_rect)
            self.hover_rect = None

        if item:
            x1 = item['x'] * scale
            y1 = item['y'] * scale
            x2 = x1 + item['width'] * scale
            y2 = y1 + item['height'] * scale

            self.hover_rect = self.card_canvas.create_rectangle(
                x1, y1, x2, y2,
                outline="#FFA500", width=2, dash=(3, 3)
            )
            self.card_canvas.tag_raise(self.hover_rect)

    def show_tooltip(self, x, y, item):
        """Показывает всплывающую подсказку в фиксированной позиции относительно курсора"""
        # Сначала скрываем предыдущую подсказку
        self.hide_tooltip()

        text = f"ID: {item['id']} | {item['width']}×{item['height']} мм"
        if item.get('rotation'):
            text += f" (повернуто)"

        # Фиксированное смещение от курсора
        offset_x = 15
        offset_y = 15

        # Создаем фон для tooltip
        self.tooltip_bg = self.card_canvas.create_rectangle(
            x + offset_x, y + offset_y,
            x + offset_x + 150, y + offset_y + 30,
            fill="#FFFFE0",
            outline="#CCCCCC",
            tags="tooltip"
        )

        # Затем создаем текст
        self.tooltip_text = self.card_canvas.create_text(
            x + offset_x + 5, y + offset_y + 5,
            text=text,
            font=("Arial", 10),
            fill="black",
            anchor="nw",
            tags="tooltip"
        )

        # Поднимаем подсказку на верхний уровень
        self.card_canvas.tag_raise(self.tooltip_bg)
        self.card_canvas.tag_raise(self.tooltip_text)

    def hide_tooltip(self, event=None):
        """Скрывает всплывающую подсказку"""
        if hasattr(self, 'tooltip_bg') and self.tooltip_bg:
            self.card_canvas.delete(self.tooltip_bg)
            self.tooltip_bg = None

        if hasattr(self, 'tooltip_text') and self.tooltip_text:
            self.card_canvas.delete(self.tooltip_text)
            self.tooltip_text = None

        if hasattr(self, 'hover_rect') and self.hover_rect:
            self.card_canvas.delete(self.hover_rect)
            self.hover_rect = None

        self.hover_item = None

    def on_canvas_click(self, event):
        """Обработчик клика с учетом масштаба"""
        try:
            if not self.groups or not self.card_listbox.curselection():
                return

            group = self.groups[self.card_listbox.curselection()[0]]
            scale = self.get_current_scale(group)

            # Преобразуем координаты клика с учетом текущего масштаба
            real_x = self.card_canvas.canvasx(event.x) / scale
            real_y = self.card_canvas.canvasy(event.y) / scale

            self.clear_selection()

            clicked_item = None
            for item in group['items']:
                x1 = item['x']
                y1 = item['y']
                x2 = x1 + item['width']
                y2 = y1 + item['height']

                if x1 <= real_x <= x2 and y1 <= real_y <= y2:
                    clicked_item = item
                    break

            if clicked_item:
                self.create_selection_rect(clicked_item, scale)
                self.selected_item = clicked_item
                self.select_order_in_list(clicked_item['id'])

        except Exception as e:
            print(f"Ошибка при клике: {e}")
            self.clear_selection()

    def canvas_to_real_coords(self, x, y):
        """Преобразует координаты холста в реальные координаты с учетом масштаба и смещения"""
        if not self.groups or not self.card_listbox.curselection():
            return x, y

        group = self.groups[self.card_listbox.curselection()[0]]
        scale = self.get_current_scale(group)

        real_x = (x - self.canvas_offset_x) / scale
        real_y = (y - self.canvas_offset_y) / scale

        return real_x, real_y

    def create_selection_rect(self, item, scale):
        """Создает прямоугольник выделения"""
        x1 = item['x'] * scale
        y1 = item['y'] * scale
        x2 = x1 + item['width'] * scale
        y2 = y1 + item['height'] * scale

        self.selection_rect = self.card_canvas.create_rectangle(
            x1, y1, x2, y2,
            outline="#00FF00", width=3, tags="selection"
        )
        self.card_canvas.tag_raise(self.selection_rect)

    def clear_selection(self):
        """Сбрасывает выделение"""
        if hasattr(self, 'selection_rect') and self.selection_rect:
            self.card_canvas.delete(self.selection_rect)
        self.selection_rect = None
        self.selected_item = None
        if hasattr(self, 'order_listbox'):
            self.order_listbox.selection_clear(0, tk.END)

    def on_card_select(self, event):
        """Обработчик выбора карты раскроя"""
        if not self.card_listbox.curselection():
            return

        selected_index = self.card_listbox.curselection()[0]
        if 0 <= selected_index < len(self.groups):
            self.display_cutting_plan(selected_index)
            self.clear_selection()  # Сбрасываем выделение при смене карты

    def select_item(self, item, scale):
        """Выделяет указанную заготовку (без предварительной очистки)"""
        x1 = item['x'] * scale
        y1 = item['y'] * scale
        x2 = x1 + item['width'] * scale
        y2 = y1 + item['height'] * scale

        self.selection_rect = self.card_canvas.create_rectangle(
            x1, y1, x2, y2,
            outline="#00FF00", width=3, tags="selection"
        )
        self.card_canvas.tag_raise(self.selection_rect)
        self.selected_item = item

    def select_order_in_list(self, order_id):
        """Выбирает заказ в списке с защитой от ошибок"""
        try:
            # Очищаем текущее выделение
            self.order_listbox.selection_clear(0, tk.END)

            # Ищем нужный заказ
            for i in range(self.order_listbox.size()):
                item_text = self.order_listbox.get(i)
                if str(order_id) in item_text.split(':')[0]:  # Ищем в начале строки (ID заказа)
                    self.order_listbox.selection_set(i)
                    self.order_listbox.see(i)
                    self.order_listbox.activate(i)  # Активируем элемент
                    break
        except Exception as e:
            print(f"Ошибка при выборе заказа: {e}")

    def get_current_scale(self, group):
        """Возвращает текущий масштаб отображения с учетом зума"""
        canvas_width = self.card_canvas.winfo_width()
        canvas_height = self.card_canvas.winfo_height()

        # Базовый масштаб для полного отображения
        base_scale = min(canvas_width / group['width'],
                         canvas_height / group['height']) * 0.95  # Небольшой отступ

        # Применяем текущий масштаб
        return base_scale * self.current_zoom

    def load_orders_from_db(self):
        """Загружает стеклопакеты из production orders с сохранением типа"""
        self.order_listbox.delete(0, tk.END)
        orders = get_production_orders()
        if not orders:
            return

        for order in orders:
            order_id = order[0]
            windows = get_windows_for_production_order(order_id)

            for window in windows:
                window_id, window_type, width, height, quantity = window
                for _ in range(quantity):
                    # Сохраняем тип в тексте заказа
                    order_text = f"Заказ {order_id}: {width}x{height} ({window_type})"
                    self.order_listbox.insert(tk.END, order_text)

    def optimize_cutting(self):
        """Запускает оптимизацию с выбранным режимом"""
        try:
            self.sheet_width = int(self.entry_sheet_width.get())
            self.sheet_height = int(self.entry_sheet_height.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Неверные размеры листа стекла")
            return

        # Показываем элементы прогресса
        self.progress_bar.pack(pady=5, fill=tk.X, padx=20)
        self.status_label.pack(pady=5)
        self.progress_bar.set(0)
        self.status_label.configure(text="Подготовка данных...")
        self.update_idletasks()

        orders = self._fetch_orders()
        if not orders:
            messagebox.showwarning("Предупреждение", "Нет заказов для оптимизации.")
            self._hide_progress()
            return

        # Запускаем в отдельном потоке
        threading.Thread(
            target=self._run_optimization,
            args=(orders,),
            daemon=True
        ).start()

    def _run_optimization(self, orders):
        """Выполняет оптимизацию в фоновом потоке с безопасными обновлениями GUI"""
        try:
            # Проверяем флаг перед началом работы
            if not getattr(self, '_is_running', False):
                return

            # Инициализация через безопасные обновления GUI
            self._safe_gui_update(self._initialize_optimization_ui)

            # Основные параметры оптимизации
            self.groups = []
            self.unused_elements = defaultdict(list)
            self.packing_cache = {}
            self.lock = threading.Lock()
            sheet_area = self.sheet_width * self.sheet_height

            # Настройки алгоритма в зависимости от режима
            optimization_params = self._get_optimization_params()

            # Группировка заказов по типам
            type_groups = defaultdict(list)
            for order in orders:
                type_groups[order['type']].append(order)

            total_types = len(type_groups)
            processed_types = 0

            # Функция обработки одного типа окон
            def process_type(window_type, type_orders):
                nonlocal processed_types
                try:
                    if not self._is_running:
                        return

                    # Подготовка элементов
                    items = self._prepare_items(type_orders)
                    self._safe_gui_update(self._update_status,
                                          f"Обработка {window_type} ({len(items)} элементов)...")

                    # Сортировка и обработка элементов
                    items.sort(key=lambda x: x['width'] * x['height'], reverse=True)
                    remaining_items = items.copy()

                    while remaining_items and self._is_running:
                        # Создаем новый лист для упаковки
                        current_sheet = self._create_new_sheet(window_type)

                        # Основной цикл заполнения листа
                        self._fill_sheet(
                            current_sheet,
                            remaining_items,
                            sheet_area,
                            optimization_params
                        )

                        # Сохраняем результат если есть элементы
                        if current_sheet['items'] and self._is_running:
                            self._save_sheet_result(
                                current_sheet,
                                window_type,
                                sheet_area,
                                remaining_items,
                                optimization_params
                            )

                    # Сохраняем неиспользованные элементы
                    with self.lock:
                        self.unused_elements[window_type] = remaining_items

                    # Обновляем прогресс
                    processed_types += 1
                    progress = processed_types / total_types
                    self._safe_gui_update(self._update_progress, progress)

                except Exception as e:
                    print(f"Ошибка при обработке типа {window_type}: {e}")

            # Запускаем обработку в пуле потоков
            with ThreadPoolExecutor(max_workers=optimization_params['max_workers']) as executor:
                futures = [
                    executor.submit(process_type, wt, to)
                    for wt, to in type_groups.items()
                ]

                for future in as_completed(futures):
                    if not self._is_running:
                        executor.shutdown(wait=False)
                        return
                    future.result()

            # Завершаем оптимизацию
            if self._is_running:
                self._safe_gui_update(self._optimization_complete)

        except Exception as e:
            print(f"Ошибка в потоке оптимизации: {e}")
            if getattr(self, '_is_running', False):
                self._safe_gui_update(self._optimization_failed, str(e))

    def _initialize_optimization_ui(self):
        """Инициализирует UI для оптимизации"""
        self.progress_bar.pack(pady=5, fill=tk.X, padx=20)
        self.status_label.pack(pady=5)
        self.progress_bar.set(0)
        self.status_label.configure(text="Инициализация...")

    def _get_optimization_params(self):
        """Возвращает параметры оптимизации в зависимости от режима"""
        if self.optimization_mode == "deep":
            return {
                'min_fill_ratio': 0.98,
                'max_workers': 2,
                'sort_algo': rectpack.SORT_SSIDE,
                'pack_algo': rectpack.MaxRectsBaf,
                'attempts_per_item': 3
            }
        else:
            return {
                'min_fill_ratio': 0.90,
                'max_workers': 4,
                'sort_algo': rectpack.SORT_AREA,
                'pack_algo': rectpack.MaxRectsBssf,
                'attempts_per_item': 1
            }

    def _prepare_items(self, type_orders):
        """Подготавливает элементы для обработки"""
        items = []
        # Создаем счетчик для каждого заказа
        order_counters = defaultdict(int)

        for o in type_orders:
            order_id = o['order_id']
            order_counters[order_id] += 1
            items.append({
                'id': f"{order_id}-{order_counters[order_id]}",  # Формат "order_id-sequence"
                'width': o['width'],
                'height': o['height'],
                'type': o['type'],
                'original': o
            })

        return items

    def _create_new_sheet(self, window_type):
        """Создает новый лист для упаковки"""
        return {
            'items': [],
            'used_area': 0,
            'type': window_type
        }

    def _fill_sheet(self, current_sheet, remaining_items, sheet_area, params):
        """Заполняет лист элементами"""
        while (current_sheet['used_area'] / sheet_area < params['min_fill_ratio'] and
               remaining_items and self._is_running):

            best_item = None
            best_packed = None
            best_used_area = None
            best_fill = current_sheet['used_area'] / sheet_area

            # Пробуем несколько вариантов для глубокого перебора
            for _ in range(params['attempts_per_item']):
                for i, item in enumerate(remaining_items):
                    if not self._is_running:
                        return

                    test_items = current_sheet['items'] + [item]
                    cache_key = self._create_cache_key(test_items)

                    if cache_key in self.packing_cache:
                        packed, used_area = self.packing_cache[cache_key]
                    else:
                        packed, used_area = self._pack_items_safe(
                            test_items,
                            sort_algo=params['sort_algo'],
                            pack_algo=params['pack_algo']
                        )
                        with self.lock:
                            self.packing_cache[cache_key] = (packed, used_area)

                    if (used_area / sheet_area > best_fill and
                            len(packed) == len(test_items)):
                        best_fill = used_area / sheet_area
                        best_item = item
                        best_packed = packed
                        best_used_area = used_area

                        if best_fill >= params['min_fill_ratio']:
                            break

                if best_item:
                    break

            if best_item:
                current_sheet['items'] = best_packed
                current_sheet['used_area'] = best_used_area
                remaining_items.remove(best_item)
            else:
                break

    def _save_sheet_result(self, current_sheet, window_type, sheet_area, remaining_items, params):
        """Сохраняет результат упаковки листа"""
        with self.lock:
            self._add_completed_sheet(current_sheet, window_type, sheet_area)

        # Дополнительная попытка добавить элементы в глубоком режиме
        if (self.optimization_mode == "deep" and
                self._is_running and
                remaining_items):
            self._try_add_remaining_items_deep(
                current_sheet,
                remaining_items,
                window_type,
                sheet_area,
                params['sort_algo'],
                params['pack_algo']
            )


    def _update_status(self, message):
        """Обновляет статус в GUI"""
        self.status_label.configure(text=message)
        self.update_idletasks()

    def _update_progress(self, value):
        """Обновляет прогресс-бар"""
        self.progress_bar.set(value)
        self.update_idletasks()

    def _optimization_complete(self):
        """Действия после завершения оптимизации"""
        self._print_final_statistics()
        self.update_interface()
        self.select_default_card()
        self.load_orders_from_db()

        self.status_label.configure(text="Оптимизация завершена!")
        self.progress_bar.set(1.0)
        self.after(2000, lambda: self._hide_progress())

    def _optimization_failed(self, error_msg):
        """Действия при ошибке оптимизации"""
        if not self.winfo_exists():
            return

        self._hide_progress()
        try:
            messagebox.showerror("Ошибка оптимизации", f"Произошла ошибка:\n{error_msg}")
            self.status_label.configure(text=f"Ошибка: {error_msg}")
        except Exception as e:
            print("Не удалось показать сообщение об ошибке:", e)

    def _hide_progress(self):
        """Скрывает элементы прогресса"""
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.status_label.pack_forget()

    def _create_cache_key(self, items):
        """Создает безопасный ключ для кэша"""
        sizes = tuple(sorted((item['width'], item['height']) for item in items))
        return hash(sizes)

    def _pack_items_safe(self, items, sort_algo=rectpack.SORT_AREA, pack_algo=rectpack.MaxRectsBssf):
        """Безопасная упаковка с настраиваемыми алгоритмами"""
        try:
            return self.pack_items(items, sort_algo, pack_algo)
        except Exception as e:
            print(f"Ошибка упаковки: {e}")
            return [], 0

    def _add_completed_sheet(self, sheet, window_type, sheet_area):
        """Добавляет завершенную карту в общий список"""
        self.groups.append({
            'width': self.sheet_width,
            'height': self.sheet_height,
            'items': sheet['items'],
            'type': window_type,
            'used_area': sheet['used_area'],
            'wasted_area': sheet_area - sheet['used_area'],
            'fill_percentage': (sheet['used_area'] / sheet_area) * 100
        })
        print(
            f"[{window_type}] Добавлена карта {len(self.groups)}. Заполнение: {self.groups[-1]['fill_percentage']:.1f}%")

    def _try_add_remaining_items_deep(self, current_sheet, remaining_items, window_type, sheet_area, sort_algo,
                                      pack_algo):
        """Дополнительная попытка добавить элементы в глубоком режиме"""
        added_count = 0
        temp_remaining = remaining_items.copy()

        for item in temp_remaining:
            if not self._is_running:
                break

            # Пробуем несколько вариантов размещения
            for rotation in [0, 90]:  # Пробуем оба варианта поворота
                width = item['width'] if rotation == 0 else item['height']
                height = item['height'] if rotation == 0 else item['width']

                test_items = current_sheet['items'] + [{
                    **item,
                    'width': width,
                    'height': height,
                    'rotation': rotation
                }]

                cache_key = self._create_cache_key(test_items)

                if cache_key in self.packing_cache:
                    packed, used_area = self.packing_cache[cache_key]
                else:
                    packed, used_area = self._pack_items_safe(
                        test_items,
                        sort_algo=sort_algo,
                        pack_algo=pack_algo
                    )
                    with self.lock:
                        self.packing_cache[cache_key] = (packed, used_area)

                if len(packed) == len(test_items) and used_area <= sheet_area:
                    current_sheet['items'] = packed
                    current_sheet['used_area'] = used_area
                    remaining_items.remove(item)
                    added_count += 1
                    print(f"[Глубокий режим] Добавлен элемент {item['id']} с поворотом {rotation}°")
                    break  # Переходим к следующему элементу

        if added_count > 0 and self._is_running:
            with self.lock:
                self.groups[-1].update({
                    'items': current_sheet['items'],
                    'used_area': current_sheet['used_area'],
                    'wasted_area': sheet_area - current_sheet['used_area'],
                    'fill_percentage': (current_sheet['used_area'] / sheet_area) * 100
                })

        return added_count

    def _print_final_statistics(self):
        """Выводит итоговую статистику после оптимизации"""
        print("\n" + "=" * 50)
        print("ИТОГОВАЯ СТАТИСТИКА")
        print("=" * 50)

        total_sheets = len(self.groups)
        total_used_area = sum(g['used_area'] for g in self.groups)
        total_wasted_area = sum(g['wasted_area'] for g in self.groups)
        try:
            avg_fill = (total_used_area / (total_sheets * self.sheet_width * self.sheet_height)) * 100
        except ZeroDivisionError:
            avg_fill = 0

        print(f"\nВсего карт раскроя: {total_sheets}")
        print(f"Среднее заполнение: {avg_fill:.1f}%")
        print(f"Общая использованная площадь: {total_used_area / 1e6:.2f} м²")
        print(f"Общие отходы: {total_wasted_area / 1e6:.2f} м²")

        for window_type, items in self.unused_elements.items():
            if items:
                print(f"\nНе упаковано элементов типа {window_type}: {len(items)}")
                for item in items[:3]:
                    print(f" - {item['id']}: {item['width']}x{item['height']} мм")
                if len(items) > 3:
                    print(f" - ...и еще {len(items) - 3} элементов")

    def _print_unused_elements(self, unused_elements):
        """Выводит информацию о неиспользованных элементах"""
        print("\nНЕИСПОЛЬЗОВАННЫЕ ЭЛЕМЕНТЫ:")
        total_unused = 0

        for window_type, items in unused_elements.items():
            if items:
                print(f"\nТип: {window_type}")
                print(f"Количество: {len(items)}")
                total_unused += len(items)

                for item in items[:5]:  # Выводим первые 5 элементов каждого типа
                    print(f" - {item['original']['order_id']}-{item['original']['window_id']}: "
                          f"{item['width']}x{item['height']} (площадь: {item['width'] * item['height']})")

                if len(items) > 5:
                    print(f" - ...и еще {len(items) - 5} элементов")

        if total_unused == 0:
            print("Все элементы успешно упакованы!")
        else:
            print(f"\nВсего неиспользованных элементов: {total_unused}")

    def _mark_unused_orders(self, unused_elements):
        """Помечает неиспользованные заказы в интерфейсе"""
        # Собираем все неиспользованные ID заказов
        unused_ids = set()
        for items in unused_elements.values():
            for item in items:
                unused_ids.add(item['original']['order_id'] + "-" + item['original']['window_id'])

        # Помечаем элементы в интерфейсе
        for i in range(self.order_listbox.size()):
            order_id = self.order_listbox.get(i)
            if order_id in unused_ids:
                self.order_listbox.itemconfig(i, {'bg': 'lightyellow', 'fg': 'red'})
                self.order_listbox.itemconfig(i, {'tags': ['unused']})

    def _print_sheet_summary(self, sheet_index, items, used_area, total_area):
        """Визуализация заполнения карты раскроя"""
        fill_ratio = used_area / total_area
        fill_percent = int(fill_ratio * 100)
        bar_length = 30
        filled = int(bar_length * fill_ratio)
        empty = bar_length - filled
        bar = '█' * filled + '-' * empty
        elem_count = len(items)

        # Цветовая индикация
        if fill_percent >= 90:
            color = "\033[92m"  # Зеленый
        elif fill_percent >= 70:
            color = "\033[93m"  # Желтый
        else:
            color = "\033[91m"  # Красный

        reset = "\033[0m"

        print(
            f"{color}[Карта #{sheet_index + 1}] Заполнение: {fill_percent:3d}% | {bar} | Элементов: {elem_count:2d}{reset}")

        # Дополнительная информация для отладки
        if fill_percent < 70:
            print(f"   Предупреждение: низкое заполнение ({fill_percent}%)")
            for item in items:
                print(
                    f"   Элемент {item['id']}: {item['width']}x{item['height']} (площадь: {item['width'] * item['height']})")

    def _add_cutting_map(self, packed_items, used_area, window_type):
        """Добавление карты раскроя с дополнительной информацией"""
        fill_percent = (used_area / (self.sheet_width * self.sheet_height)) * 100
        self.groups.append({
            'width': self.sheet_width,
            'height': self.sheet_height,
            'items': packed_items,
            'type': window_type,
            'fill_percentage': fill_percent,
            'used_area': used_area,
            'wasted_area': (self.sheet_width * self.sheet_height) - used_area
        })

    def _try_add_remaining_items_deep(self, current_sheet, remaining_items, window_type, sheet_area, sort_algo,
                                      pack_algo):
        """Дополнительная попытка добавить элементы в глубоком режиме оптимизации"""
        added_count = 0
        temp_remaining = remaining_items.copy()

        for item in temp_remaining:
            # Пробуем несколько вариантов размещения
            for rotation in [0, 90]:  # Пробуем оба варианта поворота
                width = item['width'] if rotation == 0 else item['height']
                height = item['height'] if rotation == 0 else item['width']

                test_items = current_sheet['items'] + [{
                    **item,
                    'width': width,
                    'height': height,
                    'rotation': rotation
                }]

                cache_key = self._create_cache_key(test_items)

                if cache_key in self.packing_cache:
                    packed, used_area = self.packing_cache[cache_key]
                else:
                    packed, used_area = self._pack_items_safe(
                        test_items,
                        sort_algo=sort_algo,
                        pack_algo=pack_algo
                    )
                    with self.lock:
                        self.packing_cache[cache_key] = (packed, used_area)

                if len(packed) == len(test_items) and used_area <= sheet_area:
                    current_sheet['items'] = packed
                    current_sheet['used_area'] = used_area
                    remaining_items.remove(item)
                    added_count += 1
                    print(f"[Глубокий режим] Добавлен элемент {item['id']} с поворотом {rotation}°")
                    break  # Переходим к следующему элементу

        if added_count > 0:
            with self.lock:
                self.groups[-1].update({
                    'items': current_sheet['items'],
                    'used_area': current_sheet['used_area'],
                    'wasted_area': sheet_area - current_sheet['used_area'],
                    'fill_percentage': (current_sheet['used_area'] / sheet_area) * 100
                })

        return added_count

    def pack_items(self, items, sort_algo=rectpack.SORT_AREA, pack_algo=rectpack.MaxRectsBssf):
        """Упаковка элементов с настраиваемыми параметрами"""
        if not items:
            return [], 0

        packer = rectpack.newPacker(
            rotation=True,
            sort_algo=sort_algo,
            pack_algo=pack_algo,
            bin_algo=rectpack.PackingBin.Global
        )

        # Создаем словарь для хранения элементов
        items_dict = {}
        for idx, item in enumerate(items):
            if not all(key in item for key in ['id', 'width', 'height']):
                continue

            unique_id = f"{item['id']}_{idx}"
            items_dict[unique_id] = item
            packer.add_rect(item['width'], item['height'], unique_id)

        if not items_dict:
            return [], 0

        # Добавляем область для упаковки
        packer.add_bin(self.sheet_width, self.sheet_height)

        try:
            packer.pack()
        except Exception as e:
            print(f"Ошибка при упаковке: {e}")
            return [], 0

        packed_items = []
        area_used = 0

        # Обрабатываем все упакованные прямоугольники
        for abin in packer:
            if not abin:  # Пропускаем пустые бины
                continue

            for rect in abin:
                item = items_dict.get(rect.rid)
                if not item:
                    continue

                # Определяем параметры упакованного элемента
                try:
                    rotation = 90 if (rect.width != item['width'] or
                                      rect.height != item['height']) else 0

                    packed_item = {
                        'id': item['id'],
                        'x': rect.x,
                        'y': rect.y,
                        'width': rect.width,
                        'height': rect.height,
                        'rotation': rotation,
                        'type': item.get('type', 'unknown'),
                        'original_data': item.get('original', {})
                    }
                    packed_items.append(packed_item)
                    area_used += rect.width * rect.height
                except Exception as e:
                    print(f"Ошибка обработки элемента {rect.rid}: {e}")
                    continue

        return packed_items, area_used

    def _fetch_orders(self):
        """Получение незавершенных заказов из базы данных"""
        orders = []
        production_orders = get_production_orders()

        if not production_orders:
            return []

        for order in production_orders:
            order_id = order[0]
            status = order[-1]  # Предполагаем, что статус находится во втором поле

            # Пропускаем завершенные заказы
            if status.lower() == "завершен":
                continue

            windows = get_windows_for_production_order(order_id)
            for window in windows:
                window_id, window_type, width, height, quantity = window
                for _ in range(quantity):
                    orders.append({
                        'order_id': order_id,
                        'window_id': window_id,
                        'width': width,
                        'height': height,
                        'type': window_type
                    })
        return orders


    def _item_to_tuple(self, item):
        """Преобразование элемента в хешируемый кортеж"""
        return (
            item['order_id'],
            item['window_id'],
            item['width'],
            item['height'],
            item['type']
        )

    def _tuple_to_item(self, tpl):
        """Обратное преобразование кортежа в элемент"""
        return {
            'order_id': tpl[0],
            'window_id': tpl[1],
            'width': tpl[2],
            'height': tpl[3],
            'type': tpl[4]
        }


    @lru_cache(maxsize=1000)
    def _cached_pack_items(self, items_tuple):
        """Кэшированная упаковка с хешируемыми аргументами"""
        # Преобразуем кортежи обратно в словари
        items = [self._tuple_to_item(tpl) for tpl in items_tuple]

        # Добавляем уникальные ID для rectpack
        unique_items = []
        for idx, item in enumerate(items):
            unique_items.append({
                **item,
                'id': f"{item['order_id']}-{item['window_id']}-{idx}-{hashlib.md5(str(item).encode()).hexdigest()[:6]}"
            })

        return self.pack_items(unique_items)


    # def best_fit_decreasing_algorithm(self, items: List[Dict]) -> List[Dict]:
    #     """Алгоритм с использованием rectpack для оптимального раскроя"""
    #
    #     packer = rectpack.newPacker(
    #         rotation=True,  # Разрешаем поворот
    #         sort_algo=rectpack.SORT_AREA,  # Сортировка по площади
    #         pack_algo=rectpack.MaxRectsBssf,  # Один из лучших алгоритмов (Best Short Side Fit)
    #         bin_algo=rectpack.PackingBin.BBF,  # Best Bin Fit
    #     )
    #
    #     # Добавляем прямоугольники (width, height, id)
    #     for item in items:
    #         packer.add_rect(item['width'], item['height'], item['id'])
    #
    #     # Добавляем "бин" (лист стекла)
    #     sheet_size = (self.sheet_width, self.sheet_height)
    #     max_bins = 999  # Условно большое число, rectpack сам остановится, когда всё упакует
    #     for _ in range(max_bins):
    #         packer.add_bin(*sheet_size)
    #
    #     # Выполняем упаковку
    #     packer.pack()
    #
    #     sheets = []
    #
    #     # Получаем результат
    #     for abin in packer:
    #         sheet = {
    #             'width': self.sheet_width,
    #             'height': self.sheet_height,
    #             'items': [],
    #             'type': None,  # позже подставится
    #         }
    #         for rect in abin:
    #             x, y = rect.x, rect.y
    #             w, h = rect.width, rect.height
    #             rid = rect.rid
    #
    #             # Найдём оригинальный item по id, чтобы вернуть тип
    #             orig_item = next((i for i in items if i['id'] == rid), None)
    #             if orig_item:
    #                 sheet['items'].append({
    #                     'id': rid,
    #                     'x': x,
    #                     'y': y,
    #                     'width': w,
    #                     'height': h,
    #                     'rotation': 0 if (orig_item['width'] == w and orig_item['height'] == h) else 90,
    #                     'type': orig_item['type'],
    #                 })
    #
    #         sheets.append(sheet)
    #
    #     return sheets

    # def try_place_on_sheet(self, sheet: Dict, item: Dict) -> bool:
    #     """Пытается разместить элемент на конкретном листе без наслоений"""
    #     best_rotation = None
    #     best_rect = None
    #     min_waste = float('inf')
    #
    #     # Проверяем оба варианта поворота
    #     for rotation in [0, 90]:
    #         w = item['width'] if rotation == 0 else item['height']
    #         h = item['height'] if rotation == 0 else item['width']
    #
    #         # Ищем лучшее место среди всех свободных областей
    #         for rect in sheet['remaining_rectangles']:
    #             rx, ry, rw, rh = rect
    #
    #             # Проверяем, помещается ли элемент
    #             if w <= rw and h <= rh:
    #                 # Проверяем, не пересекается ли с уже размещенными элементами
    #                 if not self.check_collision(sheet, rx, ry, w, h):
    #                     waste = (rw * rh) - (w * h)
    #                     if waste < min_waste:
    #                         min_waste = waste
    #                         best_rotation = rotation
    #                         best_rect = rect
    #
    #     # Если нашли подходящее место
    #     if best_rect:
    #         rx, ry, rw, rh = best_rect
    #         w = item['width'] if best_rotation == 0 else item['height']
    #         h = item['height'] if best_rotation == 0 else item['width']
    #
    #         # Добавляем элемент
    #         sheet['items'].append({
    #             'id': item['id'],
    #             'x': rx,
    #             'y': ry,
    #             'width': w,
    #             'height': h,
    #             'rotation': best_rotation,
    #             'type': item.get('type', 'Неизвестный тип')
    #         })
    #
    #         # Обновляем оставшееся пространство
    #         self.update_remaining_space(sheet, best_rect, w, h)
    #         return True
    #
    #     return False

    # def check_collision(self, sheet: Dict, x: int, y: int, w: int, h: int) -> bool:
    #     """Проверяет, пересекается ли новый элемент с уже размещенными"""
    #     new_rect = (x, y, x + w, y + h)
    #
    #     for item in sheet['items']:
    #         existing_rect = (item['x'], item['y'],
    #                          item['x'] + item['width'],
    #                          item['y'] + item['height'])
    #
    #         # Проверка пересечения прямоугольников
    #         if not (new_rect[2] <= existing_rect[0] or  # новый слева от существующего
    #                 new_rect[0] >= existing_rect[2] or  # новый справа от существующего
    #                 new_rect[3] <= existing_rect[1] or  # новый выше существующего
    #                 new_rect[1] >= existing_rect[3]):  # новый ниже существующего
    #             return True  # есть пересечение
    #
    #     return False  # нет пересечений


    def update_interface(self):
        """Обновляет все элементы интерфейса после оптимизации"""
        self.card_listbox.delete(0, tk.END)

        for i, group in enumerate(self.groups):
            used_area = sum(i['width'] * i['height'] for i in group['items'])
            total_area = group['width'] * group['height']
            utilization = used_area / total_area * 100
            group_type = group.get('type', 'Неизвестный тип')
            self.card_listbox.insert(tk.END, f"Карта {i + 1} ({group_type}): {utilization:.1f}%")

        if self.groups:
            self.display_cutting_plan(0)


    def draw_grid(self, group, scale):
        """Рисует сетку с шагом 1000 мм"""
        grid_color = "#cccccc"
        grid_step = 1000  # Шаг сетки в мм

        # Вертикальные линии
        for x in range(0, group['width'] + grid_step, grid_step):
            x_pos = self.canvas_offset_x + x * scale
            self.card_canvas.create_line(
                x_pos, self.canvas_offset_y,
                x_pos, self.canvas_offset_y + group['height'] * scale,
                fill=grid_color, dash=(2, 2)
            )
            # Подписи осей X
            if x > 0:
                self.card_canvas.create_text(
                    x_pos, self.canvas_offset_y + 10,
                    text=f"{x} мм",
                    font=("Arial", 8),
                    anchor=tk.N
                )

        # Горизонтальные линии
        for y in range(0, group['height'] + grid_step, grid_step):
            y_pos = self.canvas_offset_y + y * scale
            self.card_canvas.create_line(
                self.canvas_offset_x, y_pos,
                self.canvas_offset_x + group['width'] * scale, y_pos,
                fill=grid_color, dash=(2, 2)
            )
            # Подписи осей Y
            if y > 0:
                self.card_canvas.create_text(
                    self.canvas_offset_x + 10, y_pos,
                    text=f"{y} мм",
                    font=("Arial", 8),
                    anchor=tk.W
                )

    def draw_background(self, group, scale):
        """Рисует фон с учетом масштаба"""
        self.card_canvas.create_rectangle(
            self.canvas_offset_x, self.canvas_offset_y,
            self.canvas_offset_x + group['width'] * scale,
            self.canvas_offset_y + group['height'] * scale,
            fill="red", stipple="gray25", outline=""
        )

    def draw_glass_item(self, item, scale):
        """Рисует стеклопакет с учетом масштаба"""
        x1 = item['x'] * scale
        y1 = item['y'] * scale
        x2 = x1 + item['width'] * scale
        y2 = y1 + item['height'] * scale

        color = "#4CAF50" if item['rotation'] == 0 else "#2196F3"
        rect_id = self.card_canvas.create_rectangle(
            x1, y1, x2, y2,
            outline="black", fill=color, width=1,
            tags=("glass_item", f"item_{item['id']}")
        )

        # Добавляем подписи
        if (x2 - x1) > 50 * scale:
            font_size = max(8, min(12, int((x2 - x1) / 15)))
            self.card_canvas.create_text(
                (x1 + x2) / 2, y1 + 10 * scale,
                text=f"{item['width']} мм",
                font=("Arial", font_size),
                fill="black",
                anchor="n",
                tags=("glass_label", f"label_{item['id']}")
            )

        if (y2 - y1) > 50 * scale:
            font_size = max(8, min(12, int((y2 - y1) / 15)))
            self.card_canvas.create_text(
                x1 + 15 * scale, (y1 + y2) / 2 + 5 * scale,
                text=f"{item['height']} мм",
                font=("Arial", font_size),
                fill="black",
                anchor="w",
                angle=90,
                tags=("glass_label", f"label_{item['id']}")
            )

        if (x2 - x1) > 100 * scale and (y2 - y1) > 60 * scale:
            id_parts = item['id'].split('-')
            order_num = id_parts[0] if len(id_parts) > 0 else "?"
            seq_num = id_parts[1] if len(id_parts) > 1 else "?"

            self.card_canvas.create_text(
                (x1 + x2) / 2, (y1 + y2) / 2,
                text=f"{order_num}-{seq_num}",
                font=("Arial", max(8, min(12, int(min(x2 - x1, y2 - y1) / 15)))),
                fill="black",
                tags=("glass_label", f"label_{item['id']}")
            )

    def update_info_panel(self, group, index, scale):
        """Обновляет информационную панель"""
        used_area = sum(i['width'] * i['height'] for i in group['items'])
        total_area = group['width'] * group['height']
        utilization = used_area / total_area * 100

        self.unused_label.configure(
            text=f"Карта {index + 1}/{len(self.groups)} | "
                 f"Размер: {group['width']}×{group['height']} мм | "
                 f"Использовано: {utilization:.1f}% | "
                 f"Масштаб: 1:{int(1 / scale)} | "
                 f"Сетка: 1000 мм"
        )

    # def calculate_utilization(self, group):
    #     """Вычисляет процент использования листа"""
    #     used_area = sum(item['width'] * item['height'] for item in group['items'])
    #     total_area = group['width'] * group['height']
    #     return round(used_area / total_area * 100, 1)

    def display_card_details(self, event):
        selected_index = self.card_listbox.curselection()
        if selected_index:
            self.display_cutting_plan(selected_index[0])