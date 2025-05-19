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

from customtkinter import CTkLabel, CTkEntry, CTkButton, CTkFrame, CTkScrollbar
from tkinter import messagebox

from database import (get_windows_for_production_order, get_production_orders)


class GlassCuttingTab(CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.min_rotation_angle = 90  # Минимальный шаг поворота
        self.parent = parent
        self.groups = []
        self.sheet_width = 6000  # Ширина листа стекла по умолчанию
        self.sheet_height = 6000  # Высота листа стекла по умолчанию
        self.zoom_level = 0.8

        self.packing_cache = {}
        self.combination_cache = {}

        self.selected_item = None
        self.hover_item = None
        self.selection_rect = None
        self.hover_rect = None
        self.tooltip = None
        self.side_panels_width = 200  # Начальная ширина боковых панелей

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

        # Кнопка оптимизации сверху
        self.optimize_button = CTkButton(
            self.center_frame,
            text="Запустить оптимизацию",
            command=self.optimize_cutting,
            width=200
        )
        self.optimize_button.pack(pady=5, anchor='n')

        # Холст для отображения (как было раньше)
        self.card_canvas = tk.Canvas(
            self.center_frame,
            width=600,
            height=600,
            bg="white",
            highlightthickness=0
        )
        self.card_canvas.pack(pady=10, fill=tk.BOTH, expand=True)

        # Метка для неиспользованной области (как было)
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

        # Привязка событий для нового функционала
        self.card_canvas.bind("<Motion>", self.on_canvas_hover)
        self.card_canvas.bind("<Button-1>", self.on_canvas_click)
        self.card_canvas.bind("<Leave>", self.hide_tooltip)

        # Привязываем обработчик выбора карты
        self.card_listbox.bind('<<ListboxSelect>>', self.on_card_select)
        self.select_default_card()

        self.load_orders_from_db()

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
        """Всплывающая подсказка при наведении"""
        if not self.groups or not self.card_listbox.curselection():
            self.hide_tooltip()
            return

        group = self.groups[self.card_listbox.curselection()[0]]
        scale = self.get_current_scale(group)

        # Ищем заготовку под курсором
        new_hover_item = None
        for item in group['items']:
            x1 = item['x'] * scale
            y1 = item['y'] * scale
            x2 = x1 + item['width'] * scale
            y2 = y1 + item['height'] * scale

            if x1 <= event.x <= x2 and y1 <= event.y <= y2:
                new_hover_item = item
                break

        # Обновляем hover-эффект только если курсор перешел на новый элемент или вышел с элемента
        if new_hover_item != self.hover_item:
            self.update_hover_effect(new_hover_item, scale)
            self.hover_item = new_hover_item

            # Показываем tooltip только если нашли элемент
            if new_hover_item:
                self.show_tooltip(event.x, event.y, new_hover_item)
            else:
                self.hide_tooltip()  # Скрываем подсказку если курсор на пустом месте

        # Если курсор перемещается по тому же элементу, просто обновляем позицию tooltip
        elif new_hover_item and hasattr(self, 'tooltip_bg') and self.tooltip_bg:
            # Перемещаем существующую подсказку
            self.card_canvas.coords(self.tooltip_bg, event.x + 10, event.y + 10, event.x + 150, event.y + 40)
            self.card_canvas.coords(self.tooltip_text, event.x + 15, event.y + 15)

        # Дополнительная проверка: если курсор на пустом месте и есть подсказка - скрываем
        elif not new_hover_item and (hasattr(self, 'tooltip_bg') and self.tooltip_bg):
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
        """Показывает всплывающую подсказку"""
        # Сначала скрываем предыдущую подсказку
        self.hide_tooltip()

        text = f"ID: {item['id']} | {item['width']}×{item['height']} мм"
        if item.get('rotation'):
            text += f" (повернуто)"

        # Создаем фон для tooltip сначала (чтобы текст был поверх него)
        self.tooltip_bg = self.card_canvas.create_rectangle(
            x + 10, y + 10,
            x + 150, y + 40,  # Фиксированный размер, чтобы не вычислять bbox
            fill="#FFFFE0",
            outline="#CCCCCC",
            tags="tooltip"
        )

        # Затем создаем текст
        self.tooltip_text = self.card_canvas.create_text(
            x + 15, y + 15,
            text=text,
            font=("Arial", 10),
            fill="black",
            anchor="nw",
            tags="tooltip"
        )

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
        """Упрощенный обработчик клика с гарантией выбранной карты"""
        try:
            # Гарантируем, что есть выбранная карта
            if not self.groups or not self.card_listbox.curselection():
                self.card_listbox.selection_set(0)  # Форсируем выбор первой карты

            group = self.groups[self.card_listbox.curselection()[0]]
            scale = self.get_current_scale(group)

            # Остальная логика обработки клика остается без изменений
            self.clear_selection()

            clicked_item = None
            for item in group['items']:
                x1 = item['x'] * scale
                y1 = item['y'] * scale
                x2 = x1 + item['width'] * scale
                y2 = y1 + item['height'] * scale

                if x1 <= event.x <= x2 and y1 <= event.y <= y2:
                    clicked_item = item
                    break

            if clicked_item:
                self.create_selection_rect(clicked_item, scale)
                self.selected_item = clicked_item
                self.select_order_in_list(clicked_item['id'])

        except Exception as e:
            print(f"Ошибка: {e}")
            self.clear_selection()

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
            for i in range(self.order_listbox.size()):
                if str(order_id) in self.order_listbox.get(i):
                    self.order_listbox.selection_clear(0, tk.END)
                    self.order_listbox.selection_set(i)
                    self.order_listbox.see(i)
                    break
        except Exception as e:
            print(f"Ошибка при выборе заказа: {e}")

    def get_current_scale(self, group):
        """Возвращает текущий масштаб отображения"""
        return min(
            self.card_canvas.winfo_width() / group['width'],
            self.card_canvas.winfo_height() / group['height']
        )

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
        """Многопоточная оптимизация раскроя с безопасным кэшированием"""
        try:
            self.sheet_width = int(self.entry_sheet_width.get())
            self.sheet_height = int(self.entry_sheet_height.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Неверные размеры листа стекла")
            return

        orders = self._fetch_orders()
        if not orders:
            messagebox.showwarning("Предупреждение", "Нет заказов для оптимизации.")
            return

        # Инициализация структур данных
        self.groups = []
        self.unused_elements = defaultdict(list)
        self.packing_cache = {}
        self.lock = threading.Lock()
        sheet_area = self.sheet_width * self.sheet_height
        min_fill_ratio = 0.85

        # Группировка по типам стекла
        type_groups = defaultdict(list)
        for order in orders:
            type_groups[order['type']].append(order)

        # Функция для обработки одного типа стекла
        def process_type(window_type, type_orders):
            try:
                items = [{
                    'id': f"{o['order_id']}-{o['window_id']}",
                    'width': o['width'],
                    'height': o['height'],
                    'type': o['type'],
                    'original': o
                } for o in type_orders]

                # Сортировка по убыванию площади
                items.sort(key=lambda x: x['width'] * x['height'], reverse=True)
                remaining_items = items.copy()

                while remaining_items:
                    # Создаем новую карту раскроя
                    current_sheet = {
                        'items': [],
                        'used_area': 0,
                        'type': window_type
                    }

                    # Заполняем текущую карту
                    while current_sheet['used_area'] / sheet_area < min_fill_ratio and remaining_items:
                        best_item = None
                        best_fill = current_sheet['used_area'] / sheet_area

                        # Поиск лучшего элемента для добавления
                        for i, item in enumerate(remaining_items):
                            test_items = current_sheet['items'] + [item]
                            cache_key = self._create_cache_key(test_items)

                            if cache_key in self.packing_cache:
                                packed, used_area = self.packing_cache[cache_key]
                            else:
                                packed, used_area = self._pack_items_safe(test_items)
                                with self.lock:
                                    self.packing_cache[cache_key] = (packed, used_area)

                            if used_area / sheet_area > best_fill and len(packed) == len(test_items):
                                best_fill = used_area / sheet_area
                                best_item = item
                                best_packed = packed
                                best_used_area = used_area

                                if best_fill >= min_fill_ratio:
                                    break

                        if best_item:
                            current_sheet['items'] = best_packed
                            current_sheet['used_area'] = best_used_area
                            remaining_items.remove(best_item)
                        else:
                            break

                    # Если карта содержит элементы, сохраняем ее
                    if current_sheet['items']:
                        with self.lock:
                            self._add_completed_sheet(current_sheet, window_type, sheet_area)

                        # Пытаемся добавить дополнительные элементы
                        self._try_add_remaining_items(current_sheet, remaining_items, window_type, sheet_area)

                # Сохраняем неиспользованные элементы
                with self.lock:
                    self.unused_elements[window_type] = remaining_items

            except Exception as e:
                print(f"Ошибка при обработке типа {window_type}: {e}")

        # Запуск обработки в пуле потоков
        with ThreadPoolExecutor(max_workers=4) as executor:
            futures = [executor.submit(process_type, wt, to) for wt, to in type_groups.items()]
            for future in as_completed(futures):
                future.result()  # Ожидаем завершения и обрабатываем исключения

        self._print_final_statistics()
        self.update_interface()
        self.select_default_card()
        self.load_orders_from_db()

    def _create_cache_key(self, items):
        """Создает безопасный ключ для кэша"""
        sizes = tuple(sorted((item['width'], item['height']) for item in items))
        return hash(sizes)

    def _pack_items_safe(self, items):
        """Безопасная упаковка элементов с обработкой ошибок"""
        try:
            return self.pack_items(items)
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

    def _try_add_remaining_items(self, sheet, remaining_items, window_type, sheet_area):
        """Пытается добавить оставшиеся элементы в карту"""
        added_count = 0
        temp_remaining = remaining_items.copy()

        for item in temp_remaining:
            test_items = sheet['items'] + [item]
            cache_key = self._create_cache_key(test_items)

            if cache_key in self.packing_cache:
                packed, used_area = self.packing_cache[cache_key]
            else:
                packed, used_area = self._pack_items_safe(test_items)
                with self.lock:
                    self.packing_cache[cache_key] = (packed, used_area)

            if len(packed) == len(test_items) and used_area <= sheet_area:
                sheet['items'] = packed
                sheet['used_area'] = used_area
                remaining_items.remove(item)
                added_count += 1
                print(f"[{window_type}] Добавлен дополнительный элемент {item['id']}")

        if added_count > 0:
            with self.lock:
                self.groups[-1].update({
                    'items': sheet['items'],
                    'used_area': sheet['used_area'],
                    'wasted_area': sheet_area - sheet['used_area'],
                    'fill_percentage': (sheet['used_area'] / sheet_area) * 100
                })

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

    def pack_items(self, items):
        """Надежная упаковка элементов с обработкой всех крайних случаев"""
        if not items:
            return [], 0

        packer = rectpack.newPacker(
            rotation=True,
            sort_algo=rectpack.SORT_AREA,
            pack_algo=rectpack.MaxRectsBaf,  # Используем более надежный алгоритм
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
                        'id': str(item.get('original', {}).get('order_id', '')) + "_" +
                              str(item.get('original', {}).get('window_id', str(hash(item['id']))[:10])),
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
        """Получение заказов из базы данных"""
        orders = []
        production_orders = get_production_orders()
        if not production_orders:
            return []

        for order in production_orders:
            order_id = order[0]
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


    def display_cutting_plan(self, index):
        """Отображение карты раскроя с сеткой и неиспользованными областями"""
        if not self.groups or index >= len(self.groups):
            return

        group = self.groups[index]
        self.card_canvas.delete("all")

        # Масштабирование
        canvas_width = self.card_canvas.winfo_width()
        canvas_height = self.card_canvas.winfo_height()
        scale = min(canvas_width / group['width'], canvas_height / group['height'])

        # 1. Рисуем фон (красный горошек для всего листа)
        self.card_canvas.create_rectangle(
            0, 0,
            group['width'] * scale,
            group['height'] * scale,
            fill="red", stipple="gray25", outline=""
        )

        # 2. Рисуем сетку с шагом 1000 мм
        self.draw_grid(group, scale)

        # 3. Рисуем границы листа поверх сетки
        self.card_canvas.create_rectangle(
            0, 0,
            group['width'] * scale,
            group['height'] * scale,
            outline="black", width=3
        )

        # 4. Рисуем все элементы поверх всего
        for item in group['items']:
            self.draw_glass_item(item, scale)

        # 5. Информационная панель
        self.update_info_panel(group, index, scale)

        # Восстанавливаем выделение после перерисовки
        if hasattr(self, 'selected_item') and self.selected_item:
            for item in group['items']:
                if item['id'] == self.selected_item['id']:
                    self.select_item(item, scale)
                    break

    def draw_grid(self, group, scale):
        """Рисует сетку с шагом 1000 мм"""
        grid_color = "#cccccc"
        grid_step = 1000  # Шаг сетки в мм

        # Вертикальные линии
        for x in range(0, group['width'] + grid_step, grid_step):
            x_pos = x * scale
            self.card_canvas.create_line(
                x_pos, 0,
                x_pos, group['height'] * scale,
                fill=grid_color, dash=(2, 2)
            )
            # Подписи осей X
            if x > 0:
                self.card_canvas.create_text(
                    x_pos, 10,
                    text=f"{x} мм",
                    font=("Arial", 8),
                    anchor=tk.N
                )

        # Горизонтальные линии
        for y in range(0, group['height'] + grid_step, grid_step):
            y_pos = y * scale
            self.card_canvas.create_line(
                0, y_pos,
                group['width'] * scale, y_pos,
                fill=grid_color, dash=(2, 2)
            )
            # Подписи осей Y
            if y > 0:
                self.card_canvas.create_text(
                    10, y_pos,
                    text=f"{y} мм",
                    font=("Arial", 8),
                    anchor=tk.W
                )

    def draw_glass_item(self, item, scale):
        """Рисует одну заготовку с подписью"""
        x1 = item['x'] * scale
        y1 = item['y'] * scale
        x2 = x1 + item['width'] * scale
        y2 = y1 + item['height'] * scale

        # Цвет в зависимости от поворота
        color = "#4CAF50" if item['rotation'] == 0 else "#2196F3"

        # Рисуем элемент
        self.card_canvas.create_rectangle(
            x1, y1, x2, y2,
            outline="black", fill=color, width=1
        )

        # Подпись (если элемент достаточно большой)
        if (x2 - x1) > 60 and (y2 - y1) > 40:
            text = f"{item['width']}×{item['height']}\nID:{item['id']}"
            if item['rotation'] != 0:
                text += f"\n(повернуто)"

            font_size = max(8, min(12, int(min(x2 - x1, y2 - y1) / 15)))
            self.card_canvas.create_text(
                (x1 + x2) / 2, (y1 + y2) / 2,
                text=text, font=("Arial", font_size),
                fill="black"
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
    #
    # def calculate_utilization(self, group):
    #     """Вычисляет процент использования листа"""
    #     used_area = sum(item['width'] * item['height'] for item in group['items'])
    #     total_area = group['width'] * group['height']
    #     return round(used_area / total_area * 100, 1)

    def display_card_details(self, event):
        selected_index = self.card_listbox.curselection()
        if selected_index:
            self.display_cutting_plan(selected_index[0])
