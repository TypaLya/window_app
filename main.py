from tkinter import ttk
from customtkinter import CTk

from Authorization import AuthWindow
from FrameCutter import FrameCuttingTab
from GlassCutter import GlassCuttingTab
from ProductionPlanning import ProductionPlanningTab

from database import create_database
from warehouse import WarehouseTab


class CuttingOptimizer(CTk):
    def __init__(self):
        super().__init__()
        self.title("Оптимизация раскроя заготовки")
        self.geometry("1150x800")
        self.withdraw()

        # Создаем вкладки
        self.tab_control = ttk.Notebook(self)

        # Вкладка для раскроя рамки
        self.frame_tab = FrameCuttingTab(self)
        self.tab_control.add(self.frame_tab, text="Раскрой рамки")

        # Вкладка для раскроя стекла
        self.glass_tab = GlassCuttingTab(self)
        self.tab_control.add(self.glass_tab, text="Раскрой стекла")

        # Вкладка для планирования производства
        self.planning_tab = ProductionPlanningTab(self)
        self.tab_control.add(self.planning_tab, text="Планирование производства")

        # Вкладка для планирования производства
        self.warehouse_tab = WarehouseTab(self)
        self.tab_control.add(self.warehouse_tab, text="Склад")

        self.tab_control.pack(expand=1, fill="both")


        # Привязываем обработчики событий
        self.frame_tab.card_listbox.bind('<<ListboxSelect>>', self.frame_tab.display_card_details)
        self.glass_tab.card_listbox.bind('<<ListboxSelect>>', self.glass_tab.display_card_details)

    def on_orders_updated(self):
        """Обновляем списки заказов в обеих вкладках"""
        self.frame_tab.load_orders_from_db()
        self.glass_tab.load_orders_from_db()


if __name__ == "__main__":
    create_database()
    app = CuttingOptimizer()
    auth_window = AuthWindow(app)

    auth_window.mainloop()