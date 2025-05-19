import sqlite3

def safe_float(value):
    try:
        return float(value)
    except (ValueError, TypeError):
        return None

# Создаем базу данных и таблицу
def create_database():
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    # Создаем таблицу заказов с типом стеклопакета
    cursor.execute('''CREATE TABLE IF NOT EXISTS orders (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        width INTEGER NOT NULL,
                        height INTEGER NOT NULL,
                        package_type TEXT NOT NULL)''')

    # Создание таблицы производственных заказов
    cursor.execute("""
           CREATE TABLE IF NOT EXISTS production_orders (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               name TEXT NOT NULL,
               customer TEXT,
               deadline DATE NOT NULL,
               priority TEXT NOT NULL,
               status TEXT NOT NULL
           )
       """)

    # Создание таблицы стеклопакетов в производственных заказах
    cursor.execute("""
           CREATE TABLE IF NOT EXISTS production_order_windows (
               id INTEGER PRIMARY KEY AUTOINCREMENT,
               order_id INTEGER NOT NULL,
               type TEXT NOT NULL,
               width INTEGER NOT NULL,
               height INTEGER NOT NULL,
               quantity INTEGER NOT NULL,
               FOREIGN KEY(order_id) REFERENCES production_orders(id)
           )
       """)

    # Создание таблицы материалов для производственных заказов
    cursor.execute("""
            CREATE TABLE IF NOT EXISTS production_order_materials (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id INTEGER NOT NULL,
                type TEXT NOT NULL,
                amount REAL NOT NULL,
                dimension TEXT NOT NULL,
                FOREIGN KEY(order_id) REFERENCES production_orders(id)
            )
        """)

    # Таблица для материалов участка наклейки пленки
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS film_warehouse (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        num TEXT,
        name TEXT NOT NULL,
        unit TEXT NOT NULL,
        photo TEXT,
        start_balance REAL DEFAULT 0,
        income REAL DEFAULT 0,
        outcome REAL DEFAULT 0,
        end_balance REAL DEFAULT 0,
        reserved REAL DEFAULT 0,
        in_transit REAL DEFAULT 0,
        price REAL DEFAULT 0,
        total_sum REAL DEFAULT 0,
        last_income_date TEXT,
        last_move_date TEXT,
        description TEXT,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # Таблица истории изменений по материалам
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS film_warehouse_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            material_id INTEGER NOT NULL,
            change_type TEXT NOT NULL,  -- 'income', 'outcome', 'correction'
            amount REAL NOT NULL,
            document_ref TEXT,  -- Номер документа-основания
            notes TEXT,
            changed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY(material_id) REFERENCES film_warehouse(id)
        )
        ''')

    conn.commit()
    conn.close()

def add_order_to_db(package_type, width, height, quantity):
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    # Вставляем заказ с указанием типа стеклопакета
    cursor.execute("INSERT INTO orders (package_type, width, height, quantity) VALUES (?, ?, ?, ?)",
                   (package_type, width, height, quantity))
    conn.commit()
    conn.close()

def delete_order_from_db(order_id):
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM orders WHERE id = ?", (order_id,))
    conn.commit()
    conn.close()

def update_order_in_db(order_id, new_width, new_height):
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()
    cursor.execute("UPDATE orders SET width = ?, height = ? WHERE id = ?", (new_width, new_height, order_id))
    conn.commit()
    conn.close()

def get_all_orders_from_db():
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    # Добавляем выбор типа стеклопакета
    cursor.execute("SELECT id, width, height, package_type FROM orders")
    orders = cursor.fetchall()
    conn.close()
    return orders


def add_production_order(name, customer, deadline, priority, status):
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()
    cursor.execute("""
            INSERT INTO production_orders (name, customer, deadline, priority, status)
            VALUES (?, ?, ?, ?, ?)
        """, (name, customer, deadline, priority, status))
    conn.commit()
    conn.close()
    return cursor.lastrowid  # <- Это критически важно!


def get_production_orders():
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM production_orders ORDER BY deadline")
    orders = cursor.fetchall()
    conn.close()
    return orders

def update_production_order_status(order_id, new_status):
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE production_orders 
        SET status = ?
        WHERE id = ?
    """, (new_status, order_id))
    conn.commit()
    conn.close()

def delete_production_order(order_id):
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM production_orders WHERE id = ?", (order_id,))
    conn.commit()
    conn.close()

def add_window_to_production_order(order_id, window_type, width, height, quantity):
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO production_order_windows (order_id, type, width, height, quantity)
        VALUES (?, ?, ?, ?, ?)
    """, (order_id, window_type, width, height, quantity))
    conn.commit()
    conn.close()

def get_windows_for_production_order(order_id):
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, type, width, height, quantity 
        FROM production_order_windows 
        WHERE order_id = ?
        ORDER BY id
    """, (order_id,))
    windows = cursor.fetchall()
    conn.close()
    return windows

def delete_window_from_production_order(window_id):
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM production_order_windows WHERE id = ?", (window_id,))
    conn.commit()
    conn.close()

def add_material_to_production_order(order_id, material_type, amount, dimension):
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO production_order_materials (order_id, type, amount, dimension)
        VALUES (?, ?, ?, ?)
    """, (order_id, material_type, amount, dimension))
    conn.commit()
    conn.close()

def get_materials_for_production_order(order_id):
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()
    cursor.execute("""
        SELECT id, type, amount, dimension 
        FROM production_order_materials 
        WHERE order_id = ?
        ORDER BY id
    """, (order_id,))
    materials = cursor.fetchall()
    conn.close()
    return materials

def delete_material_from_production_order(material_id):
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()
    cursor.execute("DELETE FROM production_order_materials WHERE id = ?", (material_id,))
    conn.commit()
    conn.close()


def add_film_material(data):
    """Добавление материала на склад пленки (только поля из Excel)"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    INSERT INTO film_warehouse (
        num, name, unit, photo, start_balance, income, outcome,
        end_balance, reserved, in_transit, price, total_sum,
        last_income_date, last_move_date, description
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        data.get('num'),
        data['name'],
        data['unit'],
        data.get('photo'),
        safe_float(data.get('start_balance', 0)),
        safe_float(data.get('income', 0)),
        safe_float(data.get('outcome', 0)),
        safe_float(data.get('end_balance', 0)),
        safe_float(data.get('reserved', 0)),
        safe_float(data.get('in_transit', 0)),
        safe_float(data.get('price', 0)),
        safe_float(data.get('total_sum', 0)),
        data.get('last_income_date'),
        data.get('last_move_date'),
        data.get('description')
    ))

    conn.commit()
    material_id = cursor.lastrowid
    conn.close()
    return material_id


def get_all_film_materials():
    """Получение всех материалов склада пленки (только нужные поля)"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    SELECT id, num, name, unit, photo, start_balance, income, outcome,
           end_balance, reserved, in_transit, price, total_sum,
           last_income_date, last_move_date, description, updated_at
    FROM film_warehouse
    ORDER BY id
    ''')

    materials = cursor.fetchall()
    conn.close()
    return materials


def update_film_material(material_id, data):
    """Обновление данных материала (только исходные поля)"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    UPDATE film_warehouse SET
        num = ?,
        name = ?,
        unit = ?,
        photo = ?,
        start_balance = ?,
        income = ?,
        outcome = ?,
        end_balance = ?,
        reserved = ?,
        in_transit = ?,
        price = ?,
        total_sum = ?,
        last_income_date = ?,
        last_move_date = ?,
        description = ?,
        updated_at = CURRENT_TIMESTAMP
    WHERE id = ?
    ''', (
        data.get('num'),
        data['name'],
        data['unit'],
        data.get('photo'),
        safe_float(data.get('start_balance', 0)),
        safe_float(data.get('income', 0)),
        safe_float(data.get('outcome', 0)),
        safe_float(data.get('end_balance', 0)),
        safe_float(data.get('reserved', 0)),
        safe_float(data.get('in_transit', 0)),
        safe_float(data.get('price', 0)),
        safe_float(data.get('total_sum', 0)),
        data.get('last_income_date'),
        data.get('last_move_date'),
        data.get('description'),
        material_id
    ))

    conn.commit()
    conn.close()

