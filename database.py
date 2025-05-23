import sqlite3


def safe_float(value):
    try:
        return float(value)
    except (ValueError, TypeError):
        return ""


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
        start_balance REAL,
        income REAL,
        outcome REAL,
        end_balance REAL,
        reserved REAL,
        in_transit REAL,
        price REAL,
        total_sum REAL,
        last_income_date TEXT,
        last_move_date TEXT,
        description TEXT,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # Таблица для материалов участка комплектующих
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS components_warehouse (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            num TEXT,
            name TEXT NOT NULL,
            unit TEXT NOT NULL,
            photo TEXT,
            start_balance REAL,
            income REAL,
            outcome REAL,
            end_balance REAL,
            reserved REAL,
            in_transit REAL,
            price REAL,
            total_sum REAL,
            last_income_date TEXT,
            last_move_date TEXT,
            description TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
        ''')

    # Таблица для материалов участка комплектующих
    cursor.execute('''
            CREATE TABLE IF NOT EXISTS windows_warehouse (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                num TEXT,
                name TEXT NOT NULL,
                unit TEXT NOT NULL,
                photo TEXT,
                start_balance REAL,
                income REAL,
                outcome REAL,
                end_balance REAL,
                reserved REAL,
                in_transit REAL,
                price REAL,
                total_sum REAL,
                last_income_date TEXT,
                last_move_date TEXT,
                description TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
            ''')

    # Таблица для материалов участка триплекса
    cursor.execute('''
                CREATE TABLE IF NOT EXISTS triplex_warehouse (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    num TEXT,
                    name TEXT NOT NULL,
                    unit TEXT NOT NULL,
                    photo TEXT,
                    start_balance REAL,
                    income REAL,
                    outcome REAL,
                    end_balance REAL,
                    reserved REAL,
                    in_transit REAL,
                    price REAL,
                    total_sum REAL,
                    last_income_date TEXT,
                    last_move_date TEXT,
                    description TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
                ''')

    # Таблица для материалов участка триплекса
    cursor.execute('''
                    CREATE TABLE IF NOT EXISTS main_glass_warehouse (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        num TEXT,
                        name TEXT NOT NULL,
                        unit TEXT NOT NULL,
                        photo TEXT,
                        start_balance REAL,
                        income REAL,
                        outcome REAL,
                        end_balance REAL,
                        reserved REAL,
                        in_transit REAL,
                        price REAL,
                        total_sum REAL,
                        last_income_date TEXT,
                        last_move_date TEXT,
                        description TEXT,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                    ''')

    # Таблица для материалов участка резки стекла
    cursor.execute('''
                        CREATE TABLE IF NOT EXISTS cutting_warehouse (
                            id INTEGER PRIMARY KEY AUTOINCREMENT,
                            num TEXT,
                            name TEXT NOT NULL,
                            unit TEXT NOT NULL,
                            photo TEXT,
                            start_balance REAL,
                            income REAL,
                            outcome REAL,
                            end_balance REAL,
                            reserved REAL,
                            in_transit REAL,
                            price REAL,
                            total_sum REAL,
                            description TEXT,
                            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
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
    """Добавление материала на склад пленки"""
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
    """Получение всех материалов склада пленки"""
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


def add_component_material(data):
    """Добавление материала на склад комплектующих"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    INSERT INTO components_warehouse (
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


def get_all_component_materials():
    """Получение всех материалов склада комплектующих"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    SELECT id, num, name, unit, photo, start_balance, income, outcome,
           end_balance, reserved, in_transit, price, total_sum,
           last_income_date, last_move_date, description, updated_at
    FROM components_warehouse
    ORDER BY id
    ''')

    materials = cursor.fetchall()
    conn.close()
    return materials


def add_window_material(data):
    """Добавление материала на склад стеклопакетов"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    INSERT INTO windows_warehouse (
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


def get_all_window_materials():
    """Получение всех материалов склада стеклопакетов"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    SELECT id, num, name, unit, photo, start_balance, income, outcome,
           end_balance, reserved, in_transit, price, total_sum,
           last_income_date, last_move_date, description, updated_at
    FROM windows_warehouse
    ORDER BY id
    ''')

    materials = cursor.fetchall()
    conn.close()
    return materials


def add_triplex_material(data):
    """Добавление материала на склад триплекса"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    INSERT INTO triplex_warehouse (
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


def get_all_triplex_materials():
    """Получение всех материалов склада триплекса"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    SELECT id, num, name, unit, photo, start_balance, income, outcome,
           end_balance, reserved, in_transit, price, total_sum,
           last_income_date, last_move_date, description, updated_at
    FROM triplex_warehouse
    ORDER BY id
    ''')

    materials = cursor.fetchall()
    conn.close()
    return materials


def add_main_glass_material(data):
    """Добавление материала на основной склад стекла"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    INSERT INTO main_glass_warehouse (
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


def get_all_main_glass_materials():
    """Получение всех материалов основного склада стекла"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    SELECT id, num, name, unit, photo, start_balance, income, outcome,
           end_balance, reserved, in_transit, price, total_sum,
           last_income_date, last_move_date, description, updated_at
    FROM main_glass_warehouse
    ORDER BY id
    ''')

    materials = cursor.fetchall()
    conn.close()
    return materials


def add_cutting_material(data):
    """Добавление материала на склад резки стекла"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    INSERT INTO cutting_warehouse (
        num, name, unit, photo, start_balance, income, outcome,
        end_balance, reserved, in_transit, price, total_sum, description
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
        data.get('description')
    ))

    conn.commit()
    material_id = cursor.lastrowid
    conn.close()
    return material_id


def get_all_cutting_materials():
    """Получение всех материалов склада резки стекла"""
    conn = sqlite3.connect('orders.db')
    cursor = conn.cursor()

    cursor.execute('''
    SELECT id, num, name, unit, photo, start_balance, income, outcome,
           end_balance, reserved, in_transit, price, total_sum, description, updated_at
    FROM cutting_warehouse
    ORDER BY id
    ''')

    materials = cursor.fetchall()
    conn.close()
    return materials