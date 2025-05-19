users = {
        "admin": "admin",
        "": "",
        "user": "password"
    }

# Заглушка для проверки логина и пароля
def check_credentials(username, password):
    if username in users:
        if users[username] == password:
            return True
        else:
            return "wrong_password"
    return "no_user"



    # cursor.execute("""
    #             CREATE TABLE IF NOT EXISTS film_warehouse (
    #                 id INTEGER PRIMARY KEY AUTOINCREMENT,
    #                 num REAL NOT NULL,
    #                 name TEXT NOT NULL,
    #                 unit TEXT NOT NULL,
    #                 photo TEXT,
    #                 start_balance REAL,
    #                 income REAL,
    #                 outcome REAL,
    #                 end_balance REAL,
    #                 reserved REAL,
    #                 in_transit REAL,
    #                 price REAL,
    #                 total_sum REAL,
    #                 last_income_date DATE,
    #                 last_move_date DATE,
    #                 description TEXT
    #             )
    #         """)