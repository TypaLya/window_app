users = {
        "admin": "admin",
        "user1": "1234",
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