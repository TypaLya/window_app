from passlib.context import CryptContext
import secrets
import string

# Настройка криптографического контекста
pwd_context = CryptContext(
    schemes=["pbkdf2_sha256"],
    default="pbkdf2_sha256",
    pbkdf2_sha256__default_rounds=30000
)

# Храним пользователей в формате: {username: (salt, hashed_password)}
users = {
    "admin": {
        "salt": "random_salt_1",  # В реальной системе соль должна быть уникальной для каждого пользователя
        "hashed_password": pwd_context.hash("admin" + "random_salt_1")
    },
    "user": {
        "salt": "random_salt_2",
        "hashed_password": pwd_context.hash("password" + "random_salt_2")
    }
}


def generate_salt(length=16):
    """Генерация случайной соли"""
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(length))


def check_credentials(username, password):
    """Проверка учетных данных"""
    if not username or not password:
        return "empty_fields"

    user_data = users.get(username)
    if not user_data:
        return "no_user"

    # Проверяем пароль с учетом соли
    if pwd_context.verify(password + user_data["salt"], user_data["hashed_password"]):
        return True
    else:
        return "wrong_password"


def add_user(username, password):
    """Добавление нового пользователя (для админки)"""
    if username in users:
        return "user_exists"

    if len(password) < 8:
        return "weak_password"

    salt = generate_salt()
    hashed_password = pwd_context.hash(password + salt)
    users[username] = {
        "salt": salt,
        "hashed_password": hashed_password
    }
    return True

