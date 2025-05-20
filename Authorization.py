from customtkinter import CTk, CTkLabel, CTkEntry, CTkButton, CTkFrame, CTkScrollbar, CTkRadioButton
from tkinter import messagebox

from Users import check_credentials


class AuthWindow(CTk):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent
        self.geometry("300x200")
        self.title("Авторизация")
        self.protocol("WM_DELETE_WINDOW", self.on_close)  # Обрабатываем закрытие окна

        self.label_login = CTkLabel(self, text="Логин:")
        self.label_login.pack(pady=5)

        self.entry_login = CTkEntry(self, width=200)
        self.entry_login.pack(pady=5)

        self.label_password = CTkLabel(self, text="Пароль:")
        self.label_password.pack(pady=5)

        self.entry_password = CTkEntry(self, width=200, show="*")
        self.entry_password.pack(pady=5)
        self.entry_password.bind("<Return>", lambda event: self.authenticate())  # Обработка Enter в поле пароля

        self.button_login = CTkButton(self, text="Войти", command=self.authenticate)
        self.button_login.pack(pady=10)
        self.entry_login.bind("<Return>", lambda event: self.entry_password.focus())

    def authenticate(self):
        username = self.entry_login.get()
        password = self.entry_password.get()

        result = check_credentials(username, password)
        if result is True:
            self.destroy()
            self.parent.deiconify()  # Показываем главное окно
        elif result == "wrong_password":
            messagebox.showerror("Ошибка", "Неверный пароль.")
        elif result == "no_user":
            messagebox.showerror("Ошибка", "Пользователь не найден.")

    def on_close(self):
        if messagebox.askyesno("Выход", "Вы уверены, что хотите выйти?"):
            self.destroy()
            self.parent.destroy()  # Закрываем главное окно, чтобы программа завершилась корректно