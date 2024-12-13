import tkinter as tk
from login import LoginScreen
from main_screen import MainScreen

class App:
    def __init__(self, root):
        self.root = root
        self.show_login()

    def show_login(self):
        self.clear_screen()
        LoginScreen(self.root, self.show_main_screen)

    def show_main_screen(self):
        self.clear_screen()
        MainScreen(self.root)

    def clear_screen(self):
        for widget in self.root.winfo_children():
            widget.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Quản lý sinh viên")
    app = App(root)
    root.mainloop()
