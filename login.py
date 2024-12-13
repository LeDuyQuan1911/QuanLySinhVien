import tkinter as tk
from tkinter import messagebox

class LoginScreen:
    def __init__(self, root, on_login_success):
        self.root = root
        self.on_login_success = on_login_success
        

        self.root.title("Login Screen")
        self.root.geometry("400x300")
        self.root.config(bg="#f0f0f0")  

        self.frame = tk.Frame(root, bg="#ffffff", bd=5, relief="groove") 
        self.frame.pack(padx=20, pady=20)

        label_font = ("Helvetica", 12)

        tk.Label(self.frame, text="Username", bg="#ffffff", font=label_font).grid(row=0, column=0, pady=10)
        tk.Label(self.frame, text="Password", bg="#ffffff", font=label_font).grid(row=1, column=0, pady=10)

        self.username_entry = tk.Entry(self.frame, width=25)
        self.password_entry = tk.Entry(self.frame, width=25, show="*")
        self.username_entry.grid(row=0, column=1, padx=10, pady=10)
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)

        self.login_button = tk.Button(self.frame, text="Login", command=self.check_login, 
                                      bg="#4CAF50", fg="white", font=("Helvetica", 12, "bold"))
        self.signup_button = tk.Button(self.frame, text="Sign up", command=self.check_login, 
                                      bg="#4CAF50", fg="white", font=("Helvetica", 12, "bold"))
        self.login_button.grid(row=2, column=0, columnspan=2, pady=10,sticky="e")
        self.signup_button.grid(row=3, column=0, columnspan=2, pady=10,sticky="e")

    def check_login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        if username == "leduyquan" and password == "123":  
            messagebox.showinfo("Login", "Login successful!")
            self.on_login_success()  
        else:
            messagebox.showerror("Login", "Invalid credentials!")