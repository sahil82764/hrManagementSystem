from tkinter import *
from tkinter import ttk
# from PIL import ImageTk, Image
from tkinter import messagebox
import dashboard

class ClientView:
    def __init__(self, window):
        self.window = window
        window.geometry("1366x768")
        window.resizable(0, 0)
        self.window.state('zoomed')
        window.title("ADD / VIEW Mandays")
        self.txt = "MANDAYS"
        self.color = ["#4f4e4d", "#f29844", "red2"]
        self.heading = Label(self.window, text=self.txt, font=('yu gothic ui', 30, "bold"), bg="white",
                             fg='black',
                             bd=5,
                             relief=FLAT)
        self.heading.grid(row=0, column=0, columnspan=3)