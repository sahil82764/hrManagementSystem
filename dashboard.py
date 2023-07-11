from tkinter import *
from ttkthemes import themed_tk as tk
from tkinter import messagebox
import mandaysView


class Dashboard:
    def __init__(self, window):
        self.window = window
        self.window.title("ABB ASSOCIATES")
        height = 450
        width = 900
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        self.window.resizable(0, 0)


        self.text = "HR MANAGEMENT SYSTEM"
        self.heading = Label(self.window, text=self.text, font=('Arial',15, 'bold'), bg="#009aa5", fg="black", bd=5, relief=FLAT)
        self.heading.place(x=250, y=15, width=440, height=30)

        #====================== BUTTONS =================
        # HOME
        self.homeBtn = Button(self.window, text="HOME", cursor='hand2', font=('Arial',13, 'bold'), fg="white", bg='#9a258f', activebackground='white', command=lambda: self.homeView())
        self.homeBtn.place(x=37, y=112, width=118, height=45)

        # EXIT
        self.exitBtn = Button(self.window, text="EXIT", cursor='hand2', font=('Arial',13, 'bold'), fg="white", bg='#9a258f', activebackground='white', command=self.exitCommand)
        self.exitBtn.place(x=37, y=328, width=118, height=45)

        # CLIENT
        self.clientBtn = Button(self.window, text="CLIENTS", cursor='hand2', font=('Arial',13, 'bold'), fg="white", bg='#9a258f', activebackground='white', command=lambda: self.clientView())
        self.homeBtn.place(x=37, y=167, width=118, height=45)

        # MANDAYS
        self.mandayBtn = Button(self.window, text="MANDAYS", cursor='hand2', font=('Arial',13, 'bold'), fg="white", bg='#9a258f', activebackground='white', command=lambda: self.mandayView())
        self.mandayBtn.place(x=37, y=222, width=118, height=45)

        # BILL
        self.billBtn = Button(self.window, text="BILL", cursor='hand2', font=('Arial',13, 'bold'), fg="white", bg='#9a258f', activebackground='white', command=lambda: self.billView())
        self.billBtn.place(x=37, y=276, width=118, height=45)

    def homeView(self):
        win1 = Toplevel()

    def clientView(self):
        win2 = Toplevel()

    def mandayView(self):
        win3 = Toplevel()
        mandaysView.MandaysView(win3)

    def billView(self):
        pass

    def exitCommand(self):
        exit_command = messagebox.askyesno("EXIT??", "Are you sure you want to exit?")
        if exit_command > 0:
            self.window.destroy()


def win():
    window = tk.ThemedTk()
    window.get_themes()
    window.set_theme("arc")
    Dashboard(window)
    window.mainloop()


if __name__ == '__main__':
    win()