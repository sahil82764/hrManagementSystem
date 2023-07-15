from tkinter import *
from tkinter import ttk
# from PIL import ImageTk, Image
from tkinter import messagebox
from tkinter import filedialog
import generateBill
import pandas as pd
import pandas as pd

class WageView:
    def __init__(self, window, billPath, current_month_claimed_mandays, last_month_claimed_mandays, current_month_active_mandays):
        self.window = window
        window.geometry("1366x768")
        window.resizable(0, 0)
        self.window.state('zoomed')
        window.title("WAGE DETAILS")
        self.txt = "UPLOAD WAGE RATE"
        self.color = ["#4f4e4d", "#f29844", "red2"]
        self.heading = Label(self.window, text=self.txt, font=('yu gothic ui', 30, "bold"), bg="white",
                             fg='black',
                             bd=5,
                             relief=FLAT)
        self.heading.grid(row=0, column=0, columnspan=3)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=('yu gothic ui', 10, "bold"), foreground="black",
                        background="#108cff")
        
        # ============== INPUT VARIABLES =============================
        self.billPath = billPath
        self. current_month_claimed_mandays = current_month_claimed_mandays
        self.last_month_claimed_mandays = last_month_claimed_mandays
        self.current_month_active_mandays = current_month_active_mandays

        self.dsm = StringVar()
        self.dsm_addl = StringVar()
        self.tech = StringVar()
        self.tech_addl = StringVar()
        self.mgr = StringVar()
        self.mgr_addl = StringVar()

        self.dsm_fh = StringVar()
        self.dsm_fh_addl = StringVar()
        self.tech_fh = StringVar()
        self.tech_fh_addl = StringVar()
        self.mgr_fh = StringVar()
        self.mgr_fh_addl = StringVar()

        self.dsm_nh = StringVar()
        self.dsm_nh_addl = StringVar()
        self.tech_nh = StringVar()
        self.tech_nh_addl = StringVar()
        self.mgr_nh = StringVar()
        self.mgr_nh_addl = StringVar()

        self.dsm_cl = StringVar()
        self.dsm_cl_addl = StringVar()
        self.tech_cl = StringVar()
        self.tech_cl_addl = StringVar()
        self.mgr_cl = StringVar()
        self.mgr_cl_addl = StringVar()

        self.dsm_fs = StringVar()
        self.dsm_fs_addl = StringVar()
        self.tech_fs = StringVar()
        self.tech_fs_addl = StringVar()
        self.mgr_fs = StringVar()
        self.mgr_fs_addl = StringVar()

        self.wageRateFile = StringVar()

        # =============== LABELS =================================

        self.wageRate_entry = Entry(self.window, text='WAGE RATE', textvariable=self.wageRateFile)
        self.wageRate_entry.grid(row=1, column=0, columnspan=5, padx=(50, 10), pady=10, sticky="ew")

        self.upload_btn = Button(self.window, text="Upload Wage Rate", command=lambda: self.upload_wage_rate())
        self.upload_btn.grid(row=2, column=0, columnspan=3, padx=(50, 10), pady=10)

        self.bill_btn = Button(self.window, text="Generate Bill", command=lambda: self.generate_bill())
        self.bill_btn.grid(row=3, column=0, columnspan=3, padx=(50, 10), pady=10)

    def upload_wage_rate(self):
        try:
            self.filePath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
            self.wageRate_entry.config(state='normal')
            self.wageRate_entry.delete(0, END)  # Clear previous path, if any
            self.wageRate_entry.insert(END, self.filePath)  # Display the selected path
            self.wageRate_entry.config(state='readonly')
        
        except Exception as e:
            print(e)

    def generate_bill(self):
        
        current_month_claimed_mandays_df = pd.read_excel(self.current_month_claimed_mandays)

        last_month_claimed_mandays_df = pd.read_excel(self.last_month_claimed_mandays)

        current_month_active_mandays_df = pd.read_excel(self.current_month_active_mandays)

        wage_rate_df = pd.read_excel(self.filePath)

        generateBill.createBill(self.billPath, current_month_claimed_mandays_df, last_month_claimed_mandays_df, current_month_active_mandays_df, wage_rate_df)






        


