from tkinter import *
from tkinter import ttk
# from PIL import ImageTk, Image
from tkinter import messagebox
from tkinter import filedialog
from util import util
from openpyxl import load_workbook
import generateBill
import pandas as pd
import dashboard


class WageView:
    def __init__(self, window, billPath, current_month_claimed_mandays, last_month_claimed_mandays, current_month_active_mandays, lastMonth, billMonth, lastYear, billYear):
        self.window = window
        window.geometry("1366x768")
        window.resizable(0, 0)
        self.window.state('zoomed')
        window.title("WAGE DETAILS")
        self.txt = "WAGE RATE"
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
        self.lastMonth = lastMonth
        self.billMonth = billMonth
        self.lastYear = lastYear
        self.billYear = billYear

        self.wageRateFile = StringVar()

        # ==========================================================================

        self.get_wage_entries()

        # =============== HEADER LABELS =================================

        self.particular_label = Label(self.window, text="PARTICULAR")
        self.particular_label.grid(row=1, column=0, padx=(50, 10), pady=10)

        self.previousMonthWage_label = Label(self.window, text="PREVIOUS MONTH WAGE")
        self.previousMonthWage_label.grid(row=1, column=1, padx=(50, 10), pady=10)

        self.currentMonthWage_label = Label(self.window, text="CURRENT MONTH WAGE")
        self.currentMonthWage_label.grid(row=1, column=2, padx=(50, 10), pady=10)

        # =============== PARTICULAR LABELS =================================

        self.dsm_label = Label(self.window, text="DSM")
        self.dsm_label.grid(row=2, column=0, padx=(50, 10), pady=10)

        self.tech_label = Label(self.window, text="TECH")
        self.tech_label.grid(row=3, column=0, padx=(50, 10), pady=10)

        self.mgr_label = Label(self.window, text="MANAGER")
        self.mgr_label.grid(row=4, column=0, padx=(50, 10), pady=10)

        self.dsm_fh_label = Label(self.window, text="DSM FH")
        self.dsm_fh_label.grid(row=5, column=0, padx=(50, 10), pady=10)

        self.tech_fh_label = Label(self.window, text="TECH FH")
        self.tech_fh_label.grid(row=6, column=0, padx=(50, 10), pady=10)

        self.mgr_fh_label = Label(self.window, text="MANAGER FH")
        self.mgr_fh_label.grid(row=7, column=0, padx=(50, 10), pady=10)

        self.dsm_nh_label = Label(self.window, text="DSM NH")
        self.dsm_nh_label.grid(row=8, column=0, padx=(50, 10), pady=10)

        self.tech_nh_label = Label(self.window, text="TECH NH")
        self.tech_nh_label.grid(row=9, column=0, padx=(50, 10), pady=10)

        self.mgr_nh_label = Label(self.window, text="MANAGER NH")
        self.mgr_nh_label.grid(row=10, column=0, padx=(50, 10), pady=10)

        self.dsm_cl_label = Label(self.window, text="DSM CL")
        self.dsm_cl_label.grid(row=11, column=0, padx=(50, 10), pady=10)

        self.tech_cl_label = Label(self.window, text="TECH CL")
        self.tech_cl_label.grid(row=12, column=0, padx=(50, 10), pady=10)

        self.mgr_cl_label = Label(self.window, text="MANAGER CL")
        self.mgr_cl_label.grid(row=13, column=0, padx=(50, 10), pady=10)

        self.dsm_ft_label = Label(self.window, text="DSM FT")
        self.dsm_ft_label.grid(row=14, column=0, padx=(50, 10), pady=10)

        self.tech_ft_label = Label(self.window, text="TECH FT")
        self.tech_ft_label.grid(row=15, column=0, padx=(50, 10), pady=10)

        self.mgr_ft_label = Label(self.window, text="MANAGER FT")
        self.mgr_ft_label.grid(row=16, column=0, padx=(50, 10), pady=10)

        # =============== PREVIOUS MONTH ENTRIES =================================

        self.dsm_pm_entry = Entry(self.window, width=20, textvariable=self.dsm_pm, state='readonly')
        self.dsm_pm_entry.grid(row=2, column=1, padx=(50, 10), pady=10)

        self.tech_pm_entry = Entry(self.window, width=20, textvariable=self.tech_pm)
        # self.tech_pm_entry.insert(0, self.tech_pm.get())
        self.tech_pm_entry.config(state='readonly')
        self.tech_pm_entry.grid(row=3, column=1, padx=(50, 10), pady=10)

        self.mgr_pm_entry = Entry(self.window, width=20, textvariable=self.mgr_pm)
        # self.mgr_pm_entry.insert(0, self.mgr_pm.get())
        self.mgr_pm_entry.config(state='readonly')
        self.mgr_pm_entry.grid(row=4, column=1, padx=(50, 10), pady=10)

        self.dsm_fh_pm_entry = Entry(self.window, width=20, textvariable=self.dsm_fh_pm)
        # self.dsm_fh_pm_entry.insert(0, self.dsm_fh_pm.get())
        self.dsm_fh_pm_entry.config(state='readonly')
        self.dsm_fh_pm_entry.grid(row=5, column=1, padx=(50, 10), pady=10)

        self.tech_fh_pm_entry = Entry(self.window, width=20, textvariable=self.tech_fh_pm)
        # self.tech_fh_pm_entry.insert(0, self.tech_fh_pm.get())
        self.tech_fh_pm_entry.config(state='readonly')
        self.tech_fh_pm_entry.grid(row=6, column=1, padx=(50, 10), pady=10)

        self.mgr_fh_pm_entry = Entry(self.window, width=20, textvariable=self.mgr_fh_pm)
        # self.mgr_fh_pm_entry.insert(0, self.mgr_fh_pm.get())
        self.mgr_fh_pm_entry.config(state='readonly')
        self.mgr_fh_pm_entry.grid(row=7, column=1, padx=(50, 10), pady=10)

        self.dsm_nh_pm_entry = Entry(self.window, width=20, textvariable=self.dsm_nh_pm)
        # self.dsm_nh_pm_entry.insert(0, self.dsm_nh_pm.get())
        self.dsm_nh_pm_entry.config(state='readonly')
        self.dsm_nh_pm_entry.grid(row=8, column=1, padx=(50, 10), pady=10)

        self.tech_nh_pm_entry = Entry(self.window, width=20, textvariable=self.tech_nh_pm)
        # self.tech_nh_pm_entry.insert(0, self.tech_nh_pm.get())
        self.tech_nh_pm_entry.config(state='readonly')
        self.tech_nh_pm_entry.grid(row=9, column=1, padx=(50, 10), pady=10)

        self.mgr_nh_pm_entry = Entry(self.window, width=20, textvariable=self.mgr_nh_pm)
        # self.mgr_nh_pm_entry.insert(0, self.mgr_nh_pm.get())
        self.mgr_nh_pm_entry.config(state='readonly')
        self.mgr_nh_pm_entry.grid(row=10, column=1, padx=(50, 10), pady=10)

        self.dsm_cl_pm_entry = Entry(self.window, width=20, textvariable=self.dsm_cl_pm)
        # self.dsm_cl_pm_entry.insert(0, self.dsm_cl_pm.get())
        self.dsm_cl_pm_entry.config(state='readonly')
        self.dsm_cl_pm_entry.grid(row=11, column=1, padx=(50, 10), pady=10)

        self.tech_cl_pm_entry = Entry(self.window, width=20, textvariable=self.tech_cl_pm)
        # self.tech_cl_pm_entry.insert(0, self.tech_cl_pm.get())
        self.tech_cl_pm_entry.config(state='readonly')
        self.tech_cl_pm_entry.grid(row=12, column=1, padx=(50, 10), pady=10)

        self.mgr_cl_pm_entry = Entry(self.window, width=20, textvariable=self.mgr_cl_pm)
        # self.mgr_cl_pm_entry.insert(0, self.mgr_cl_pm.get())
        self.mgr_cl_pm_entry.config(state='readonly')
        self.mgr_cl_pm_entry.grid(row=13, column=1, padx=(50, 10), pady=10)

        self.dsm_ft_pm_entry = Entry(self.window, width=20, textvariable=self.dsm_fs_pm)
        # self.dsm_ft_pm_entry.insert(0, self.dsm_fs_pm.get())
        self.dsm_ft_pm_entry.config(state='readonly')
        self.dsm_ft_pm_entry.grid(row=14, column=1, padx=(50, 10), pady=10)

        self.tech_ft_pm_entry = Entry(self.window, width=20, textvariable=self.tech_fs_pm)
        # self.tech_ft_pm_entry.insert(0, self.tech_fs_pm.get())
        self.tech_ft_pm_entry.config(state='readonly')
        self.tech_ft_pm_entry.grid(row=15, column=1, padx=(50, 10), pady=10)

        self.mgr_ft_pm_entry = Entry(self.window, width=20, textvariable=self.mgr_fs_pm)
        # self.mgr_ft_pm_entry.insert(0, self.mgr_fs_pm.get())
        self.mgr_ft_pm_entry.config(state='readonly')
        self.mgr_ft_pm_entry.grid(row=16, column=1, padx=(50, 10), pady=10)

        # =============== CURRENT MONTH ENTRIES =================================

        self.dsm_cm_entry = Entry(self.window, width=20, textvariable=self.dsm_cm)
        # self.dsm_cm_entry.insert(0, self.dsm_cm.get())
        self.dsm_cm_entry.config(state='readonly')
        self.dsm_cm_entry.grid(row=2, column=2, padx=(50, 10), pady=10)

        self.tech_cm_entry = Entry(self.window, width=20, textvariable=self.tech_cm)
        # self.tech_cm_entry.insert(0, self.tech_cm.get())
        self.tech_cm_entry.config(state='readonly')
        self.tech_cm_entry.grid(row=3, column=2, padx=(50, 10), pady=10)

        self.mgr_cm_entry = Entry(self.window, width=20, textvariable=self.mgr_cm)
        # self.mgr_cm_entry.insert(0, self.mgr_cm.get())
        self.mgr_cm_entry.config(state='readonly')
        self.mgr_cm_entry.grid(row=4, column=2, padx=(50, 10), pady=10)

        self.dsm_fh_cm_entry = Entry(self.window, width=20, textvariable=self.dsm_fh_cm)
        # self.dsm_fh_cm_entry.insert(0, self.dsm_fh_cm.get())
        self.dsm_fh_cm_entry.config(state='readonly')
        self.dsm_fh_cm_entry.grid(row=5, column=2, padx=(50, 10), pady=10)

        self.tech_fh_cm_entry = Entry(self.window, width=20, textvariable=self.tech_fh_cm)
        # self.tech_fh_cm_entry.insert(0, self.tech_fh_cm.get())
        self.tech_fh_cm_entry.config(state='readonly')
        self.tech_fh_cm_entry.grid(row=6, column=2, padx=(50, 10), pady=10)

        self.mgr_fh_cm_entry = Entry(self.window, width=20, textvariable=self.mgr_fh_cm)
        # self.mgr_fh_cm_entry.insert(0, self.mgr_fh_cm.get())
        self.mgr_fh_cm_entry.config(state='readonly')
        self.mgr_fh_cm_entry.grid(row=7, column=2, padx=(50, 10), pady=10)

        self.dsm_nh_cm_entry = Entry(self.window, width=20, textvariable=self.dsm_nh_cm)
        # self.dsm_nh_cm_entry.insert(0, self.dsm_nh_cm.get())
        self.dsm_nh_cm_entry.config(state='readonly')
        self.dsm_nh_cm_entry.grid(row=8, column=2, padx=(50, 10), pady=10)

        self.tech_nh_cm_entry = Entry(self.window, width=20, textvariable=self.tech_nh_cm)
        # self.tech_nh_cm_entry.insert(0, self.tech_nh_cm.get())
        self.tech_nh_cm_entry.config(state='readonly')
        self.tech_nh_cm_entry.grid(row=9, column=2, padx=(50, 10), pady=10)

        self.mgr_nh_cm_entry = Entry(self.window, width=20, textvariable=self.mgr_nh_cm)
        # self.mgr_nh_cm_entry.insert(0, self.mgr_nh_cm.get())
        self.mgr_nh_cm_entry.config(state='readonly')
        self.mgr_nh_cm_entry.grid(row=10, column=2, padx=(50, 10), pady=10)

        self.dsm_cl_cm_entry = Entry(self.window, width=20, textvariable=self.dsm_cl_cm)
        # self.dsm_cl_cm_entry.insert(0, self.dsm_cl_cm.get())
        self.dsm_cl_cm_entry.config(state='readonly')
        self.dsm_cl_cm_entry.grid(row=11, column=2, padx=(50, 10), pady=10)

        self.tech_cl_cm_entry = Entry(self.window, width=20, textvariable=self.tech_cl_cm)
        # self.tech_cl_cm_entry.insert(0, self.tech_cl_cm.get())
        self.tech_cl_cm_entry.config(state='readonly')
        self.tech_cl_cm_entry.grid(row=12, column=2, padx=(50, 10), pady=10)

        self.mgr_cl_cm_entry = Entry(self.window, width=20, textvariable=self.mgr_cl_cm)
        # self.mgr_cl_cm_entry.insert(0, self.mgr_cl_cm.get())
        self.mgr_cl_cm_entry.config(state='readonly')
        self.mgr_cl_cm_entry.grid(row=13, column=2, padx=(50, 10), pady=10)

        self.dsm_ft_cm_entry = Entry(self.window, width=20, textvariable=self.dsm_fs_cm)
        # self.dsm_ft_cm_entry.insert(0, self.dsm_fs_cm.get())
        self.dsm_ft_cm_entry.config(state='readonly')
        self.dsm_ft_cm_entry.grid(row=14, column=2, padx=(50, 10), pady=10)

        self.tech_ft_cm_entry = Entry(self.window, width=20, textvariable=self.tech_fs_cm)
        # self.tech_ft_cm_entry.insert(0, self.tech_fs_cm.get())
        self.tech_ft_cm_entry.config(state='readonly')
        self.tech_ft_cm_entry.grid(row=15, column=2, padx=(50, 10), pady=10)

        self.mgr_ft_cm_entry = Entry(self.window, width=20, textvariable=self.mgr_fs_cm)
        # self.mgr_ft_cm_entry.insert(0, self.mgr_fs_cm.get())
        self.mgr_ft_cm_entry.config(state='readonly')
        self.mgr_ft_cm_entry.grid(row=16, column=2, padx=(50, 10), pady=10)

        
        # Create the "Change Entries" button
        self.change_btn = Button(self.window, text="Change Entries", command=lambda: self.toggleEntry())
        self.change_btn.grid(row=17, column=0, columnspan=2, padx=(50, 10), pady=10, sticky="ew")

        # Create the "Save Entries" button
        self.save_btn = Button(self.window, text="Save Entries", command=lambda: self.saveEntry(), state='disabled')
        self.save_btn.grid(row=17, column=1, columnspan=2, padx=(50, 10), pady=10, sticky="ew")

        # Create the "Generate Bill" button
        self.bill_btn = Button(self.window, text="Generate Bill", command=lambda: self.generate_bill())
        self.bill_btn.grid(row=19, column=1, columnspan=6, padx=(50, 10), pady=10, sticky="ew")

        for i in range(2, 18):
            self.label = Label(self.window, text=" || ")
            self.label.grid(row=i, column=4, padx=(50, 10), pady=10)

        self.wageRate_entry = Entry(self.window, textvariable=self.wageRateFile, width=60)
        self.wageRate_entry.grid(row=8, column=5, columnspan=5, padx=(50, 10), pady=10, sticky="ew")

        self.upload_btn = Button(self.window, text="Upload Wage Rate", command=lambda: self.upload_wage_rate(), width= 50)
        self.upload_btn.grid(row=9, column=5, padx=(50, 10), pady=10)


    def get_wage_entries(self):

        self.wageWorkbook = load_workbook(util.get_wage_template())
        self.wage_sheet = self.wageWorkbook.active

        # ============== PREVIOUS MONTH WAGE VARIABLES =============================

        self.dsm_pm = IntVar()
        self.dsm_pm.set(self.wage_sheet['B2'].value)
        self.tech_pm = IntVar()
        self.tech_pm.set(self.wage_sheet['B4'].value)
        self.mgr_pm = IntVar()
        self.mgr_pm.set(self.wage_sheet['B6'].value)
        
        self.dsm_fh_pm = IntVar()
        self.dsm_fh_pm.set(self.wage_sheet['B8'].value)
        self.tech_fh_pm = IntVar()
        self.tech_fh_pm.set(self.wage_sheet['B10'].value)
        self.mgr_fh_pm = IntVar()
        self.mgr_fh_pm.set(self.wage_sheet['B12'].value)

        self.dsm_nh_pm = IntVar()
        self.dsm_nh_pm.set(self.wage_sheet['B14'].value)
        self.tech_nh_pm = IntVar()
        self.tech_nh_pm.set(self.wage_sheet['B16'].value)
        self.mgr_nh_pm = IntVar()
        self.mgr_nh_pm.set(self.wage_sheet['B18'].value)
        
        self.dsm_cl_pm = IntVar()
        self.dsm_cl_pm.set(self.wage_sheet['B20'].value)
        self.tech_cl_pm = IntVar()
        self.tech_cl_pm.set(self.wage_sheet['B22'].value)
        self.mgr_cl_pm = IntVar()
        self.mgr_cl_pm.set(self.wage_sheet['B24'].value)

        self.dsm_fs_pm = IntVar()
        self.dsm_fs_pm.set(self.wage_sheet['B26'].value)
        self.tech_fs_pm = IntVar()
        self.tech_fs_pm.set(self.wage_sheet['B28'].value)    
        self.mgr_fs_pm = IntVar()
        self.mgr_fs_pm.set(self.wage_sheet['B30'].value)

        # ============== CURRENT MONTH WAGE VARIABLES =============================

        self.dsm_cm = IntVar()
        self.dsm_cm.set(self.wage_sheet['C2'].value)
        self.tech_cm = IntVar()
        self.tech_cm.set(self.wage_sheet['C4'].value)
        self.mgr_cm = IntVar()
        self.mgr_cm.set(self.wage_sheet['C6'].value)
        
        self.dsm_fh_cm = IntVar()
        self.dsm_fh_cm.set(self.wage_sheet['C8'].value)
        self.tech_fh_cm = IntVar()
        self.tech_fh_cm.set(self.wage_sheet['C10'].value)
        self.mgr_fh_cm = IntVar()
        self.mgr_fh_cm.set(self.wage_sheet['C12'].value)

        self.dsm_nh_cm = IntVar()
        self.dsm_nh_cm.set(self.wage_sheet['C14'].value)
        self.tech_nh_cm = IntVar()
        self.tech_nh_cm.set(self.wage_sheet['C16'].value)
        self.mgr_nh_cm = IntVar()
        self.mgr_nh_cm.set(self.wage_sheet['C18'].value)
        
        self.dsm_cl_cm = IntVar()
        self.dsm_cl_cm.set(self.wage_sheet['C20'].value)
        self.tech_cl_cm = IntVar()
        self.tech_cl_cm.set(self.wage_sheet['C22'].value)
        self.mgr_cl_cm = IntVar()
        self.mgr_cl_cm.set(self.wage_sheet['C24'].value)

        self.dsm_fs_cm = IntVar()
        self.dsm_fs_cm.set(self.wage_sheet['C26'].value)
        self.tech_fs_cm = IntVar()
        self.tech_fs_cm.set(self.wage_sheet['C28'].value)    
        self.mgr_fs_cm = IntVar()
        self.mgr_fs_cm.set(self.wage_sheet['C30'].value)



    def toggleEntry(self):

        self.dsm_pm_entry.config(state='normal')
        self.tech_pm_entry.config(state='normal')
        self.mgr_pm_entry.config(state='normal')
        self.dsm_fh_pm_entry.config(state='normal')
        self.tech_fh_pm_entry.config(state='normal')
        self.mgr_fh_pm_entry.config(state='normal')
        self.dsm_nh_pm_entry.config(state='normal')
        self.tech_nh_pm_entry.config(state='normal')
        self.mgr_nh_pm_entry.config(state='normal')
        self.dsm_cl_pm_entry.config(state='normal')
        self.tech_cl_pm_entry.config(state='normal')
        self.mgr_cl_pm_entry.config(state='normal')
        self.dsm_ft_pm_entry.config(state='normal')
        self.tech_ft_pm_entry.config(state='normal')
        self.mgr_ft_pm_entry.config(state='normal')

        self.dsm_cm_entry.config(state='normal')
        self.tech_cm_entry.config(state='normal')
        self.mgr_cm_entry.config(state='normal')
        self.dsm_fh_cm_entry.config(state='normal')
        self.tech_fh_cm_entry.config(state='normal')
        self.mgr_fh_cm_entry.config(state='normal')
        self.dsm_nh_cm_entry.config(state='normal')
        self.tech_nh_cm_entry.config(state='normal')
        self.mgr_nh_cm_entry.config(state='normal')
        self.dsm_cl_cm_entry.config(state='normal')
        self.tech_cl_cm_entry.config(state='normal')
        self.mgr_cl_cm_entry.config(state='normal')
        self.dsm_ft_cm_entry.config(state='normal')
        self.tech_ft_cm_entry.config(state='normal')
        self.mgr_ft_cm_entry.config(state='normal')

        # Toggle the button text between "Change" and "Save"
        self.change_btn.config(state='disabled')
        self.save_btn.config(state='normal')


    def saveEntry(self):        
        if (

            self.dsm_pm.get() and
            self.tech_pm.get() and
            self.mgr_pm.get() and
            self.dsm_fh_pm.get() and
            self.tech_fh_pm.get() and
            self.mgr_fh_pm.get() and
            self.dsm_nh_pm.get() and
            self.tech_nh_pm.get() and
            self.mgr_nh_pm.get() and
            self.dsm_cl_pm.get() and
            self.tech_cl_pm.get() and
            self.mgr_cl_pm.get() and
            self.dsm_fs_pm.get() and
            self.tech_fs_pm.get() and
            self.mgr_fs_pm.get() and
            self.dsm_cm.get() and
            self.tech_cm.get() and
            self.mgr_cm.get() and
            self.dsm_fh_cm.get() and
            self.tech_fh_cm.get() and
            self.mgr_fh_cm.get() and
            self.dsm_nh_cm.get() and
            self.tech_nh_cm.get() and
            self.mgr_nh_cm.get() and
            self.dsm_cl_cm.get() and
            self.tech_cl_cm.get() and
            self.mgr_cl_cm.get() and
            self.dsm_fs_cm.get() and
            self.tech_fs_cm.get() and
            self.mgr_fs_cm.get()
            
        ):
            
            self.wage_sheet['B2'] = self.dsm_pm.get()
            self.wage_sheet['B4'] = self.tech_pm.get()
            self.wage_sheet['B6'] = self.mgr_pm.get()
            self.wage_sheet['B8'] = self.dsm_fh_pm.get()
            self.wage_sheet['B10'] = self.tech_fh_pm.get()
            self.wage_sheet['B12'] = self.mgr_fh_pm.get()
            self.wage_sheet['B14'] = self.dsm_nh_pm.get()
            self.wage_sheet['B16'] = self.tech_nh_pm.get()
            self.wage_sheet['B18'] = self.mgr_nh_pm.get()
            self.wage_sheet['B20'] = self.dsm_cl_pm.get()
            self.wage_sheet['B22'] = self.tech_cl_pm.get()
            self.wage_sheet['B24'] = self.mgr_cl_pm.get()
            self.wage_sheet['B26'] = self.dsm_fs_pm.get()
            self.wage_sheet['B28'] = self.tech_fs_pm.get()
            self.wage_sheet['B30'] = self.mgr_fs_pm.get()

            self.wage_sheet['C2'] = self.dsm_cm.get()
            self.wage_sheet['C4'] = self.tech_cm.get()
            self.wage_sheet['C6'] = self.mgr_cm.get()
            self.wage_sheet['C8'] = self.dsm_fh_cm.get()
            self.wage_sheet['C10'] = self.tech_fh_cm.get()
            self.wage_sheet['C12'] = self.mgr_fh_cm.get()
            self.wage_sheet['C14'] = self.dsm_nh_cm.get()
            self.wage_sheet['C16'] = self.tech_nh_cm.get()
            self.wage_sheet['C18'] = self.mgr_nh_cm.get()
            self.wage_sheet['C20'] = self.dsm_cl_cm.get()
            self.wage_sheet['C22'] = self.tech_cl_cm.get()
            self.wage_sheet['C24'] = self.mgr_cl_cm.get()
            self.wage_sheet['C26'] = self.dsm_fs_cm.get()
            self.wage_sheet['C28'] = self.tech_fs_cm.get()
            self.wage_sheet['C30'] = self.mgr_fs_cm.get()

            self.wageWorkbook.save(util.get_wage_template())

            self.dsm_pm_entry.config(state='readonly')
            self.tech_pm_entry.config(state='readonly')
            self.mgr_pm_entry.config(state='readonly')
            self.dsm_fh_pm_entry.config(state='readonly')
            self.tech_fh_pm_entry.config(state='readonly')
            self.mgr_fh_pm_entry.config(state='readonly')
            self.dsm_nh_pm_entry.config(state='readonly')
            self.tech_nh_pm_entry.config(state='readonly')
            self.mgr_nh_pm_entry.config(state='readonly')
            self.dsm_cl_pm_entry.config(state='readonly')
            self.tech_cl_pm_entry.config(state='readonly')
            self.mgr_cl_pm_entry.config(state='readonly')
            self.dsm_ft_pm_entry.config(state='readonly')
            self.tech_ft_pm_entry.config(state='readonly')
            self.mgr_ft_pm_entry.config(state='readonly')

            self.dsm_cm_entry.config(state='readonly')
            self.tech_cm_entry.config(state='readonly')
            self.mgr_cm_entry.config(state='readonly')
            self.dsm_fh_cm_entry.config(state='readonly')
            self.tech_fh_cm_entry.config(state='readonly')
            self.mgr_fh_cm_entry.config(state='readonly')
            self.dsm_nh_cm_entry.config(state='readonly')
            self.tech_nh_cm_entry.config(state='readonly')
            self.mgr_nh_cm_entry.config(state='readonly')
            self.dsm_cl_cm_entry.config(state='readonly')
            self.tech_cl_cm_entry.config(state='readonly')
            self.mgr_cl_cm_entry.config(state='readonly')
            self.dsm_ft_cm_entry.config(state='readonly')
            self.tech_ft_cm_entry.config(state='readonly')
            self.mgr_ft_cm_entry.config(state='readonly')

            self.change_btn.config(state='normal')
            self.save_btn.config(state='disabled')

        else:
            # Display an error message if any field is empty
            messagebox.showerror("Error", "Please fill in all the fields.")


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

        if self.filePath is not None:
            wage_rate_df = pd.read_excel(self.filePath)
        else:
            wage_rate_df = pd.read_excel(util.get_wage_template())

        generateBill.createBill(self.billPath, current_month_claimed_mandays_df, last_month_claimed_mandays_df, current_month_active_mandays_df, wage_rate_df, self.lastMonth, self.billMonth, self.lastYear, self.billYear)

        win = Toplevel()
        dashboard.Dashboard(win)
        self.window.withdraw()
        win.deiconify()






        

