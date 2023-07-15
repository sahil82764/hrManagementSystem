from tkinter import *
from tkinter import ttk
import datetime
from util import util
from openpyxl import load_workbook
from tkinter import messagebox
import dashboard
import mandaysView


class AddMandays:
    def __init__(self, window, contractEnd, vendorName, stationName):
        self.window = window
        self.window.geometry("1366x768")
        self.window.title("Add Mandays")
        self.window.resizable(False, False)
        self.window.config(background='#98a65d')
        self.txt = "ADD MANDAYS"
        self.heading = Label(self.window, text=self.txt, font=('yu gothic ui', 30, "bold"), bg="white",
                             fg='black',
                             bd=5,
                             relief=FLAT)
        self.heading.grid(row=0, column=0, columnspan=10, pady=(50, 20))
        
        # self.window.grid_columnconfigure(0, weight=1)
        # self.window.grid_rowconfigure(0, weight=0)

        # self.contractStart = contractStart
        # self.contractEnd = contractEnd

        # =================== INPUT VARIABLES =================
        self.mgr_wd = StringVar()
        self.mgr_off = StringVar()
        self.mgr_cl = StringVar()
        self.mgr_ft = StringVar()
        self.mgr_fh = StringVar()
        self.mgr_nh = StringVar()
        self.dsm_wd = StringVar()
        self.dsm_off = StringVar()
        self.dsm_cl = StringVar()
        self.dsm_ft = StringVar()
        self.dsm_fh = StringVar()
        self.dsm_nh = StringVar()
        self.tech_wd = StringVar()
        self.tech_off = StringVar()
        self.tech_cl = StringVar()
        self.tech_ft = StringVar()
        self.tech_fh = StringVar()
        self.tech_nh = StringVar()
        self.year = StringVar()
        self.month = StringVar()
        self.vName = vendorName
        self.sName = stationName

        # ================ LABELS =============================

        # YEAR
        self.year_label = Label(self.window, text='YEAR')
        self.year_label.grid(row=1, column=0, padx=(50, 10), pady=10)

        self.year_entry = ttk.Combobox(self.window, textvariable=self.year, state='readonly')
        
        yearList = []
        for year in range(datetime.datetime.now().year, contractEnd.year + 1):
            yearList.append(year)
        
        self.year_entry['values'] = yearList
        self.year_entry.grid(row=1, column=1, padx=(50, 10), pady=10)

        # MONTH
        self.month_label = Label(self.window, text='MONTH')
        self.month_label.grid(row=2, column=0, padx=(50, 10), pady=10)

        # monthList = []
        # for month in range(datetime.datetime.now().month, 13):
        #     monthList.append(month)

        self.month_entry = ttk.Combobox(self.window, values=[1,2,3,4,5,6,7,8,9,10,11,12], textvariable=self.month, state='readonly')
        self.month_entry.grid(row=2, column=1, padx=(50, 10), pady=10)

        # MANDAYS LABEL

        mt_label = Label(self.window, text='', bg="#98a65d", fg="#98a65d")
        mt_label.grid(row=3, column=0, padx=(100, 100), pady=30)

        wd_label = Label(self.window, text='WD')
        wd_label.grid(row=4, column=1, padx=(50, 10), pady=10)

        off_label = Label(self.window, text='OFF')
        off_label.grid(row=4, column=2, padx=(50, 10), pady=10)

        cl_label = Label(self.window, text='CL')
        cl_label.grid(row=4, column=3, padx=(50, 10), pady=10)

        ft_label = Label(self.window, text='FT')
        ft_label.grid(row=4, column=4, padx=(50, 10), pady=10)

        fh_label = Label(self.window, text='FH')
        fh_label.grid(row=4, column=5, padx=(50, 10), pady=10)

        nh_label = Label(self.window, text='NH')
        nh_label.grid(row=4, column=6, padx=(50, 10), pady=10)

        mgr_label = Label(self.window, text='MANAGER')
        mgr_label.grid(row=5, column=0, padx=(50, 10), pady=10)

        dsm_label = Label(self.window, text='TECH')
        dsm_label.grid(row=6, column=0, padx=(50, 10), pady=10)

        tech_label = Label(self.window, text='DSM')
        tech_label.grid(row=7, column=0, padx=(50, 10), pady=10)

        # MANDAYS ENTRY

        # MANAGER
        mgr_wd_entry = Entry(self.window, textvariable=self.mgr_wd)
        mgr_wd_entry.grid(row=5, column=1, padx=(50, 10), pady=10)

        mgr_off_entry = Entry(self.window, textvariable=self.mgr_off)
        mgr_off_entry.grid(row=5, column=2, padx=(50, 10), pady=10)

        mgr_cl_entry = Entry(self.window, textvariable=self.mgr_cl)
        mgr_cl_entry.grid(row=5, column=3, padx=(50, 10), pady=10)

        mgr_ft_entry = Entry(self.window, textvariable=self.mgr_ft)
        mgr_ft_entry.grid(row=5, column=4, padx=(50, 10), pady=10)

        mgr_fh_entry = Entry(self.window, textvariable=self.mgr_fh)
        mgr_fh_entry.grid(row=5, column=5, padx=(50, 10), pady=10)

        mgr_nh_entry = Entry(self.window, textvariable=self.mgr_nh)
        mgr_nh_entry.grid(row=5, column=6, padx=(50, 10), pady=10)

        # TECH
        tech_wd_entry = Entry(self.window, textvariable=self.tech_wd)
        tech_wd_entry.grid(row=6, column=1, padx=(50, 10), pady=10)

        tech_off_entry = Entry(self.window, textvariable=self.tech_off)
        tech_off_entry.grid(row=6, column=2, padx=(50, 10), pady=10)

        tech_cl_entry = Entry(self.window, textvariable=self.tech_cl)
        tech_cl_entry.grid(row=6, column=3, padx=(50, 10), pady=10)

        tech_ft_entry = Entry(self.window, textvariable=self.tech_ft)
        tech_ft_entry.grid(row=6, column=4, padx=(50, 10), pady=10)

        tech_fh_entry = Entry(self.window, textvariable=self.tech_fh)
        tech_fh_entry.grid(row=6, column=5, padx=(50, 10), pady=10)

        tech_nh_entry = Entry(self.window, textvariable=self.tech_nh)
        tech_nh_entry.grid(row=6, column=6, padx=(50, 10), pady=10)

        # DSM
        dsm_wd_entry = Entry(self.window, textvariable=self.dsm_wd)
        dsm_wd_entry.grid(row=7, column=1, padx=(50, 10), pady=10)

        dsm_off_entry = Entry(self.window, textvariable=self.dsm_off)
        dsm_off_entry.grid(row=7, column=2, padx=(50, 10), pady=10)

        dsm_cl_entry = Entry(self.window, textvariable=self.dsm_cl)
        dsm_cl_entry.grid(row=7, column=3, padx=(50, 10), pady=10)

        dsm_ft_entry = Entry(self.window, textvariable=self.dsm_ft)
        dsm_ft_entry.grid(row=7, column=4, padx=(50, 10), pady=10)

        dsm_fh_entry = Entry(self.window, textvariable=self.dsm_fh)
        dsm_fh_entry.grid(row=7, column=5, padx=(50, 10), pady=10)

        dsm_nh_entry = Entry(self.window, textvariable=self.dsm_nh)
        dsm_nh_entry.grid(row=7, column=6, padx=(50, 10), pady=10)

        mt_label = Label(self.window, text='', bg="#98a65d", fg="#98a65d")
        mt_label.grid(row=8, column=0, padx=(100, 100), pady=10)

        self.saveMandays_btn = Button(self.window, text="SAVE", width=15, height=1, font=('Arial', 15, "bold"), command=lambda: self.save_mandays())
        self.saveMandays_btn.grid(row=9, column=0, columnspan=3,  padx=(50, 10), pady=10)

        self.back_btn = Button(self.window, text="BACK", width=15, height=1, font=('Arial', 15, "bold"), command=lambda: self.back_operation())
        self.back_btn.grid(row=9, column=1, columnspan=3,  padx=(50, 10), pady=10)

    def save_mandays(self):
        if (
            self.mgr_wd.get() and self.mgr_off.get() and self.mgr_cl.get() and self.mgr_ft.get() and self.mgr_fh.get() and self.mgr_nh.get() and self.dsm_wd.get() and self.dsm_off.get() and self.dsm_cl.get() and self.dsm_ft.get() and self.dsm_fh.get() and self.dsm_nh.get() and self.tech_wd.get() and self.tech_off.get() and self.tech_cl.get() and self.tech_ft.get() and self.tech_fh.get() and self.tech_nh.get() and self.year.get() and self.month.get()):
            # All fields are filled, perform the save operation
            self.perform_save_operation()
        else:
            # Display an error message if any field is empty
            messagebox.showerror("Error", "Please fill in all the fields.")

    def back_operation(self):
        win = Toplevel()
        mandaysView.MandaysView(win)
        self.window.withdraw()
        win.deiconify()
        

    def perform_save_operation(self):
        try:
            fPath = util.get_custom_template('Mandays')
            customMandays = load_workbook(fPath)
            activeSheet = customMandays.active

            # manager entries
            activeSheet['B2'] = self.mgr_wd.get()
            activeSheet['C2'] = self.mgr_off.get()
            activeSheet['D2'] = self.mgr_cl.get()
            activeSheet['E2'] = self.mgr_ft.get()
            activeSheet['F2'] = self.mgr_fh.get()
            activeSheet['G2'] = self.mgr_nh.get()

            # technician entries
            activeSheet['B3'] = self.tech_wd.get()
            activeSheet['C3'] = self.tech_off.get()
            activeSheet['D3'] = self.tech_cl.get()
            activeSheet['E3'] = self.tech_ft.get()
            activeSheet['F3'] = self.tech_fh.get()
            activeSheet['G3'] = self.tech_nh.get()

            # dsm entries
            activeSheet['B4'] = self.dsm_wd.get()
            activeSheet['C4'] = self.dsm_off.get()
            activeSheet['D4'] = self.dsm_cl.get()
            activeSheet['E4'] = self.dsm_ft.get()
            activeSheet['F4'] = self.dsm_fh.get()
            activeSheet['G4'] = self.dsm_nh.get()

            #saving the modified workbook
            savePath = util.save_mandays('Claimed', str(self.year.get()), str(self.month.get()), self.vName.get(), str(self.sName.get()))
            customMandays.save(savePath)

            win = Toplevel()
            dashboard.Dashboard(win)
            self.window.withdraw()
            win.deiconify()

        except Exception as e:
            print(e)
            
            
                         
        