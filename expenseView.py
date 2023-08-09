# from tkinter import *
# from tkinter import ttk
# # from PIL import ImageTk, Image
# from tkinter import messagebox
# # from tkinter import filedialog
# from openpyxl import load_workbook
# import dashboard
# import wageView
# import pandas as pd
# import generateBill

# class ExpenseView:
#     def __init__(self, window, billPath, current_month_claimed_mandays, last_month_claimed_mandays, current_month_active_mandays_df, wage_rate_df, lastMonth, billMonth, lastYear, billYear):
#         self.window = window
#         window.geometry("1366x768")
#         window.resizable(0, 0)
#         self.window.state('zoomed')
#         window.title("VENDOR EXPENSES")
#         self.txt = "EXPENSES REPORT"
#         self.color = ["#4f4e4d", "#f29844", "red2"]
#         self.heading = Label(self.window, text=self.txt, font=('yu gothic ui', 30, "bold"), bg="white",
#                              fg='black',
#                              bd=5,
#                              relief=FLAT)
#         self.heading.grid(row=0, column=0, columnspan=3)

#         # style = ttk.Style()
#         # style.theme_use("clam")
#         # style.configure("Treeview.Heading", font=('yu gothic ui', 10, "bold"), foreground="black",
#         #                 background="#108cff")
        
#         # ============ INPUT VARIABLES =================================
#         self.billPath = billPath
#         self.current_month_claimed_mandays = current_month_claimed_mandays
#         self.last_month_claimed_mandays = last_month_claimed_mandays
#         self.current_month_active_mandays_df = current_month_active_mandays_df
#         self.wage_rate_df= wage_rate_df
#         self.lastMonth = lastMonth
#         self.billMonth =billMonth
#         self.lastYear = lastYear
#         self.billYear = billYear
        
#         # ============ LUMPSUM REIMBURSEMENT VARIABLES =============================
#         self.lumpsum1 = StringVar()        # MONTHLY MAINTAINENCE 
#         self.lumpsum2 = StringVar()        # Monthly Water Charges
#         self.lumpsum3 = StringVar()        # Monthly Horticulture Charges
#         self.lumpsum4 = StringVar()        # Montly Air Filler Charges

#         # ============ OTHER EXPENSES REIMBURSEMENT VARIABLES ======================
#         self.reimburse1 = StringVar()       # Petrol for Water Mist Trolley
#         self.reimburse2 = StringVar()       # Delhi Jal Board Water Bill
#         self.reimburse3 = StringVar()       # MCD Trade Licence Fee
#         self.reimburse4 = StringVar()       # Medical Check Up Charges
#         self.reimburse5 = StringVar()       # Insurance
#         self.reimburse6 = StringVar()       # Mediclaim
#         self.reimburse7 = StringVar()       # Diwali Illumination
#         self.reimburse8 = StringVar()       # QMS APP
#         self.reimburse9 = StringVar()       # Water Tanker Charges
#         self.reimburse10 = StringVar()      # Telephone Bill for EDC

#         # ============ OTHER SERVICE CHARGES VARIABLES =============================
#         self.charges1 = StringVar()        # Operator Service Charges
#         self.charges2 = StringVar()        # SPI Claim

#         # ============ MANPOWER DEPLOYED VARIABLES =================================
#         self.dsm_deployed = StringVar()
#         self.tech_deployed = StringVar()
#         self.mgr_deployed = StringVar()

#         self.dsm_roll = StringVar()
#         self.tech_roll = StringVar()
#         self.mgr_roll = StringVar()


#         # ====================== LUMPSUM REIMBURSEMENT LABELS ======================
#         self.lumpsum_label = Label(self.window, text="LUMPSUM REIMBURSEMENT",font=('yu gothic ui', 18, "bold"))
#         self.lumpsum_label.grid(row=1, column=0, columnspan=2, sticky="ew")

#         self.lumpsum1_label = Label(self.window, text="MONTHLY MAINTAINENCE")
#         self.lumpsum1_label.grid(row=2, column=1, padx=(50, 10), pady=6)

#         self.lumpsum2_label = Label(self.window, text="MONTHLY WATER CHARGES")
#         self.lumpsum2_label.grid(row=3, column=1, padx=(50, 10), pady=6)

#         self.lumpsum3_label = Label(self.window, text="MONTHLY HORTICULTURE CHARGES")
#         self.lumpsum3_label.grid(row=4, column=1, padx=(50, 10), pady=6)

#         self.lumpsum4_label = Label(self.window, text="MONTHLY AIR FILLER CHARGES")
#         self.lumpsum4_label.grid(row=5, column=1, padx=(50, 10), pady=6)

#         # ================ OTHER EXPENSES REIMBURSEMENT LABELS ======================
#         self.reimburse_label = Label(self.window, text="OTHER EXPENSES REIMBURSEMENT", font=('yu gothic ui', 18, "bold"))
#         self.reimburse_label.grid(row=6, column=0, columnspan=2, sticky="ew")

#         self.reimburse1_label = Label(self.window, text="PETROL FOR WATER MIST TROLLEY")
#         self.reimburse1_label.grid(row=7, column=1, padx=(50, 10), pady=6)

#         self.reimburse2_label = Label(self.window, text="DELHI JAL BOARD WATER BILL")
#         self.reimburse2_label.grid(row=8, column=1, padx=(50, 10), pady=6)

#         self.reimburse3_label = Label(self.window, text="MCD TRADE LICENCE FEE")
#         self.reimburse3_label.grid(row=9, column=1, padx=(50, 10), pady=6)

#         self.reimburse4_label = Label(self.window, text="MEDICAL CHECK UP CHARGES")
#         self.reimburse4_label.grid(row=10, column=1, padx=(50, 10), pady=6)

#         self.reimburse5_label = Label(self.window, text="INSURANCE")
#         self.reimburse5_label.grid(row=11, column=1, padx=(50, 10), pady=6)

#         self.reimburse6_label = Label(self.window, text="MEDICLAIM")
#         self.reimburse6_label.grid(row=12, column=1, padx=(50, 10), pady=6)

#         self.reimburse7_label = Label(self.window, text="DIWALI ILLUMINATION")
#         self.reimburse7_label.grid(row=13, column=1, padx=(50, 10), pady=6)

#         self.reimburse8_label = Label(self.window, text="QMS APP")
#         self.reimburse8_label.grid(row=14, column=1, padx=(50, 10), pady=6)

#         self.reimburse9_label = Label(self.window, text="WATER TANKER CHARGES")
#         self.reimburse9_label.grid(row=15, column=1, padx=(50, 10), pady=6)

#         self.reimburse10_label = Label(self.window, text="TELEPHONE BILL FOR EDC")
#         self.reimburse10_label.grid(row=16, column=1, padx=(50, 10), pady=6)

#         # ================ OTHER SERVICE CHARGES LABELS ======================
#         self.charges_label = Label(self.window, text="OTHER SERVICE CHARGES", font=('yu gothic ui', 18, "bold"))
#         self.charges_label.grid(row=17, column=0, columnspan=2, sticky="ew")

#         self.charges1_label = Label(self.window, text="OPERATOR SERVICE CHARGES")
#         self.charges1_label.grid(row=18, column=1, padx=(50, 10), pady=6)

#         self.charges2_label = Label(self.window, text="SPI CLAIM")
#         self.charges2_label.grid(row=19, column=1, padx=(50, 10), pady=6)


#         # ====================== LUMPSUM REIMBURSEMENT ENTRIES ====================
#         self.lumpsum1_entry = Entry(self.window, textvariable=self.lumpsum1, width=30)
#         self.lumpsum1_entry.grid(row=2, column=2, padx=(50, 10), pady=6)

#         self.lumpsum2_entry = Entry(self.window, textvariable=self.lumpsum2, width=30)
#         self.lumpsum2_entry.grid(row=3, column=2, padx=(50, 10), pady=6)

#         self.lumpsum3_entry = Entry(self.window, textvariable=self.lumpsum3, width=30)
#         self.lumpsum3_entry.grid(row=4, column=2, padx=(50, 10), pady=6)

#         self.lumpsum4_entry = Entry(self.window, textvariable=self.lumpsum4, width=30)
#         self.lumpsum4_entry.grid(row=5, column=2, padx=(50, 10), pady=6)

#         # ================ OTHER EXPENSES REIMBURSEMENT ENTRIES ====================
#         self.reimburse1_entry = Entry(self.window, textvariable=self.reimburse1, width=30)
#         self.reimburse1_entry.grid(row=7, column=2, padx=(50, 10), pady=6)

#         self.reimburse2_entry = Entry(self.window, textvariable=self.reimburse2, width=30)
#         self.reimburse2_entry.grid(row=8, column=2, padx=(50, 10), pady=6)

#         self.reimburse3_entry = Entry(self.window, textvariable=self.reimburse3, width=30)
#         self.reimburse3_entry.grid(row=9, column=2, padx=(50, 10), pady=6)

#         self.reimburse4_entry = Entry(self.window, textvariable=self.reimburse4, width=30)
#         self.reimburse4_entry.grid(row=10, column=2, padx=(50, 10), pady=6)

#         self.reimburse5_entry = Entry(self.window, textvariable=self.reimburse5, width=30)
#         self.reimburse5_entry.grid(row=11, column=2, padx=(50, 10), pady=6)

#         self.reimburse6_entry = Entry(self.window, textvariable=self.reimburse6, width=30)
#         self.reimburse6_entry.grid(row=12, column=2, padx=(50, 10), pady=6)

#         self.reimburse7_entry = Entry(self.window, textvariable=self.reimburse7, width=30)
#         self.reimburse7_entry.grid(row=13, column=2, padx=(50, 10), pady=6)

#         self.reimburse8_entry = Entry(self.window, textvariable=self.reimburse8, width=30)
#         self.reimburse8_entry.grid(row=14, column=2, padx=(50, 10), pady=6)

#         self.reimburse9_entry = Entry(self.window, textvariable=self.reimburse9, width=30)
#         self.reimburse9_entry.grid(row=15, column=2, padx=(50, 10), pady=6)

#         self.reimburse10_entry = Entry(self.window, textvariable=self.reimburse10, width=30)
#         self.reimburse10_entry.grid(row=16, column=2, padx=(50, 10), pady=6)

#         # ================ OTHER SERVICE CHARGES ENTRIES ======================
#         self.charges1_entry = Entry(self.window, textvariable=self.charges1, width=30)
#         self.charges1_entry.grid(row=18, column=2, padx=(50, 10), pady=6)

#         self.charges2_entry = Entry(self.window, textvariable=self.charges2, width=30)
#         self.charges2_entry.grid(row=19, column=2, padx=(50, 10), pady=6)

#         # SUBMIT BUTTON
#         self.submit_btn = Button(self.window, text=" SUBMIT ENTRIES AND GENERATE BILL ", font=('yu gothic ui', 12, "bold"), command=lambda: self.submit_charges())
#         self.submit_btn.grid(row=20, column=1, columnspan=4, padx=(50, 10), pady=10, sticky='ew')

#         # BACK BUTTON
#         self.back_btn = Button(self.window, text=" BACK ", font=('yu gothic ui', 12, "bold"), command=lambda: self.back_operation())
#         self.back_btn.grid(row=20, column=5, columnspan=4, padx=(50, 10), pady=10, sticky='ew')

#         for i in range(1, 20):
#             self.label = Label(self.window, text="    ||")
#             self.label.grid(row=i, column=3)

#         # ============== MANPOWER DETAILS LABELS ===============================
#         self.manpowerDetails_label = Label(self.window, text="MANPOWER DETAILS", font=('yu gothic ui', 14, "bold"))
#         self.manpowerDetails_label.grid(row=7, column=4, padx=(50, 10), pady=10)

#         self.manpowerDeployed_label = Label(self.window, text="MANPOWER DEPLOYED", font=('yu gothic ui', 10, "bold"))
#         self.manpowerDeployed_label.grid(row=8, column=4, padx=(50, 10), pady=10)

#         self.manpowerOnRoll_label = Label(self.window, text="MANPOWER ON MUSTER ROLL", font=('yu gothic ui', 10, "bold"))
#         self.manpowerOnRoll_label.grid(row=9, column=4, padx=(50, 10), pady=10)

#         self.manpowerDSM_label = Label(self.window, text="DSM / TM", font=('yu gothic ui', 10, "bold"))
#         self.manpowerDSM_label.grid(row=7, column=5, padx=(50, 10), pady=10)

#         self.manpowerTECH_label = Label(self.window, text="TECH", font=('yu gothic ui', 10, "bold"))
#         self.manpowerTECH_label.grid(row=7, column=6, padx=(50, 10), pady=10)

#         self.manpowerMGR_label = Label(self.window, text="MGR", font=('yu gothic ui', 10, "bold"))
#         self.manpowerMGR_label.grid(row=7, column=7, padx=(50, 10), pady=10)

#         # ============== MANPOWER DETAILS ENTRIES ===============================
#         self.dsmDeployed_entry = Entry(self.window, textvariable=self.dsm_deployed)
#         self.dsmDeployed_entry.grid(row=8, column=5, padx=(50, 10), pady=10)

#         self.dsmOnRoll_entry = Entry(self.window, textvariable=self.dsm_roll)
#         self.dsmOnRoll_entry.grid(row=9, column=5, padx=(50, 10), pady=10)

#         self.techDeployed_entry = Entry(self.window, textvariable=self.tech_deployed)
#         self.techDeployed_entry.grid(row=8, column=6, padx=(50, 10), pady=10)

#         self.techOnRoll_entry = Entry(self.window, textvariable=self.tech_roll)
#         self.techOnRoll_entry.grid(row=9, column=6, padx=(50, 10), pady=10)

#         self.mgrDeployed_entry = Entry(self.window, textvariable=self.mgr_deployed)
#         self.mgrDeployed_entry.grid(row=8, column=7, padx=(50, 10), pady=10)

#         self.mgrOnRoll_entry = Entry(self.window, textvariable=self.mgr_roll)
#         self.mgrOnRoll_entry.grid(row=9, column=7, padx=(50, 10), pady=10)

        

#     def submit_charges(self):

#         if (
#             self.lumpsum1.get() and
#             self.lumpsum2.get() and
#             self.lumpsum3.get() and
#             self.lumpsum4.get() and
#             self.reimburse1.get() and
#             self.reimburse2.get() and
#             self.reimburse3.get() and
#             self.reimburse4.get() and
#             self.reimburse5.get() and
#             self.reimburse6.get() and
#             self.reimburse7.get() and
#             self.reimburse8.get() and
#             self.reimburse9.get() and
#             self.reimburse10.get() and
#             self.charges1.get() and
#             self.charges2.get() and
#             self.dsm_deployed.get() and
#             self.tech_deployed.get() and
#             self.mgr_deployed.get() and
#             self.dsm_roll.get() and
#             self.tech_roll.get() and
#             self.mgr_roll.get()
#         ):
#             generateBill.createBill(self.billPath, self.current_month_claimed_mandays_df, self.last_month_claimed_mandays_df, self.current_month_active_mandays_df, self.wage_rate_df, self.lastMonth, self.billMonth, self.lastYear, self.billYear)
    
#             billWOrkbook = load_workbook(self.billPath)
#             active_bill_sheet = billWOrkbook.active

#             # =============== LUMPSUM REIMBURSEMENT CELL: H50-HH53 ===============
#             active_bill_sheet['H50'] = int(self.lumpsum1.get())
#             active_bill_sheet['H51'] = int(self.lumpsum2.get())
#             active_bill_sheet['H52'] = int(self.lumpsum3.get())
#             active_bill_sheet['H53'] = int(self.lumpsum4.get())

#             # =============== OTHER EXPENSES REIMBURSEMENT CELL: H55-H64 =========
#             active_bill_sheet['H55'] = int(self.reimburse1.get())
#             active_bill_sheet['H56'] = int(self.reimburse2.get())
#             active_bill_sheet['H57'] = int(self.reimburse3.get())
#             active_bill_sheet['H58'] = int(self.reimburse4.get())
#             active_bill_sheet['H59'] = int(self.reimburse5.get())
#             active_bill_sheet['H60'] = int(self.reimburse6.get())
#             active_bill_sheet['H61'] = int(self.reimburse7.get() )
#             active_bill_sheet['H62'] = int(self.reimburse8.get())
#             active_bill_sheet['H63'] = int(self.reimburse9.get())
#             active_bill_sheet['H64'] = int(self.reimburse10.get())

#             # =============== OPERATOR SERVICE CHARGES CELL: H66-H67 ==============
#             active_bill_sheet['H66'] = int(self.charges1.get())
#             active_bill_sheet['H67'] = int(self.charges2.get())

#             # =============== MANPOWER DEPLOYED CELLS =============================
#             active_bill_sheet['F8'] = int(self.dsm_deployed.get())
#             active_bill_sheet['F9'] = int(self.dsm_roll.get())
#             active_bill_sheet['G8'] = int(self.tech_deployed.get())
#             active_bill_sheet['G9'] = int(self.tech_roll.get())
#             active_bill_sheet['H8'] = int(self.mgr_deployed.get())
#             active_bill_sheet['H9'] = int(self.mgr_roll.get())

#             # =============== Saving Workbook  ===============
#             billWOrkbook.save(self.billPath)

#             win = Toplevel()
#             dashboard.Dashboard(win)
#             self.window.withdraw()
#             win.deiconify()

#         else:
#             # Display an error message if any field is empty
#             messagebox.showerror("Error", "Please fill in all the fields.")

#     def back_operation(self):
#         win = Toplevel()
#         wageView.WageView(win)
#         self.window.withdraw()
#         win.deiconify()


