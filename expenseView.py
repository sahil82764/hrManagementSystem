from tkinter import *
from tkinter import ttk
# from PIL import ImageTk, Image
from tkinter import messagebox
from tkinter import filedialog
from openpyxl import load_workbook
import datetime
import dashboard
import database
import pandas as pd
from util import util
import wageView
import calendar

class ExpenseView:
    def __init__(self, window, billPath, current_month_claimed_mandays_df, last_month_claimed_mandays_df, current_month_active_mandays_df, wage_rate_df, lastMonth, billMonth, lastYear, billYear):
        self.window = window
        window.geometry("1366x768")
        window.resizable(0, 0)
        self.window.state('zoomed')
        window.title("VENDOR EXPENSES")
        self.txt = "EXPENSES REPORT"
        self.color = ["#4f4e4d", "#f29844", "red2"]
        self.heading = Label(self.window, text=self.txt, font=('yu gothic ui', 30, "bold"), bg="white",
                             fg='black',
                             bd=5,
                             relief=FLAT)
        self.heading.grid(row=0, column=0, columnspan=3)

        # style = ttk.Style()
        # style.theme_use("clam")
        # style.configure("Treeview.Heading", font=('yu gothic ui', 10, "bold"), foreground="black",
        #                 background="#108cff")
        
        # ============ INPUT VARIABLES =================================
        self.billPath = billPath
        self.current_month_claimed_mandays_df = current_month_claimed_mandays_df
        self.last_month_claimed_mandays_df = last_month_claimed_mandays_df
        self.current_month_active_mandays_df = current_month_active_mandays_df
        self.wage_rate_df= wage_rate_df
        self.lastMonth = lastMonth
        self.billMonth =billMonth
        self.lastYear = lastYear
        self.billYear = billYear
        
        # ============ LUMPSUM REIMBURSEMENT VARIABLES =============================
        self.lumpsum1 = IntVar()        # MONTHLY MAINTAINENCE 
        self.lumpsum2 = IntVar()        # Monthly Water Charges
        self.lumpsum3 = IntVar()        # Monthly Horticulture Charges
        self.lumpsum4 = IntVar()        # Montly Air Filler Charges

        # ============ OTHER EXPENSES REIMBURSEMENT VARIABLES ======================
        self.reimburse1 = IntVar()       # Petrol for Water Mist Trolley
        self.reimburse2 = IntVar()       # Delhi Jal Board Water Bill
        self.reimburse3 = IntVar()       # MCD Trade Licence Fee
        self.reimburse4 = IntVar()       # Medical Check Up Charges
        self.reimburse5 = IntVar()       # Insurance
        self.reimburse6 = IntVar()       # Mediclaim
        self.reimburse7 = IntVar()       # Diwali Illumination
        self.reimburse8 = IntVar()       # QMS APP
        self.reimburse9 = IntVar()       # Water Tanker Charges
        self.reimburse10 = IntVar()      # Telephone Bill for EDC

        # ============ OTHER SERVICE CHARGES VARIABLES =============================
        self.charges1 = IntVar()        # Operator Service Charges
        self.charges2 = IntVar()        # SPI Claim


        # ====================== LUMPSUM REIMBURSEMENT LABELS ======================
        self.lumpsum_label = Label(self.window, text="LUMPSUM REIMBURSEMENT",font=('yu gothic ui', 18, "bold"))
        self.lumpsum_label.grid(row=1, column=0, columnspan=2, sticky="ew")

        self.lumpsum1_label = Label(self.window, text="MONTHLY MAINTAINENCE")
        self.lumpsum1_label.grid(row=2, column=1, padx=(50, 10), pady=6)

        self.lumpsum2_label = Label(self.window, text="MONTHLY WATER CHARGES")
        self.lumpsum2_label.grid(row=3, column=1, padx=(50, 10), pady=6)

        self.lumpsum3_label = Label(self.window, text="MONTHLY HORTICULTURE CHARGES")
        self.lumpsum3_label.grid(row=4, column=1, padx=(50, 10), pady=6)

        self.lumpsum4_label = Label(self.window, text="MONTHLY AIR FILLER CHARGES")
        self.lumpsum4_label.grid(row=5, column=1, padx=(50, 10), pady=6)

        # ================ OTHER EXPENSES REIMBURSEMENT LABELS ======================
        self.reimburse_label = Label(self.window, text="OTHER EXPENSES REIMBURSEMENT", font=('yu gothic ui', 18, "bold"))
        self.reimburse_label.grid(row=6, column=0, columnspan=2, sticky="ew")

        self.reimburse1_label = Label(self.window, text="PETROL FOR WATER MIST TROLLEY")
        self.reimburse1_label.grid(row=7, column=1, padx=(50, 10), pady=6)

        self.reimburse2_label = Label(self.window, text="DELHI JAL BOARD WATER BILL")
        self.reimburse2_label.grid(row=8, column=1, padx=(50, 10), pady=6)

        self.reimburse3_label = Label(self.window, text="MCD TRADE LICENCE FEE")
        self.reimburse3_label.grid(row=9, column=1, padx=(50, 10), pady=6)

        self.reimburse4_label = Label(self.window, text="MEDICAL CHECK UP CHARGES")
        self.reimburse4_label.grid(row=10, column=1, padx=(50, 10), pady=6)

        self.reimburse5_label = Label(self.window, text="INSURANCE")
        self.reimburse5_label.grid(row=11, column=1, padx=(50, 10), pady=6)

        self.reimburse6_label = Label(self.window, text="MEDICLAIM")
        self.reimburse6_label.grid(row=12, column=1, padx=(50, 10), pady=6)

        self.reimburse7_label = Label(self.window, text="DIWALI ILLUMINATION")
        self.reimburse7_label.grid(row=13, column=1, padx=(50, 10), pady=6)

        self.reimburse8_label = Label(self.window, text="QMS APP")
        self.reimburse8_label.grid(row=14, column=1, padx=(50, 10), pady=6)

        self.reimburse9_label = Label(self.window, text="WATER TANKER CHARGES")
        self.reimburse9_label.grid(row=15, column=1, padx=(50, 10), pady=6)

        self.reimburse10_label = Label(self.window, text="TELEPHONE BILL FOR EDC")
        self.reimburse10_label.grid(row=16, column=1, padx=(50, 10), pady=6)

        # ================ OTHER SERVICE CHARGES LABELS ======================
        self.charges_label = Label(self.window, text="OTHER SERVICE CHARGES", font=('yu gothic ui', 18, "bold"))
        self.charges_label.grid(row=17, column=0, columnspan=2, sticky="ew")

        self.charges1_label = Label(self.window, text="OPERATOR SERVICE CHARGES")
        self.charges1_label.grid(row=18, column=1, padx=(50, 10), pady=6)

        self.charges2_label = Label(self.window, text="SPI CLAIM")
        self.charges2_label.grid(row=19, column=1, padx=(50, 10), pady=6)


        # ====================== LUMPSUM REIMBURSEMENT ENTRIES ====================
        self.lumpsum1_entry = Entry(self.window, textvariable=self.lumpsum1, width=30)
        self.lumpsum1_entry.grid(row=2, column=2, padx=(50, 10), pady=6)

        self.lumpsum2_entry = Entry(self.window, textvariable=self.lumpsum2, width=30)
        self.lumpsum2_entry.grid(row=3, column=2, padx=(50, 10), pady=6)

        self.lumpsum3_entry = Entry(self.window, textvariable=self.lumpsum3, width=30)
        self.lumpsum3_entry.grid(row=4, column=2, padx=(50, 10), pady=6)

        self.lumpsum4_entry = Entry(self.window, textvariable=self.lumpsum4, width=30)
        self.lumpsum4_entry.grid(row=5, column=2, padx=(50, 10), pady=6)

        # ================ OTHER EXPENSES REIMBURSEMENT ENTRIES ====================
        self.reimburse1_entry = Entry(self.window, textvariable=self.reimburse1, width=30)
        self.reimburse1_entry.grid(row=7, column=2, padx=(50, 10), pady=6)

        self.reimburse2_entry = Entry(self.window, textvariable=self.reimburse2, width=30)
        self.reimburse2_entry.grid(row=8, column=2, padx=(50, 10), pady=6)

        self.reimburse3_entry = Entry(self.window, textvariable=self.reimburse3, width=30)
        self.reimburse3_entry.grid(row=9, column=2, padx=(50, 10), pady=6)

        self.reimburse4_entry = Entry(self.window, textvariable=self.reimburse4, width=30)
        self.reimburse4_entry.grid(row=10, column=2, padx=(50, 10), pady=6)

        self.reimburse5_entry = Entry(self.window, textvariable=self.reimburse5, width=30)
        self.reimburse5_entry.grid(row=11, column=2, padx=(50, 10), pady=6)

        self.reimburse6_entry = Entry(self.window, textvariable=self.reimburse6, width=30)
        self.reimburse6_entry.grid(row=12, column=2, padx=(50, 10), pady=6)

        self.reimburse7_entry = Entry(self.window, textvariable=self.reimburse7, width=30)
        self.reimburse7_entry.grid(row=13, column=2, padx=(50, 10), pady=6)

        self.reimburse8_entry = Entry(self.window, textvariable=self.reimburse8, width=30)
        self.reimburse8_entry.grid(row=14, column=2, padx=(50, 10), pady=6)

        self.reimburse9_entry = Entry(self.window, textvariable=self.reimburse9, width=30)
        self.reimburse9_entry.grid(row=15, column=2, padx=(50, 10), pady=6)

        self.reimburse10_entry = Entry(self.window, textvariable=self.reimburse10, width=30)
        self.reimburse10_entry.grid(row=16, column=2, padx=(50, 10), pady=6)

        # ================ OTHER SERVICE CHARGES ENTRIES ======================
        self.charges1_entry = Entry(self.window, textvariable=self.charges1, width=30)
        self.charges1_entry.grid(row=18, column=2, padx=(50, 10), pady=6)

        self.charges2_entry = Entry(self.window, textvariable=self.charges2, width=30)
        self.charges2_entry.grid(row=19, column=2, padx=(50, 10), pady=6)

        # SUBMIT BUTTON
        self.submit_btn = Button(self.window, text="SUBMIT ENTRIES", command=lambda: self.submit_charges())
        self.submit_btn.grid(row=20, column=0, columnspan=2, padx=(50, 10), pady=6)

        # # Place the content frame inside a canvas
        # self.canvas = Canvas(window, width=300, height=200)
        # self.canvas.grid(row=1, column=0, rowspan=20, sticky=NSEW)

        # # Create a scrollbar and attach it to the canvas
        # self.scrollbar = Scrollbar(window, orient=VERTICAL, command=self.canvas.yview)
        # self.scrollbar.grid(row=1, column=1, rowspan=20, sticky=NS)
        # self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # # Configure canvas scrolling with the mouse wheel
        # self.canvas.bind("<MouseWheel>", self.on_canvas_scroll)

        # # Create a frame to hold all the content
        # self.content_frame = Frame(self.canvas)

        # # Put the content frame inside the canvas
        # self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        # # Adjust the scroll region to update as the content frame size changes
        # self.content_frame.bind("<Configure>", self.on_frame_configure)

        # # Call the on_frame_configure method once to set the initial scroll region
        # self.on_frame_configure(None)


    def submit_charges(self):
        pass

    # def on_canvas_scroll(self, event):
    #     self.canvas.yview_scroll(-1 * int(event.delta / 120), "units")

    # def on_frame_configure(self, event):
    #     self.canvas.configure(scrollregion=self.canvas.bbox("all"))
