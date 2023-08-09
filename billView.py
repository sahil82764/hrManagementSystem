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
import shutil

class BillView:
    def __init__(self, window):
        self.window = window
        window.geometry("1366x768")
        window.resizable(0, 0)
        self.window.state('zoomed')
        window.title("VENDOR BILLING")
        self.txt = "CREATE BILL"
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

        # ============ INPUT VARIABLES =============================
        self.vendor = StringVar()
        self.po = StringVar()
        self.vendorCode = StringVar()
        self.station = []
        self.selectedStation = StringVar()
        self.contractStart = StringVar()
        self.gst = StringVar()
        self.pan = StringVar()
        self.operator = StringVar()
        self.attendencePath = StringVar()
        self.estimatePath = StringVar()
        self.billYear = IntVar()
        self.billMonth = IntVar()
        
        # ============ VENDOR NAME =================================
        self.vendor_label = Label(self.window, text="Vendor Name:")
        self.vendor_label.grid(row=1, column=0, padx=(50, 10), pady=10)

        self.vendorNames = database.get_vendors_only()
    
        self.vendor_combo = ttk.Combobox(self.window, values= self.vendorNames, textvariable=self.vendor, state='readonly')
        self.vendor_combo.grid(row=1, column=1, padx=(50, 10), pady=10)
        self.vendor_combo.bind('<<ComboboxSelected>>', self.fetchData)
    
        # PO NO
        self.po_label = Label(self.window, text="PO No.:")
        self.po_label.grid(row=2, column=0, padx=(50, 10), pady=10)

        
        self.po_no = Entry(self.window, textvariable=self.po)
        self.po_no.grid(row=2, column=1, padx=(50, 10), pady=10)

        # Vendor Code
        self.vendor_code_label = Label(self.window, text="Vendor Code")
        self.vendor_code_label.grid(row=3, column=0, padx=(50, 10), pady=10)

        self.vendor_code = Entry(self.window, textvariable=self.vendorCode)
        self.vendor_code.grid(row=3, column=1, padx=(50, 10), pady=10)

        # Station Name
        self.station_label = Label(self.window, text="Station Name:")
        self.station_label.grid(row=4, column=0, padx=(50, 10), pady=10)

        self.station_name = ttk.Combobox(self.window, values=[" "], state='readonly', textvariable=self.selectedStation)
        self.station_name.grid(row=4, column=1, padx=(50, 10), pady=10)

        # Contract Date
        self.contractStart_label = Label(self.window, text="Contract Start Date:")
        self.contractStart_label.grid(row=5, column=0, padx=(50, 10), pady=10)

        self.contractStart_date = Entry(self.window, textvariable=self.contractStart)
        self.contractStart_date.grid(row=5, column=1, padx=(50, 10), pady=10)

        # Operator Name
        self.operator_label = Label(self.window, text="Operator Name:")
        self.operator_label.grid(row=6, column=0, padx=(50, 10), pady=10)

        self.operator_name = Entry(self.window, textvariable=self.operator)
        self.operator_name.grid(row=6, column=1, padx=(50, 10), pady=10)

        # GST No
        self.gst_label = Label(self.window, text="GST No:")
        self.gst_label.grid(row=7, column=0, padx=(50, 10), pady=10)

        self.gst_no = Entry(self.window, textvariable=self.gst)
        self.gst_no.grid(row=7, column=1, padx=(50, 10), pady=10)

        # PAN No
        self.pan_label = Label(self.window, text="PAN No:")
        self.pan_label.grid(row=8, column=0, padx=(50, 10), pady=10)

        self.pan_no = Entry(self.window, textvariable=self.pan)
        self.pan_no.grid(row=8, column=1, padx=(50, 10), pady=10)

        # SALARY ESTIMATE
        self.estimate_label = Label(self.window, text="Salary Estimate Path")
        self.estimate_label.grid(row=9, column=0, padx=(50, 10), pady=10)

        self.estimate_entry = Entry(self.window, textvariable=self.estimatePath)
        self.estimate_entry.grid(row=9, column=1, padx=(50, 10), pady=10)

        self.upload_estimate_btn = Button(self.window, text="Upload Salary Estimate", command=lambda: self.upload_estimate())
        self.upload_estimate_btn.grid(row=9, column=2, columnspan=2, padx=(50, 10), pady=10)

        # Attendence File
        self.attendencePath_label = Label(self.window, text="Attendance Path")
        self.attendencePath_label.grid(row=10, column=0, padx=(50, 10), pady=10)

        self.attendencePath_entry = Entry(self.window, textvariable=self.attendencePath)
        self.attendencePath_entry.grid(row=10, column=1, padx=(50, 10), pady=10)

        self.upload_attendance_btn = Button(self.window, text="Upload Attendence", command=lambda: self.upload_attendence())
        self.upload_attendance_btn.grid(row=10, column=2, columnspan=2, padx=(50, 10), pady=10)

        # Bill Year
        self.billYear_label = Label(self.window, text="YEAR")
        self.billYear_label.grid(row=11, column=0, padx=(50, 10), pady=10)

        self.billYear_entry = ttk.Combobox(self.window, values=[" "], textvariable=self.billYear, state='readonly')
        self.billYear_entry.grid(row=11, column=1, padx=(50, 10), pady=10)

        # Bill Month
        self.billMonth_label = Label(self.window, text="MONTH")
        self.billMonth_label.grid(row=12, column=0, padx=(50, 10), pady=10)

        self.billMonth_entry = ttk.Combobox(self.window, values=[" "], textvariable=self.billMonth, state='readonly')
        self.billMonth_entry.grid(row=12, column=1, padx=(50, 10), pady=10)

        # Next Button
        self.next_btn = Button(self.window, text="NEXT", command=lambda: self.next_operation())
        self.next_btn.grid(row=13, column=0, columnspan=2, padx=(50, 10), pady=10)

        # Back Button
        self.back_btn = Button(self.window, text="BACK", command=lambda: self.back_operation())
        self.back_btn.grid(row=13, column=1, columnspan=2, padx=(50, 10), pady=10)

    def fetchData(self, event):

        if self.vendor.get() is not None:
                
                self.vendor_data_dictionary = database.get_vendor_data(self.vendor.get())

                self.contractEnd = datetime.datetime.strptime(self.vendor_data_dictionary['Contract Date'], '%Y-%m-%d').date()  + datetime.timedelta(days=365*5)

                yearList = []
                for year in range(datetime.datetime.now().year, self.contractEnd.year + 1):
                    yearList.append(year)

                
                self.po_no.config(state='normal')
                self.po_no.delete(0, END)
                self.po_no.insert(0, self.vendor_data_dictionary['PO No'])
                self.po_no.config(state='readonly')

                self.vendor_code.config(state='normal')
                self.vendor_code.delete(0, END)
                self.vendor_code.insert(0, self.vendor_data_dictionary['Vendor Code'])
                self.vendor_code.config(state='readonly')
                

                self.station_name.delete(0, END)
                self.station_name.config(values=self.vendor_data_dictionary['Station Name'])
                self.station_name.current(0)


                self.contractStart_date.config(state='normal')
                self.contractStart_date.delete(0, END)
                self.contractStart_date.insert(0, self.vendor_data_dictionary['Contract Date'])
                self.contractStart_date.config(state='readonly')

                self.gst_no.config(state='normal')
                self.gst_no.delete(0, END)
                self.gst_no.insert(0, self.vendor_data_dictionary['GST No'])
                self.gst_no.config(state='readonly')

                self.pan_no.config(state='normal')
                self.pan_no.delete(0, END)
                self.pan_no.insert(0, self.vendor_data_dictionary['PAN No'])
                self.pan_no.config(state='readonly')

                self.operator_name.config(state='normal')
                self.operator_name.delete(0, END)
                self.operator_name.insert(0, self.vendor_data_dictionary['Operator Name'])
                self.operator_name.config(state='readonly')

                self.attendencePath_entry.config(state='normal')
                self.attendencePath_entry.delete(0, END)
                self.attendencePath_entry.config(state='readonly')

                self.estimate_entry.config(state='normal')
                self.estimate_entry.delete(0, END)
                self.estimate_entry.config(state='readonly')

                self.billYear_entry.delete(0, END)
                self.billYear_entry.config(values= yearList)
                self.billYear_entry.current(0)

                self.billMonth_entry.delete(0, END)
                self.billMonth_entry.config(values=[1,2,3,4,5,6,7,8,9,10,11,12])
                self.billMonth_entry.current(0)


    def upload_estimate(self):
        try:
            self.filePath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
            self.estimate_entry.config(state='normal')
            self.estimate_entry.delete(0, END)  # Clear previous path, if any
            self.estimate_entry.insert(END, self.filePath)  # Display the selected path
            self.estimate_entry.config(state='readonly')
        
        except Exception as e:
            print(e)
    
    def upload_attendence(self):
        try:
            self.filePath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
            self.attendencePath_entry.config(state='normal')
            self.attendencePath_entry.delete(0, END)  # Clear previous path, if any
            self.attendencePath_entry.insert(END, self.filePath)  # Display the selected path
            self.attendencePath_entry.config(state='readonly')
        
        except Exception as e:
            print(e)


    def next_operation(self):
        if (
            self.vendor.get() and 
            self.po.get() and 
            self.selectedStation.get() and 
            self.contractStart.get() and 
            self.operator.get() and 
            self.gst.get() and 
            self.pan.get() and 
            self.attendencePath.get() and 
            self.estimatePath.get() and 
            self.billYear.get() and 
            self.billMonth.get()):
            # All fields are filled, perform the add mandays operation
            self.copy_file()
            self.get_active_mandays()
            self.create_bill_file()
            self.perform_wage_operation()
            
        else:
            # Display an error message if any field is empty
            messagebox.showerror("Error", "Please fill in all the fields.")

    def back_operation(self):
        win = Toplevel()
        dashboard.Dashboard(win)
        self.window.withdraw()
        win.deiconify()

    def get_active_mandays(self):
        
        all_designations = ['DSM', 'TECH', 'MGR']
        activeMandaysDF = pd.read_excel(str(self.attendencePath.get()), header=2)
        activeMandaysDF = activeMandaysDF.iloc[:, 1:]
        activeMandaysDF = activeMandaysDF.dropna(axis=0, subset=['Name of Employee', "Father's Name", 'Designation']).reset_index(drop=True)

        col = list(activeMandaysDF.columns)[2:]
        activeMandaysDF_attendance = pd.DataFrame(columns=['Designation'] + list(activeMandaysDF.columns)[3:])
        attendance = activeMandaysDF_attendance.copy()
        activeMandaysDF_attendance['Designation'] = all_designations

        for i in activeMandaysDF.Designation.unique():
            tempList = [i] + [activeMandaysDF.loc[activeMandaysDF['Designation'] == i, j].sum() for j in col[1:]]
            attendance = pd.concat([attendance, pd.DataFrame([tempList], columns=col)], ignore_index=True)

        activeMandaysDF_attendance.set_index('Designation', inplace=True)
        attendance.set_index('Designation', inplace=True)

        mergedDf = pd.concat([activeMandaysDF_attendance, attendance], axis=0)

        outputDf = mergedDf.groupby(level=0).sum().reset_index()
        outputDf.columns = ['FT' if col == 'F&S' else col for col in outputDf.columns]
        outputDf.columns = ['OFF' if col == 'Off' else col for col in outputDf.columns]

        activeMandaysPath = util.save_mandays('Active', str(self.billYear.get()), str(self.billMonth.get()), str(self.vendor.get()), str(self.selectedStation.get()))
        outputDf.to_excel(activeMandaysPath, index=False)

    def create_bill_file(self):
        try:
            billPath = util.get_custom_template('Bill')
            customBill = load_workbook(billPath)
            activeSheet = customBill.active

            activeSheet['B2'] = self.po.get()
            activeSheet['B3'] = self.vendor_code.get()
            activeSheet['B4'] = self.vendor.get()
            activeSheet['B5'] = self.selectedStation.get()
            activeSheet['B6'] = self.contractStart.get()
            activeSheet['G2'] = f"{calendar.month_abbr[self.billMonth.get()]}-{self.billYear.get()}"
            activeSheet['G3'] = self.gst.get()
            activeSheet['G4'] = self.pan.get()
            activeSheet['G5'] = self.operator.get()
            activeSheet['G6'] = f"{calendar.month_abbr[self.billMonth.get()]}-{self.billYear.get()}"

            #saving the modified workbook
            self.billSavePath = util.save_bill(str(self.billYear.get()), str(self.billMonth.get()), str(self.vendor.get()), str(self.selectedStation.get()))
            customBill.save(self.billSavePath)
        
        except Exception as e:
            print(e)

    def copy_file(self):
        try:
            sourcePath_estimate = self.estimatePath
            destinationPath_estimate = util.get_estimate_path(self.billYear, self.billMonth, self.vendor.get(), self.selectedStation.get())
            shutil.copy(sourcePath_estimate, destinationPath_estimate)
            print(f"File copied successfully from '{sourcePath_estimate}' to '{destinationPath_estimate}'.")


            sourcePath_attendance = self.attendencePath
            destinationPath_attendance = util.get_attendance_path(self.billYear, self.billMonth, self.vendor.get(), self.selectedStation.get())
            shutil.copy(sourcePath_attendance, destinationPath_attendance)
            print(f"File copied successfully from '{sourcePath_attendance}' to '{destinationPath_attendance}'.")

        except Exception as e:
            print(f"An error occurred: {e}")
    

    def perform_wage_operation(self):
        if self.billMonth.get() == 1:
            lastMonth = 12
            lastYear = self.billYear.get() - 1

        else:
            lastMonth = self.billMonth.get() - 1
            lastYear = self.billYear.get()

        current_month_claimed_mandays = util.get_estimate_path(self.billYear.get(), self.billMonth.get(), self.vendor.get(), self.selectedStation.get())
        last_month_claimed_mandays = util.get_estimate_path(lastYear, lastMonth, self.vendor.get(), self.selectedStation.get())
        current_month_active_mandays = util.get_mandays(self.billYear.get(), self.billMonth.get(), self.vendor.get(), self.selectedStation.get())

        win = Toplevel()
        wageView.WageView(win, self.billSavePath, current_month_claimed_mandays, last_month_claimed_mandays, current_month_active_mandays, lastMonth, self.billMonth.get(), lastYear, self.billYear.get())
        self.window.withdraw()
        win.deiconify()

        
    
