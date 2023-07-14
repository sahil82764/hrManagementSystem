from tkinter import *
from tkinter import ttk
# from PIL import ImageTk, Image
from tkinter import messagebox
from tkinter import filedialog
import datetime
import dashboard
import database

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
        self.vendor_code_label = Label(self.window, text="PO No.:")
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

        # Upload Button
        self.upload_btn = Button(self.window, text="Upload Attendence", command=lambda: self.upload_attendence())
        self.upload_btn.grid(row=9, column=0, columnspan=2, padx=(50, 10), pady=10)

        # Attendence Excel File
        self.attendencePath_label = Label(window, text="File Path:")
        self.attendencePath_label.grid(row=10, column=0, padx=(50, 10), pady=10)

        self.attendencePath_entry = Entry(self.window, textvariable=self.attendencePath)
        self.attendencePath_entry.grid(row=10, column=1, padx=(50, 10), pady=10)

        # Next Button
        self.next_btn = Button(self.window, text="NEXT", command=lambda: self.next_operation())
        self.next_btn.grid(row=11, column=0, columnspan=3, padx=(50, 10), pady=10)

    def fetchData(self, event):

        if self.vendor.get() is not None:
                
                self.vendor_data_dictionary = database.get_vendor_data(self.vendor.get())

                self.contractEnd = datetime.datetime.strptime(self.vendor_data_dictionary['Contract Date'], '%Y-%m-%d').date()  + datetime.timedelta(days=365*5)
                
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
            self.vendor.get() and self.po.get() and self.selectedStation.get() and self.contractStart.get() and self.operator_name.get() and self.gst.get() and self.pan.get() and self.attendencePath.get()):
            # All fields are filled, perform the add mandays operation
            self.perform_bill_operation()
        else:
            # Display an error message if any field is empty
            messagebox.showerror("Error", "Please fill in all the fields.")

    def perform_bill_operation(self):
        pass


        
         
         

        
