from tkinter import *
from tkinter import ttk
# from PIL import ImageTk, Image
from tkinter import messagebox
import database
import datetime
import addMandays
import dashboard


class MandaysView:
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

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=('yu gothic ui', 10, "bold"), foreground="black",
                        background="#108cff")
        
        # ============ INPUT VARIABLES =============================
        self.vendor = StringVar()
        self.po = StringVar()
        self.station = []
        self.selectedStation = StringVar()
        self.contractStart = StringVar()
        self.contractEnd = StringVar()
        self.operator = StringVar()

        
        # ============ VENDOR NAME =================================
        self.vendor_label = Label(self.window, text="Vendor Name:")
        self.vendor_label.grid(row=1, column=0, padx=(50, 10), pady=10)

        databaseConnection = database.connectSQL()
        dbCursor = databaseConnection.cursor()
        dbCursor.execute('SELECT DISTINCT Vendor_Name FROM vendor')
        results = dbCursor.fetchall()
        self.vendorNames = [row[0] for row in results]
        dbCursor.close()
        databaseConnection.close()

        
        self.vendor_combo = ttk.Combobox(self.window, values= self.vendorNames, textvariable=self.vendor, state='readonly')
        self.vendor_combo.grid(row=1, column=1, padx=(50, 10), pady=10)
        self.vendor_combo.bind('<<ComboboxSelected>>', self.fetchData)

        self.po_label = Label(self.window, text="PO No.:")
        self.po_label.grid(row=2, column=0, padx=(50, 10), pady=10)

        self.po_no = Entry(self.window, textvariable=self.po)
        self.po_no.grid(row=2, column=1, padx=(50, 10), pady=10)

        self.station_label = Label(self.window, text="Station Name:")
        self.station_label.grid(row=3, column=0, padx=(50, 10), pady=10)

        self.station_name = ttk.Combobox(self.window, values=[" "], state='readonly', textvariable=self.selectedStation)
        self.station_name.grid(row=3, column=1, padx=(50, 10), pady=10)

        self.contractStart_label = Label(self.window, text="Contract Start Date:")
        self.contractStart_label.grid(row=4, column=0, padx=(50, 10), pady=10)

        self.contractStart_date = Entry(self.window, textvariable=self.contractStart)
        self.contractStart_date.grid(row=4, column=1, padx=(50, 10), pady=10)

        self.operator_label = Label(self.window, text="Operator Name:")
        self.operator_label.grid(row=6, column=0, padx=(50, 10), pady=10)

        self.operator_name = Entry(self.window, textvariable=self.operator)
        self.operator_name.grid(row=6, column=1, padx=(50, 10), pady=10)

        self.addMandays_btn = Button(self.window, text="Add Mandays", command=lambda: self.add_mandays())
        self.addMandays_btn.grid(row=7, column=0, columnspan=2, padx=(50, 10), pady=10)

        self.back_btn = Button(self.window, text="BACK", command=lambda: self.back_operation())
        self.back_btn.grid(row=7, column=1, columnspan=2, padx=(50, 10), pady=10)

    def fetchData(self, event):

        if self.vendor.get() is not None:
            try:    
                self.vendor_data_dictionary = database.get_vendor_data(self.vendor.get())
                
                self.contractEnd = datetime.datetime.strptime(self.vendor_data_dictionary['Contract Date'], '%Y-%m-%d').date()  + datetime.timedelta(days=365*5)
                
                self.po_no.config(state='normal')
                self.po_no.delete(0, END)
                self.po_no.insert(0, self.vendor_data_dictionary['PO No'])
                self.po_no.config(state='readonly')
                
                self.station_name.delete(0, END)
                self.station_name.config(values=self.vendor_data_dictionary['Station Name'])
                self.station_name.current(0)

                self.contractStart_date.config(state='normal')
                self.contractStart_date.delete(0, END)
                self.contractStart_date.insert(0, self.vendor_data_dictionary['Contract Date'])
                self.contractStart_date.config(state='readonly')

                self.operator_name.config(state='normal')
                self.operator_name.delete(0, END)
                self.operator_name.insert(0, self.vendor_data_dictionary['Operator Name'])
                self.operator_name.config(state='readonly')
            except Exception as e:
                print(e)

        else:
            # Display an error message if any field is empty
            messagebox.showerror("Error", "Please fill in all the fields.")

    def add_mandays(self):
        
        if ( self.vendor.get() and self.po.get() and self.selectedStation.get() and self.contractStart.get() and self.operator.get() ):
            # All fields are filled, perform the add mandays operation
            self.perform_add_mandays_operation()
        else:
            # Display an error message if any field is empty
            messagebox.showerror("Error", "Please fill in all the fields.")

    def back_operation(self):
        win = Toplevel()
        dashboard.Dashboard(win)
        self.window.withdraw()
        win.deiconify()
            
        
    def perform_add_mandays_operation(self):

        winAddMandays = Toplevel()
        addMandays.AddMandays(winAddMandays, self.contractEnd, self.vendor, self.selectedStation)
        self.window.withdraw()
        winAddMandays.deiconify()
            
        
