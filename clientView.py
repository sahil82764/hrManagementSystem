from tkinter import *
from tkinter import ttk
from ttkthemes import themed_tk as tk
# from PIL import ImageTk, Image
from tkinter import messagebox
import dashboard
import database
from datetime import datetime
import sys

class ClientView:
    def __init__(self, window):
        self.window = window
        window.geometry("1366x768")
        window.resizable(0, 0)
        self.window.state('zoomed')
        window.title("CLIENT VIEW")
        self.txt = "MANAGE CLIENT RECORDS"
        self.color = ["#4f4e4d", "#f29844", "red2"]
        self.heading = Label(self.window, text=self.txt, font=('yu gothic ui', 30, "bold"), bg="white",
                             fg='black',
                             bd=5,
                             relief=FLAT)
        self.heading.place(x=430, y=25, width=600)

        #=============== Declaring variables =================================
        self.id = StringVar()
        self.po_no = StringVar()
        self.vendor_code = StringVar()
        self.vendor_name = StringVar()
        self.station_name = StringVar()
        self.contract_date = StringVar()
        self.gst_no = StringVar()
        self.pan_no = StringVar()
        self.operator_name = StringVar()

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview.Heading", font=('yu gothic ui', 10, "bold"), foreground="black",
                        background="#108cff")
        
         # ========================Tree View=============================
        self.scrollbarx = Scrollbar(self.window, orient=HORIZONTAL)
        self.scrollbary = Scrollbar(self.window, orient=VERTICAL)
        self.tree = ttk.Treeview(self.window)
        self.tree.place(relx=0.307, rely=0.203, width=880, height=510)
        self.tree.configure(
            yscrollcommand=self.scrollbary.set, xscrollcommand=self.scrollbarx.set
        )
        self.tree.configure(selectmode="extended")

        self.scrollbary.configure(command=self.tree.yview)
        self.scrollbarx.configure(command=self.tree.xview)

        self.scrollbary.place(relx=0.954, rely=0.203, width=22, height=509)
        self.scrollbarx.place(relx=0.307, rely=0.892, width=884, height=22)

        self.tree.configure(
            columns=(
                "ID",
                "PO No",
                "Vendor Code",
                "Vendor Name",
                "Station Name",
                "Contract Date",
                "GST No",
                "PAN No",
                "Operator Name"
            )
        )

        self.tree.heading("ID", text="ID", anchor="center")
        self.tree.heading("PO No", text="PO No", anchor="center")
        self.tree.heading("Vendor Code", text="Vendor Code", anchor="center")
        self.tree.heading("Vendor Name", text="Vendor Name", anchor="center")
        self.tree.heading("Station Name", text="Station Name", anchor="center")
        self.tree.heading("Contract Date", text="Contract Date", anchor="center")
        self.tree.heading("GST No", text="GST No", anchor="center")
        self.tree.heading("PAN No", text="PAN No", anchor="center")
        self.tree.heading("Operator Name", text="Operator Name", anchor="center")

        self.tree.column("#0", stretch=NO, minwidth=0, width=0)
        self.tree.column("#1", stretch=NO, minwidth=0, width=100)
        self.tree.column("#2", stretch=NO, minwidth=0, width=140)
        self.tree.column("#3", stretch=NO, minwidth=0, width=160)
        self.tree.column("#4", stretch=NO, minwidth=0, width=140)
        self.tree.column("#5", stretch=NO, minwidth=0, width=120)
        self.tree.column("#6", stretch=NO, minwidth=0, width=120)
        self.tree.column("#7", stretch=NO, minwidth=0, width=110)
        self.tree.column("#8", stretch=NO, minwidth=0, width=110)
        self.show_data()
        self.tree.bind("<ButtonRelease-1>", self.client_info)
        # self.tree.bind("<ButtonRelease-1>", self.reset_entry_fields)

        #================= ID ========================================
        self.id_label = Label(self.window, text="ID: ", bg="white", fg="#4f4e4d",
                                  font=("yu gothic ui", 13, "bold"))
        self.id_label.place(x=22, y=220, height=25)

        self.id_entry = Entry(self.window, highlightthickness=0, relief=FLAT, bg="white", fg="#6b6a69",
                                  font=("yu gothic ui semibold", 12), textvariable=self.id, state="readonly")
        self.id_entry.place(x=190, y=220, width=250)  # trebuchet ms

        self.id_line = Canvas(self.window, width=250, height=1.5, bg="#bdb9b1", highlightthickness=0)
        self.id_line.place(x=190, y=243)
        
        #================= PO NO ========================================
        self.po_label = Label(self.window, text="PO No: ", bg="white", fg="#4f4e4d",
                                  font=("yu gothic ui", 13, "bold"))
        self.po_label.place(x=22, y=260, height=25)

        self.po_entry = Entry(self.window, highlightthickness=0, relief=FLAT, bg="white", fg="#6b6a69",
                                  font=("yu gothic ui semibold", 12), textvariable=self.po_no)
        self.po_entry.place(x=190, y=260, width=250)  # trebuchet ms

        self.po_line = Canvas(self.window, width=250, height=1.5, bg="#bdb9b1", highlightthickness=0)
        self.po_line.place(x=190, y=283)

        #================= VENDOR CODE ========================================
        self.code_label = Label(self.window, text="VENDOR CODE: ", bg="white", fg="#4f4e4d",
                                  font=("yu gothic ui", 13, "bold"))
        self.code_label.place(x=22, y=300, height=25)

        self.code_entry = Entry(self.window, highlightthickness=0, relief=FLAT, bg="white", fg="#6b6a69",
                                  font=("yu gothic ui semibold", 12), textvariable=self.vendor_code)
        self.code_entry.place(x=190, y=300, width=250)  # trebuchet ms

        self.code_line = Canvas(self.window, width=250, height=1.5, bg="#bdb9b1", highlightthickness=0)
        self.code_line.place(x=190, y=323)

        #================= VENDOR NAME ========================================
        self.name_label = Label(self.window, text="VENDOR NAME: ", bg="white", fg="#4f4e4d",
                                  font=("yu gothic ui", 13, "bold"))
        self.name_label.place(x=22, y=340, height=25)

        self.name_entry = Entry(self.window, highlightthickness=0, relief=FLAT, bg="white", fg="#6b6a69",
                                  font=("yu gothic ui semibold", 12), textvariable=self.vendor_name)
        self.name_entry.place(x=190, y=340, width=250)  # trebuchet ms

        self.name_line = Canvas(self.window, width=250, height=1.5, bg="#bdb9b1", highlightthickness=0)
        self.name_line.place(x=190, y=363)

        #================= STATION NAME ========================================
        self.station_label = Label(self.window, text="STATION NAME: ", bg="white", fg="#4f4e4d",
                                  font=("yu gothic ui", 13, "bold"))
        self.station_label.place(x=22, y=380, height=25)

        self.station_entry = Entry(self.window, highlightthickness=0, relief=FLAT, bg="white", fg="#6b6a69",
                                  font=("yu gothic ui semibold", 12), textvariable=self.station_name)
        self.station_entry.place(x=190, y=380, width=250)  # trebuchet ms

        self.station_line = Canvas(self.window, width=250, height=1.5, bg="#bdb9b1", highlightthickness=0)
        self.station_line.place(x=190, y=403)

        #================= CONTRACT DATE ========================================
        self.contract_label = Label(self.window, text="CONTRACT DATE: ", bg="white", fg="#4f4e4d",
                                  font=("yu gothic ui", 13, "bold"))
        self.contract_label.place(x=22, y=420, height=25)

        self.contract_entry = Entry(self.window, highlightthickness=0, relief=FLAT, bg="white", fg="#6b6a69",
                                  font=("yu gothic ui semibold", 12), textvariable=self.contract_date)
        self.contract_entry.insert(0, "yyyy-mm-dd")
        self.contract_entry.place(x=190, y=420, width=250)  # trebuchet ms

        self.contract_line = Canvas(self.window, width=250, height=1.5, bg="#bdb9b1", highlightthickness=0)
        self.contract_line.place(x=190, y=443)

        #================= GST NO ========================================
        self.gst_label = Label(self.window, text="GST NO: ", bg="white", fg="#4f4e4d",
                                  font=("yu gothic ui", 13, "bold"))
        self.gst_label.place(x=22, y=460, height=25)

        self.gst_entry = Entry(self.window, highlightthickness=0, relief=FLAT, bg="white", fg="#6b6a69",
                                  font=("yu gothic ui semibold", 12), textvariable=self.gst_no)
        self.gst_entry.place(x=190, y=460, width=250)  # trebuchet ms

        self.gst_line = Canvas(self.window, width=250, height=1.5, bg="#bdb9b1", highlightthickness=0)
        self.gst_line.place(x=190, y=483)

        #================= PAN NO ========================================
        self.pan_label = Label(self.window, text="PAN NO: ", bg="white", fg="#4f4e4d",
                                  font=("yu gothic ui", 13, "bold"))
        self.pan_label.place(x=22, y=500, height=25)

        self.pan_entry = Entry(self.window, highlightthickness=0, relief=FLAT, bg="white", fg="#6b6a69",
                                  font=("yu gothic ui semibold", 12), textvariable=self.pan_no)
        self.pan_entry.place(x=190, y=500, width=250)  # trebuchet ms

        self.pan_line = Canvas(self.window, width=250, height=1.5, bg="#bdb9b1", highlightthickness=0)
        self.pan_line.place(x=190, y=523)

        #================= OPERATOR NAME ========================================
        self.operator_label = Label(self.window, text="OPERATOR NAME: ", bg="white", fg="#4f4e4d",
                                  font=("yu gothic ui", 13, "bold"))
        self.operator_label.place(x=22, y=540, height=25)

        self.operator_entry = Entry(self.window, highlightthickness=0, relief=FLAT, bg="white", fg="#6b6a69",
                                  font=("yu gothic ui semibold", 12), textvariable=self.operator_name)
        self.operator_entry.place(x=190, y=540, width=250)  # trebuchet ms

        self.operator_line = Canvas(self.window, width=250, height=1.5, bg="#bdb9b1", highlightthickness=0)
        self.operator_line.place(x=190, y=563)

        #================= BUTTONS ========================================
        self.submit_btn = Button(self.window, text="UPDATE", width=10, height=1, cursor="hand2", font=("yu gothic ui", 15, "bold"), command=lambda: self.update())
        self.submit_btn.place(x=50, y=600)

        self.add_btn = Button(self.window, text="ADD NEW", width=10, height=1, cursor="hand2", font=("yu gothic ui", 15, "bold"), command=lambda: self.add())
        self.add_btn.place(x=180, y=600)

        self.clear_btn = Button(self.window, text="CLEAR", width=10, height=1, cursor="hand2", font=("yu gothic ui", 15, "bold"), command=lambda: self.clear())
        self.clear_btn.place(x=310, y=600)

        self.delete_btn = Button(self.window, text="DELETE", width=10, height=1, cursor="hand2", font=("yu gothic ui", 15, "bold"), command=lambda: self.delete())
        self.delete_btn.place(x=50, y=670)

        self.exit_btn = Button(self.window, text="EXIT", width=22, height=1, cursor="hand2", font=("yu gothic ui", 15, "bold"), command=lambda: self.exit())
        self.exit_btn.place(x=180, y=670)


    def clear(self):
        self.id_entry.config(state='normal')
        self.id.set("")
        self.id_entry.delete(0, END)
        self.id_entry.config(state='readonly')        
        
        self.po_no.set("")
        self.vendor_code.set("")
        self.vendor_name.set("")
        self.station_name.set("")

        self.contract_entry.delete(0, END)
        self.contract_date.set("")
        self.contract_entry.insert(0, "yyyy-mm-dd")

        self.gst_no.set("")
        self.pan_no.set("")
        self.operator_name.set("")

    def client_info(self, ev):
        viewInfo = self.tree.focus()
        learner_data = self.tree.item(viewInfo)
        row = learner_data["values"]
        self.id.set(row[0])
        self.po_no.set(row[1])
        self.vendor_code.set(row[2])
        self.vendor_name.set(row[3])
        self.station_name.set(row[4])
        self.contract_date.set(row[5])
        self.gst_no.set(row[6])
        self.pan_no.set(row[7])
        self.operator_name.set(row[8])

    def update(self):
        selected_item = self.tree.focus()  # Get the selected item in the Treeview
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select a record to update.")
            return
        if self.po_no.get() == "":
            messagebox.showwarning("No Selection", "Please select a record to update.")
            return
        
        connection = database.connectSQL()
        dbCursor = connection.cursor()
        
        # Get the values from the selected item
        selected_values = self.tree.item(selected_item, "values")

         # Check if the contract date is the placeholder
        if self.contract_date.get() == "yyyy-mm-dd":
            # contract_date = None  # Set contract_date to None
            messagebox.showwarning("INVALID RECORD", "Please Enter Correct Date.")
            self.clear()
            return
        
        # Update only the selected record
        dbCursor.execute("""UPDATE vendor 
                            SET Station_Name=?,
                                Contract_Date=?,
                            WHERE id=?""",
                        (
                            self.station_name.get(),
                            datetime.strptime(self.contract_date.get(), "%Y-%m-%d").date(),
                            int(selected_values[0])  # Use id from the selected row
                        )
                        )
        connection.commit()
        self.show_data()
        connection.close()
        self.clear()
        messagebox.showinfo("", "Client Record Updated Successfully")

    def add(self):
        connection = database.connectSQL()
        dbCursor = connection.cursor()
        dbCursor.execute("INSERT INTO vendor VALUES (?,?,?,?,?,?,?,?)",
                         (int(self.po_no.get()),
                          int(self.vendor_code.get()),
                          self.vendor_name.get().upper(),
                          self.station_name.get().upper(),
                          datetime.strptime(self.contract_date.get(), "%Y-%m-%d").date(),
                          self.gst_no.get(),
                          self.pan_no.get().upper(),
                          self.operator_name.get().upper()
                          )
                         )
        connection.commit()
        connection.close()
        self.show_data()
        self.clear()
        messagebox.showinfo("", "New Client Record Added Successfully")   

    def delete(self):
        try:
            treeViewContent = self.tree.focus()
            treeViewItems = self.tree.item(treeViewContent)
            treeViewValues = treeViewItems["values"][3]
            ask = messagebox.askyesno("Warning",
                                      f"Are you sure you want to delete records of {treeViewValues}")
            
            if ask is True:
                connection = database.connectSQL()
                dbCursor = connection.cursor()
                dbCursor.execute("""DELETE FROM vendor 
                                    WHERE id=?""",
                                (
                                    int(self.id.get())  # Use id from the selected row
                                )
                                )
                connection.commit()
                self.show_data()
                connection.close()
                self.clear()
                messagebox.showinfo("Success", f" {treeViewValues} records has been deleted Successfully")

        except Exception as e:
            print(f"An error occurred: {e} at line {sys.exc_info()[-1].tb_lineno}")
    

    def show_data(self):
        rows = database.get_all()
        if len(rows) != 0:
            self.tree.delete(*self.tree.get_children())
            for row in rows:
                self.tree.insert('', END, values=(row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8]))


    def exit(self):
        exit_command = messagebox.askyesno("Edit Client Records", "Are you sure you want to exit")
        if exit_command > 0:
            win = Toplevel()
            dashboard.Dashboard(win)
            self.window.withdraw()
            win.deiconify()


def win():
    window = tk.ThemedTk()
    window.get_themes()
    window.set_theme("arc")
    ClientView(window)
    window.mainloop()


if __name__ == '__main__':
    win()