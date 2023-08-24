from tkinter import *
from tkinter import ttk
from ttkthemes import themed_tk as tk
# from PIL import ImageTk, Image
from tkinter import messagebox
from tkinter import filedialog
import dashboard
import database
from datetime import datetime
import os, sys

class MergeView:
    def __init__(self, window):
        self.window = window
        window.geometry("1366x768")
        window.resizable(0, 0)
        self.window.state('zoomed')
        window.title("COMBINE BILLS")
        self.txt = "UPLOAD BILLS TO COMBINE"
        self.color = ["#4f4e4d", "#f29844", "red2"]
        self.heading = Label(self.window, text=self.txt, font=('yu gothic ui', 30, "bold"), bg="white",
                             fg='black',
                             bd=5,
                             relief=FLAT)
        self.heading.place(x=430, y=100, width=600)

        # ========== DECLARING VARIABLES ==================
        self.bill1 = StringVar()
        self.bill2 = StringVar()
        self.bill3 = StringVar()

        # Call the create_segment method for each bill
        self.create_segment("BILL 1:", 318, self.bill1)
        self.create_segment("BILL 2:", 398, self.bill2)
        self.create_segment("BILL 3:", 478, self.bill3)

        self.gen_btn = Button(self.window, text="GENERATE BILL", width=15, height=1, cursor="hand2", font=("yu gothic ui", 20, "bold"), command=self.generate)
        self.gen_btn.place(x=600, y=600)

    def create_segment(self, label_text, y_offset, bill_variable):
        label = Label(self.window, text=label_text, bg="white", fg="#4f4e4d", font=("yu gothic ui", 20, "bold"))
        label.place(x=120, y=y_offset, height=35)

        entry = Entry(self.window, highlightthickness=1, highlightbackground="black", relief=FLAT, bg="white", fg="#6b6a69",
                      font=("yu gothic ui semibold", 15), textvariable=bill_variable)
        entry.place(x=220, y=y_offset + 2, width=850)

        btn = Button(self.window, text="UPLOAD", width=10, height=1, cursor="hand2", font=("yu gothic ui", 14, "bold"), command=lambda: self.upload(bill_variable))
        btn.place(x=1100, y=y_offset - 6)

    def upload(self, bill_variable):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
            bill_variable.set(file_path)
        except Exception as e:
            print(f"An error occurred: {e} at line {sys.exc_info()[-1].tb_lineno}")

    def generate(self):
        pass


        #     # Keep track of the number of segments added
        #     self.segment_count = 0

        #     # Store widgets for each segment
        #     self.segments = {}

        #     # Create the "ADD MORE" button
        #     self.add_more_btn = Button(self.window, text="ADD BILL", width=15, height=1, cursor="hand2", font=("yu gothic ui", 14, "bold"), command=self.add_more_segment)
        #     self.add_more_btn.place(x=300, y=400)  # Initial position

    # def add_more_segment(self):
    #     if self.segment_count <= 3:  # You can adjust the maximum number of segments
    #         y_offset = 320 + self.segment_count * 80  # Adjust the y-coordinate for new elements

    #         # Create new widgets for the additional segment
    #         new_label = Label(self.window, text=f"BILL {self.segment_count + 1}:", bg="white", fg="#4f4e4d",
    #                         font=("yu gothic ui", 20, "bold"))
    #         new_label.place(x=120, y=y_offset, height=35)

    #         new_entry = Entry(self.window, highlightthickness=1, highlightbackground="black", relief=FLAT, bg="white", fg="#6b6a69",
    #                         font=("yu gothic ui semibold", 15), textvariable=self.bill1, state="readonly")
    #         new_entry.place(x=220, y=y_offset + 2, width=650)

    #         new_btn = Button(self.window, text="UPLOAD", width=10, height=1, cursor="hand2", font=("yu gothic ui", 14, "bold"), command=lambda: self.upload())
    #         new_btn.place(x=900, y=y_offset - 6)

    #         # Create a "DELETE BILL" button for the new segment
    #         delete_btn = Button(self.window, text="DELETE BILL", width=15, height=1, cursor="hand2", font=("yu gothic ui", 12, "bold"), command=lambda idx=self.segment_count: self.delete_segment(idx))
    #         delete_btn.place(x=1050, y=y_offset + 6)

    #         # Adjust the position of the "ADD MORE" button
    #         self.add_more_btn.place(y=y_offset + 80)

    #         # Store widgets in the segments dictionary
    #         self.segments[self.segment_count] = {
    #             "label": new_label,
    #             "entry": new_entry,
    #             "upload_btn": new_btn,
    #             "delete_btn": delete_btn
    #         }            

    #         # Increment the segment count
    #         self.segment_count += 1

    #         if self.segment_count == 3:
    #             self.add_more_btn.config(state='disabled')

    #     else:
    #         # Remove the "ADD MORE" button after reaching the maximum number of segments
    #         self.add_more_btn.place_forget()

    # def delete_segment(self, segment_index):
    #     if segment_index in self.segments:  # Ensure the segment exists

    #         # Get widgets for the specified segment
    #         widgets_to_remove = self.segments[segment_index]

    #         for widget in widgets_to_remove.values():
    #             widget.destroy()

    #         # Remove the segment from the segments dictionary
    #         del self.segments[segment_index]

    #         # Identify the y-coordinate of the segment to be deleted
    #         y_offset = 320 + (segment_index - 1) * 80

    #         # Adjust the position of the "ADD MORE" button
    #         self.add_more_btn.place(y=y_offset + 80)

    #         # Adjust the positions of the remaining segments
    #         for idx in range(segment_index + 1, self.segment_count + 1):
    #             segment_widgets = self.segments[idx]

    #             for widget in segment_widgets.values():
    #                 widget.place_configure(y=widget.winfo_y() - 80)

    #         # Decrement the segment count
    #         self.segment_count -= 1

    #         # Update the state of the "ADD BILL" button
    #         self.add_more_btn.config(state='active')


def win():
    window = tk.ThemedTk()
    window.get_themes()
    window.set_theme("arc")
    MergeView(window)
    window.mainloop()


if __name__ == '__main__':
    win()