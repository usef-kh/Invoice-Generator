import os
import tkinter as tk
from tkinter import ttk

import pandas as pd
from win32com import client


class Module:

    def __init__(self, root):
        self.root = root
        self.root.wm_iconbitmap('Graphics/logo.ico')

        self.template_path = os.path.abspath('Templates.xlsx')
        self.final_file_path = None
        self.module_name = None

        self.master = tk.LabelFrame(self.root, borderwidth=0, highlightthickness=0)
        self.master.pack()
        self.load_database()
        self.build_window()

    def start(self):
        self.root.mainloop()

    def load_database(self):
        self.database = pd.read_excel("database.xlsx", sheet_name=None, index_col=None)

    def build_window(self):
        raise NotImplementedError

    def place_admin_fields(self):
        frame = tk.LabelFrame(self.master, text='Admin Info', font=('Arial', 14), padx=10, pady=10, borderwidth=0,
                              highlightthickness=0)

        tk.Label(frame, text="Name", font=('Arial', 10), padx=10, width=7, anchor=tk.W).grid(row=0, column=0, padx=10)
        self.admin = tk.Entry(frame, fg="black", bg="white", width=33)
        self.admin.grid(row=0, column=1, padx=10)

        frame.pack(padx=20, pady=(20, 0), anchor=tk.W)

    def place_customer_fields(self):
        frame = tk.LabelFrame(self.master, text='Customer Info', width=10, font=('Arial', 14), padx=10, pady=10,
                              borderwidth=0, highlightthickness=0)

        entries = []
        for i, text in enumerate(['Name', 'Company', 'Credit']):
            tk.Label(frame, text=text, font=('Arial', 10), padx=10, width=7, anchor=tk.W).grid(row=i, column=0, padx=10)
            entry = tk.Entry(frame, fg="black", bg="white", width=33)
            entry.grid(row=i, column=1, padx=10)
            entries.append(entry)

        self.customer, self.company, self.credit = entries
        frame.pack(padx=20, pady=2, anchor=tk.W)

    def place_item_reservation_fields(self):
        frame = tk.LabelFrame(self.master, text='Item Reservation', width=10, font=('Arial', 14), padx=10, pady=10,
                              borderwidth=0, highlightthickness=0)

        self.table_frame = tk.LabelFrame(frame, width=10, font=('Arial', 14), borderwidth=0, highlightthickness=0)

        self.table_frame.pack(pady=5)

        button_frame = tk.LabelFrame(frame, width=10, font=('Arial', 14), borderwidth=0, highlightthickness=0)

        # Creating a photoimage object to use image
        img = tk.PhotoImage(file="Graphics/plus.png")
        img = img.subsample(64)

        button = tk.Button(button_frame, image=img, command=self.add_table_row, borderwidth=0)
        button.image = img
        button.pack()
        button_frame.pack(side=tk.RIGHT)

        for i, text in enumerate(['Item', 'Rate', 'Quantity']):
            tk.Label(self.table_frame, text=text, font=('Arial', 12), padx=10, width=6).grid(row=0, column=i, padx=10)

        self.counter = 1
        self.table_fields = dict()
        self.add_table_row()

        frame.pack(padx=20, pady=2, anchor=tk.W)

    def add_table_row(self):
        def set_rate(event):
            table_entry = self.table_fields[event.widget]
            choice = event.widget.get()

            table_entry[0].delete(0, "end")
            table_entry[0].insert(0, self.ITEMS['Rate'][choice])

            table_entry[1].delete(0, "end")
            table_entry[1].insert(0, 1)

        item_field = ttk.Combobox(self.table_frame, values=list(self.ITEMS.index), font=('Arial', 9), width=25)
        item_field.grid(row=self.counter, column=0, sticky="NSEW")

        item_field.bind("<<ComboboxSelected>>", set_rate)

        rate_field = tk.Entry(self.table_frame, fg="black", bg="white", font=('Arial', 11), width=3)
        rate_field.grid(row=self.counter, column=1, sticky="NSEW")

        qty_field = tk.Entry(self.table_frame, fg="black", bg="white", font=('Arial', 11), width=4)
        qty_field.grid(row=self.counter, column=2, sticky="NSEW")
        self.table_fields[item_field] = (rate_field, qty_field)

        self.counter += 1
        self.root.geometry(str(self.root.winfo_width()) + 'x' + str(self.root.winfo_height() + 20))

        # print(self.root.winfo_height())

    def place_notes(self):
        frame = tk.LabelFrame(self.master, text='Notes', font=('Arial', 14), padx=10, pady=10, borderwidth=0,
                              highlightthickness=0)
        self.notes = tk.Text(frame, width=49, height=2)
        self.notes.pack()

        frame.pack(padx=20, pady=2, anchor=tk.W)

    def read_inputs(self, *args):
        raise NotImplementedError

    def preview(self):
        self.save()
        if self.final_file_path:
            os.remove(self.final_file_path)

    def clear(self, *args):
        raise NotImplementedError

    def save(self, choice='PDF'):

        inputs = self.read_inputs()
        if isinstance(inputs, Exception):
            tk.messagebox.showinfo("Error", message=inputs)
            return

        target, entries, hidden_entries = inputs

        # paths
        target_path = self.template_path[:-14] + target
        try:
            # open excel
            app = client.DispatchEx("Excel.Application")
            app.Interactive = False
            app.Visible = False

            # load template and open required sheet
            Workbook = app.Workbooks.Open(self.template_path)
            Workbook.WorkSheets(self.module_name).Select()
            Worksheet = Workbook.WorkSheets(self.module_name)

            conv = {
                'A': 1,
                'B': 2,
                'C': 3,
                'D': 4,
                'E': 5,
                'F': 6,
                'G': 7,
                'H': 8,
                'I': 9,
                'K': 10
            }

            # Place Information in cells
            for cell, info in entries:
                column, row = conv[cell[0]], cell[1:]
                Worksheet.Cells(row, column).Value = info

            for cell, info in hidden_entries:
                column, row = conv[cell[0]], cell[1:]
                if info != 0:
                    Worksheet.Rows(row).Hidden = False
                    Worksheet.Cells(row, column).Value = info
                else:
                    Worksheet.Rows(row).Hidden = True

            # Insert Logo
            logo_path = os.path.abspath("Graphics\logo.png")
            Worksheet.Shapes.AddPicture(logo_path, True, True, 0, 15, 60, 65)

            # Save as PDF AND Discard Changes to template
            if choice == 'PDF':
                self.final_file_path = target_path + '.pdf'
                Workbook.SaveAs(self.final_file_path, FileFormat=57)

            elif choice == 'Excel':
                self.final_file_path = target_path + '.xlsx'
                Workbook.SaveAs(self.final_file_path)

            os.startfile(self.final_file_path)

        except Exception as e:
            print("Failed to convert")
            print(str(e))

        finally:
            Workbook.Close(SaveChanges=False)
            app = None


if __name__ == '__main__':
    window = tk.Tk()
    Module(window)
