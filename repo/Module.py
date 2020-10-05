import os
import tkinter as tk
from tkinter import ttk

import pandas as pd


class Module:

    def __init__(self, root):
        self.root = root
        self.root.wm_iconbitmap('Graphics/logo.ico')

        self.target_xlsx = None
        self.target_pdf = None

        self.master = tk.LabelFrame(self.root, borderwidth=0, highlightthickness=0)
        self.master.pack()
        self.load_database()
        self.build_window()

    def start(self):
        self.root.mainloop()

    def load_database(self):
        self.database = pd.read_excel("../database.xlsx", sheet_name=None, index_col=None)

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

        button = tk.Button(button_frame, text="Add Row", command=self.add_table_row)
        button.pack()
        button_frame.pack()

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

    def generate(self, *args):
        raise NotImplementedError

    def preview(self):
        self.generate()
        if self.target_pdf:
            os.remove(self.target_pdf)

    def clear(self, *args):
        raise NotImplementedError


if __name__ == '__main__':
    window = tk.Tk()
    Module(window)
