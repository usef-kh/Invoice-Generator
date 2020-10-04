import tkinter as tk
import os
import pandas as pd


class Module:

    def __init__(self, root):
        self.root = root
        self.root.wm_iconbitmap('logo.ico')

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