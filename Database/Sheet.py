import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox


class Sheet:

    def __init__(self, sheet_name='', db='Database.db'):
        self.conn = sqlite3.connect(db)
        self.cursor = self.conn.cursor()

        self.cursor.execute("select * from " + sheet_name)
        self.columns = [description[0] for description in self.cursor.description]
        self.sheet_name = sheet_name

    def get_table(self):
        return list(self.cursor.execute("select * from " + self.sheet_name))

    def insert(self, record):
        assert len(record) == len(self.columns)

        command = "insert into " + self.sheet_name + " values("
        placements = {}
        for i, info in enumerate(record):
            command += ':a' + str(i) + ','
            placements['a' + str(i)] = info

        command = command[:-1] + ")"

        self.cursor.execute(command, placements)
        self.conn.commit()

    def delete(self, record):
        assert len(record) == len(self.columns)

        command = "delete from " + self.sheet_name + " where "
        for col, val in zip(self.columns[:-1], record[:-1]):
            command += col + "=\"" + str(val) + "\" AND "
        command += self.columns[-1] + "=\"" + str(record[-1]) + "\""

        self.cursor.execute(command)
        self.conn.commit()

    def edit(self):
        pass


class GUISheet(Sheet):

    def __init__(self, sheet_name, master):
        super(GUISheet, self).__init__(sheet_name)

        self.master = master
        self.tree = ttk.Treeview(self.master, columns=[str(i) for i in range(len(self.columns))],
                                 height='10', show='headings', selectmode="extended")

        self.tree.pack(padx=10, pady=10)

        for i, column in enumerate(self.columns):
            self.tree.heading(str(i), text=column)
            self.tree.column(str(i), minwidth=10, width=150)

        for record in self.get_table():
            self.tree.insert('', tk.END, value=record)

        add_button = ttk.Button(self.master, text='Add Entry', command=self.insert)
        edit_button = ttk.Button(self.master, text='Edit', command=self.edit)
        delete_button = ttk.Button(self.master, text='Delete', command=self.delete)

        add_button.pack(side=tk.RIGHT, padx=10, pady=(0, 10))
        edit_button.pack(side=tk.RIGHT, padx=10, pady=(0, 10))
        delete_button.pack(side=tk.RIGHT, padx=10, pady=(0, 10))

    def insert(self):

        def submit():
            record = []
            for entry in entries:
                record.append(entry.get())

            # Insert into database
            super(GUISheet, self).insert(record)

            # Insert into tree
            self.tree.insert('', tk.END, value=record)

        window = tk.Tk()
        info_frame = ttk.Frame(window)
        button_frame = ttk.Frame(window)

        info_frame.pack(padx=20, pady=(20, 0))
        button_frame.pack(anchor=tk.E, padx=20, pady=(0, 20))

        entries = []
        for i, txt in enumerate(self.columns):
            tk.Label(info_frame, text=txt, font=('Arial', 10), anchor=tk.W).grid(row=i + 1, column=0, padx=10,
                                                                                 sticky=tk.W)
            entry = ttk.Entry(info_frame, width=30)
            entry.grid(row=i + 1, column=1, padx=10)
            entries.append(entry)

        submit_button = ttk.Button(button_frame, text="Submit", command=submit)
        submit_button.pack(side=tk.RIGHT, padx=10, pady=(10, 0))

    def delete(self):
        if self.tree.selection():
            check = messagebox.askquestion("Confirm", message="Are you sure you want to delete selection?")

            if check == 'yes':
                for selection in self.tree.selection():
                    record = self.tree.item(selection)['values']

                    # delete from database
                    super(GUISheet, self).delete(record)

                    # delete from tree
                    self.tree.delete(selection)

    def edit(self):
        pass

    def run(self):
        self.master.mainloop()


if __name__ == '__main__':
    sheet = GUISheet('morespace_spaces', tk.Tk())
    sheet.run()