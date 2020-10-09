import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox


class DatabaseModule:

    def __init__(self, root, db='Database.db'):
        self.database_name = db
        self.root = root
        self.conn = sqlite3.connect(db)
        self.cursor = self.conn.cursor()

    def load_sheet(self, frame, sheet):
        def add():
            def submit():

                string, placements = '', {}
                for i, entry in enumerate(entries):
                    string += ':a' + str(i) + ','
                    placements['a' + str(i)] = entry.get()

                string = string[:-1]

                conn = sqlite3.connect('Database.db')
                c = conn.cursor()

                # Insert into table
                command = "INSERT INTO " + sheet + " VALUES(" + string + ")"
                c.execute(command, placements)

                # Commit Changes
                conn.commit()

                select = c.execute("SELECT oid, * FROM " + sheet + " order by oid")
                select = list(select)
                tree.insert('', tk.END, value=select[-1])
                conn.close()

                for entry in entries:
                    entry.delete(0, tk.END)

            window = tk.Tk()
            entries = []
            for i, txt in enumerate(columns[1:]):
                tk.Label(window, text=txt, font=('Arial', 10), padx=10, anchor=tk.W).grid(row=i + 1, column=0,
                                                                                          padx=10)
                entry = tk.Entry(window, fg="black", bg="white", width=30)
                entry.grid(row=i + 1, column=1, padx=10)

                entries.append(entry)

            submit_button = tk.Button(window, text="Submit", command=submit)
            submit_button.grid(row=i + 2, column=0, columnspan=2)

        def delete():

            if tree.selection():
                check = messagebox.askquestion("Confirm", message="Are you sure you want to delete selection?")

                if check == 'yes':
                    conn = sqlite3.connect(self.database_name)
                    c = conn.cursor()

                    for selection in tree.selection():
                        record = tree.item(selection)['values']
                        key = record[0]
                        c.execute("delete from " + sheet + " where oid =" + str(key))

                        tree.delete(selection)

                    conn.commit()

        conn = sqlite3.connect(self.database_name)
        c = conn.cursor()
        records = c.execute("select oid, * from " + sheet)

        columns = [description[0] for description in c.description]

        tree = ttk.Treeview(frame, columns=([str(i) for i in range(len(columns))]), height='10', show='headings',
                            selectmode="extended")

        tree.pack(padx=10, pady=10)

        for i, txt in enumerate(columns):
            tree.heading(str(i), text=txt)
            tree.column(str(i), minwidth=10, width=150)

        for entry in records:
            tree.insert('', tk.END, value=entry)

        delete_button = tk.Button(frame, text='Delete', command=delete)
        add_button = tk.Button(frame, text='Add Entry', command=add)
        delete_button.pack(side=tk.LEFT, padx=10)
        add_button.pack(side=tk.RIGHT, padx=10)


class Customers(DatabaseModule):

    def __init__(self, root):
        super(Customers, self).__init__(root)
        self.module_name = "Customers"

        self.load_sheet(root, 'customers')


class MoreSpace(DatabaseModule):

    def __init__(self, root):
        super(MoreSpace, self).__init__(root)
        self.module_name = "More Space"

        # create tabs
        tabs = ttk.Notebook(root)
        tabs.pack(expand=True, fill="both")

        spaces_tab = ttk.Frame(tabs)
        items_tab = ttk.Frame(tabs)
        extras_tab = ttk.Frame(tabs)

        tabs.add(spaces_tab, text='Spaces')
        tabs.add(items_tab, text='Items')
        tabs.add(extras_tab, text='Extras')

        # load sheets into tabs
        all_tabs = [spaces_tab, items_tab, extras_tab]
        all_sheets = ['morespace_spaces', 'morespace_items', 'morespace_extras']

        for tab, sheet in zip(all_tabs, all_sheets):
            self.load_sheet(tab, sheet)


class MediaSpace(DatabaseModule):

    def __init__(self, root):
        super(MediaSpace, self).__init__(root)
        self.module_name = "Media Space"

        # create tabs
        tabs = ttk.Notebook(root)
        tabs.pack(expand=True, fill="both")

        spaces_tab = ttk.Frame(tabs)
        items_tab = ttk.Frame(tabs)
        extras_tab = ttk.Frame(tabs)

        tabs.add(spaces_tab, text='Spaces')
        tabs.add(items_tab, text='Items')
        tabs.add(extras_tab, text='Extras')

        # load sheets into tabs
        all_tabs = [spaces_tab, items_tab, extras_tab]
        all_sheets = ['mediaspace_spaces', 'mediaspace_items', 'mediaspace_extras']

        for tab, sheet in zip(all_tabs, all_sheets):
            self.load_sheet(tab, sheet)


if __name__ == '__main__':
    root = tk.Tk()
    tabs = ttk.Notebook(root)
    tabs.pack(expand=True, fill="both")

    modules = [Customers, MoreSpace, MediaSpace]
    for module in modules:
        tab = ttk.Frame(tabs)
        instance = module(tab)

        tabs.add(tab, text=instance.module_name)

    root.mainloop()
