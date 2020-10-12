import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox


class DatabaseModule:

    def __init__(self, root):
        self.root = root


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

                # Insert into tree
                tree.insert('', tk.END, value=list(placements.values()))

                # Commit Changes
                conn.commit()
                c.close()

                # Clear entry fields
                for entry in entries:
                    entry.delete(0, tk.END)

            window = tk.Tk()
            info_frame = ttk.Frame(window)
            button_frame = ttk.Frame(window)

            info_frame.pack(padx=20, pady=(20, 0))
            button_frame.pack(anchor=tk.E, padx=20, pady=(0, 20))

            entries = []
            for i, txt in enumerate(columns):
                tk.Label(info_frame, text=txt, font=('Arial', 10), anchor=tk.W).grid(row=i + 1, column=0, padx=10,
                                                                                     sticky=tk.W)
                entry = ttk.Entry(info_frame, width=30)
                entry.grid(row=i + 1, column=1, padx=10)
                entries.append(entry)

            submit_button = ttk.Button(button_frame, text="Submit", command=submit)
            submit_button.pack(side=tk.RIGHT, padx=10, pady=(10, 0))

        def delete():
            if tree.selection():
                check = messagebox.askquestion("Confirm", message="Are you sure you want to delete selection?")

                if check == 'yes':
                    conn = sqlite3.connect(self.database_name)
                    c = conn.cursor()

                    for selection in tree.selection():
                        record = tree.item(selection)['values']

                        command = "delete from " + sheet + " where "
                        for col, val in zip(columns[:-1], record[:-1]):
                            command += col + "=\"" + str(val) + "\" AND "
                        command += columns[-1] + "=\"" + str(record[-1]) + "\""

                        c.execute(command)
                        tree.delete(selection)

                    conn.commit()
                    c.close()

        def edit():
            def submit():

                conn = sqlite3.connect('Database.db')
                c = conn.cursor()

                query = "select *  from " + sheet + " where oid =" + str(key)
                record = c.execute(query)
                record = list(record)[0]

                for i, column in enumerate(c.description):
                    column_name = column[0]

                    new_value = entries[i].get()
                    old_value = record[i]

                    if old_value != new_value:
                        query = "update " + sheet + " set " + column_name + " = ? where oid = ?"
                        data = (new_value, key)
                        c.execute(query, data)

                conn.commit()
                c.close()

            if tree.selection():

                if len(tree.selection()) > 1:
                    messagebox.showinfo("Error", message="You can only update one entry at a time")
                    return

                record = tree.item(tree.selection()[0])['values']

                key = record[0]

                window = tk.Tk()
                info_frame = ttk.Frame(window)
                button_frame = ttk.Frame(window)

                info_frame.pack(padx=20, pady=(20, 0))
                button_frame.pack(anchor=tk.E, padx=20, pady=(0, 20))

                entries = []
                for i, txt in enumerate(columns):
                    tk.Label(info_frame, text=txt, font=('Arial', 10), anchor=tk.W).grid(row=i + 1, column=0, padx=10, sticky=tk.W)
                    entry = ttk.Entry(info_frame, width=30)
                    entry.grid(row=i + 1, column=1, padx=10)
                    entries.append(entry)

                for entry, value in zip(entries, record):
                    entry.insert(0, value)

                submit_button = ttk.Button(button_frame, text="Submit", command=submit)
                submit_button.pack(side=tk.RIGHT, padx=10, pady = (10, 0))

        conn = sqlite3.connect(self.database_name)
        c = conn.cursor()
        records = c.execute("select * from " + sheet)

        columns = [description[0] for description in c.description]

        tree = ttk.Treeview(frame, columns=([str(i) for i in range(len(columns))]), height='10', show='headings',
                            selectmode="extended")

        tree.pack(padx=10, pady=10)

        for i, txt in enumerate(columns):
            tree.heading(str(i), text=txt)
            tree.column(str(i), minwidth=10, width=150)

        for entry in records:
            tree.insert('', tk.END, value=entry)

        add_button = ttk.Button(frame, text='Add Entry', command=add)
        edit_button = ttk.Button(frame, text='Edit', command=edit)
        delete_button = ttk.Button(frame, text='Delete', command=delete)

        add_button.pack(side=tk.RIGHT, padx=10)
        edit_button.pack(side=tk.RIGHT, padx=10)
        delete_button.pack(side=tk.RIGHT, padx=10)

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
        all_sheets = ['morespace_spaces', 'morespace_items', 'morespace_extras']

        for tab, sheet in zip(tabs.children.values(), all_sheets):
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
        all_sheets = ['mediaspace_spaces', 'mediaspace_items', 'mediaspace_extras']

        for tab, sheet in zip(tabs.children.values(), all_sheets):
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
