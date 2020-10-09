import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox


def make_connection(db='Database.db'):
    conn = sqlite3.connect(db)
    return conn


root = tk.Tk()
main_tabs = ttk.Notebook(root)
main_tabs.pack()

customers_tab = ttk.Frame(main_tabs)
morespace_tab = ttk.Frame(main_tabs)
mediaspace_tab = ttk.Frame(main_tabs)

main_tabs.add(customers_tab, text='Customers')
main_tabs.add(morespace_tab, text='More Space')
main_tabs.add(mediaspace_tab, text='Media Space')
main_tabs.pack(expand=True, fill="both")

morespace_child_tabs = ttk.Notebook(morespace_tab)
morespace_child_tabs.pack(expand=True, fill="both")

morespace_spaces_tab = ttk.Frame(morespace_child_tabs)
morespace_items_tab = ttk.Frame(morespace_child_tabs)
morespace_extras_tab = ttk.Frame(morespace_child_tabs)

morespace_child_tabs.add(morespace_spaces_tab, text='Spaces')
morespace_child_tabs.add(morespace_items_tab, text='Items')
morespace_child_tabs.add(morespace_extras_tab, text='Extras')

mediaspace_child_tabs = ttk.Notebook(mediaspace_tab)
mediaspace_child_tabs.pack(expand=True, fill="both")

mediaspace_spaces_tab = ttk.Frame(mediaspace_child_tabs)
mediaspace_items_tab = ttk.Frame(mediaspace_child_tabs)
mediaspace_extras_tab = ttk.Frame(mediaspace_child_tabs)

mediaspace_child_tabs.add(mediaspace_spaces_tab, text='Spaces')
mediaspace_child_tabs.add(mediaspace_items_tab, text='Items')
mediaspace_child_tabs.add(mediaspace_extras_tab, text='Extras')


def load_database(frame, sheet):
    def submit():
        def confirm():

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

        confirm_button = tk.Button(window, text="Confirm", command=confirm)
        confirm_button.grid(row=i + 2, column=0, columnspan=2)

    def delete():

        if tree.selection():
            check = messagebox.askquestion("Confirm", message="Are you sure you want to delete selection?")

            if check == 'yes':
                conn = make_connection()
                c = conn.cursor()

                for selection in tree.selection():
                    record = tree.item(selection)['values']
                    key = record[0]
                    c.execute("delete from " + sheet + " where oid =" + str(key))

                    tree.delete(selection)

                conn.commit()

    conn = make_connection()
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
    add_button = tk.Button(frame, text='Add Entry', command=submit)
    delete_button.pack(side=tk.LEFT, padx=10)
    add_button.pack(side=tk.RIGHT, padx=10)


load_database(customers_tab, 'customers')
load_database(morespace_spaces_tab, 'morespace_spaces')
load_database(morespace_items_tab, 'morespace_items')
load_database(morespace_extras_tab, 'morespace_extras')

load_database(mediaspace_spaces_tab, 'mediaspace_spaces')
load_database(mediaspace_items_tab, 'mediaspace_items')
load_database(mediaspace_extras_tab, 'mediaspace_extras')

conn = make_connection()

c = conn.cursor()

# records = c.execute("select * from morespace_spaces where code=" + "'A'")
# print(list(records))

root.mainloop()
