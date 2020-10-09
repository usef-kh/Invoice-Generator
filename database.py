import sqlite3
import tkinter as tk

root = tk.Tk()
root.geometry("400x400")

conn = sqlite3.connect('database.db')

# create cursor
cursor = conn.cursor()


# create table
# cursor.execute("""CREATE TABLE Customers (
#         name text,
#         company real,
#         email text
# )
# """)


def submit():
    conn = sqlite3.connect('database.db')
    cursor = conn.cursor()

    # insert into table

    cursor.execute("INSERT INTO Customers VALUES (:customer, :company, :email)",
                   {
                       'customer': customer.get(),
                       'company': company.get(),
                       'email': email.get()
                   }
                   )

    # commit changes
    conn.commit()
    conn.close()

    customer.delete(0, tk.END)
    company.delete(0, tk.END)
    email.delete(0, tk.END)


def query():
    conn = sqlite3.connect('database.db')
    cursor = conn.cursor()

    cursor.execute("SELECT *, oid FROM Customers WHERE company='Individual'")
    records = cursor.fetchall()

    print(records)

    # commit changes
    conn.commit()
    conn.close()


tk.Label(text='Customer').grid(row=1, column=0)
customer = tk.Entry(root, width=30)
customer.grid(row=1, column=1, padx=20)

tk.Label(text='Comapny').grid(row=2, column=0)
company = tk.Entry(root, width=30)
company.grid(row=2, column=1)

tk.Label(text='Email').grid(row=3, column=0)
email = tk.Entry(root, width=30)
email.grid(row=3, column=1)

add_entry = tk.Button(root, text="Add Record", command=submit)
add_entry.grid(row=4, column=1)

show_records = tk.Button(root, text="Show Records", command=query)
show_records.grid(row=5, column=1)

# commit changes
conn.commit()

conn.close()

root.mainloop()
