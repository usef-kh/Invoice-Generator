import os
import tkinter as tk
from tkcalendar import DateEntry
from datetime import date
from datetime import datetime
import shutil
import openpyxl
import pandas as pd
from tkinter import ttk
from win32com import client
from testing import *


window = tk.Tk()
window.geometry('450x530')
window.title("Invoice Generator")
window.wm_iconbitmap('logo.ico')

database = pd.read_excel("database.xlsx", sheet_name=None, index_col=None)
SPACES = database['Spaces']
SPACES.set_index('Options', inplace=True)

ITEMS = database['Items']
ITEMS.set_index('Item', inplace=True)

ypos = 0
# Customer Info
ypos += 20
header = tk.Label(text='Customer Info', font=('Arial', 14))
header.place(x=20, y=ypos)

ypos += 30
name_text = tk.Label(text='Name', font=('Arial', 10))
name_text.place(x=50, y=ypos)
name_field = tk.Entry(fg="black", bg="white", width=33)
name_field.place(x=120, y=ypos)

ypos += 21
company_text = tk.Label(text='Company', font=('Arial', 10))
company_text.place(x=50, y=ypos)
company_field = tk.Entry(fg="black", bg="white", width=33)
company_field.place(x=120, y=ypos)

# Space Reservation
ypos += 40
header = tk.Label(text='Space Reservation', font=('Arial', 14))
header.place(x=20, y=ypos)

ypos += 30
spaces_text = tk.Label(text='Space(s)', font=('Arial', 10))
spaces_text.place(x=50, y=ypos)

box_value = tk.StringVar()
space_field = AutocompleteCombobox(textvariable=box_value, width=30)

space_field.set_completion_list(list(SPACES.index))
space_field.place(x=120, y=ypos)

space_fields = {space_field: (None, None)}

# def add_space(*args):
#     pass
#
# add_space_button = tk.Button(text="  +  ", font=("Arial", 8))
# add_space_button.place(x=310, y=ypos)
# add_space_button.bind("<Button-1>", add_space)


ypos += 23
start_date_text = tk.Label(text="Dates", font=("Arial", 10))
start_date_text.place(x=50, y=ypos)
start_date_field = DateEntry(window, width=8, background='black', foreground='white', borderwidth=2)
start_date_field.place(x=120, y=ypos)

# ypos += 23
# end_date_text = tk.Label(text="End date", font=("Arial", 10))
# end_date_text.place(x=200, y=ypos)
end_date_field = DateEntry(window, width=8, background='black', foreground='white', borderwidth=2)
end_date_field.place(x=193, y=ypos)

# Item Reservation
ypos += 40
header = tk.Label(text='Item Reservation', font=('Arial', 14))
header.place(x=20, y=ypos)

ypos += 30
item_text = tk.Label(text="Item", font=("Arial", 11))
item_text.place(x=125, y=ypos)

rate_text = tk.Label(text="Rate", font=("Arial", 11))
rate_text.place(x=263, y=ypos)

qty_text = tk.Label(text="Quantity", font=("Arial", 11))
qty_text.place(x=343, y=ypos)

ypos += 20
table_fields = dict()

def set_rate(event):
    table_entry = table_fields[event.widget]
    choice = event.widget.get()

    table_entry[0].delete(0, "end")
    table_entry[0].insert(0, ITEMS['Rate'][choice])

    table_entry[1].delete(0, "end")
    table_entry[1].insert(0, 1)


for i in range(5):
    item_field = ttk.Combobox(window, values=list(ITEMS.index), width=27)
    item_field.place(x=50, y=ypos)

    item_field.bind("<<ComboboxSelected>>", set_rate)

    rate_field = tk.Entry(fg="black", bg="white", width=15)
    rate_field.place(x=233, y=ypos)

    qty_field = tk.Entry(fg="black", bg="white", width=15)
    qty_field.place(x=326, y=ypos)

    table_fields[item_field] = (rate_field, qty_field)

    ypos += 18

# Notes
ypos += 40 - 18
header = tk.Label(text='Notes', font=('Arial', 14))
header.place(x=20, y=ypos)

ypos += 30
notes_field = tk.Entry(fg="black", bg="white", width=65)
notes_field.place(x=25, y=ypos)


# Generate Invoice
def generate(*args):

    def separate(string_date, split_char ='/'):
        string_date = string_date.split(split_char)
        m, d, y = [int(num) for num in string_date]
        return y, m, d

    def reformat(string_date, char='/'):
        date = string_date.split('/')
        for i, entry in enumerate(date):
            if len(entry) == 1:
                date[i] = '0' + entry

        return char.join(date)

    name = name_field.get()
    company = company_field.get()

    invoice_date = datetime.date(datetime.now()).strftime("%m/%d/%y")

    start = start_date_field.get()
    end = end_date_field.get()
    notes = notes_field.get()

    num_of_days = date(*separate(end)) - date(*separate(start))
    num_of_days = num_of_days.days

    template = os.path.abspath('Template.xlsx')
    target = reformat(start, char='-') + "_" + name + '.xlsx'

    # Create duplicate of template
    shutil.copyfile(template, target)

    table_nums = [str(num) for num in range(15, 27)]

    def generate_col_ids(chars):
        res = []

        for num in range(15, 27):
            entry = []
            for char in chars:
                entry.append(char + str(num))
            res.append(entry)

        return res

    placements = {'name': 'A7',
                  'company': 'A8',
                  'invoice_date': 'H6',
                  'due_date': 'H7',
                  'start_date': 'C11',
                  'end_date': 'C12',
                  'table': (generate_col_ids(['A', 'G', 'F'])),
                  'notes': 'A36'}


    workbook = openpyxl.load_workbook(target)
    worksheet = workbook['Sheet1']

    worksheet[placements['name']] = name
    worksheet[placements['company']] = company

    worksheet[placements['invoice_date']] = invoice_date
    worksheet[placements['due_date']] = reformat(end)

    all_entries = {**space_fields, **table_fields}
    for placement, entry in zip(placements['table'], all_entries):

        worksheet[placement[0]] = entry.get()
        rate_field, qty_field = all_entries[entry]

        if (rate_field, qty_field) == (None, None):
            if entry.get() not in SPACES['Rate']:
                pass
            else:
                worksheet[placement[1]] = SPACES['Rate'][entry.get()]
                worksheet[placement[2]] = num_of_days

        else:
            worksheet[placement[1]] = rate_field.get()
            worksheet[placement[2]] = qty_field.get()

        worksheet[placement[0]] = entry.get()

    worksheet[placements['start_date']] = reformat(start)
    worksheet[placements['end_date']] = reformat(end)

    worksheet[placements['notes']] = notes



    # Save Excel sheet
    workbook.save(target)

    # Convert to pdf
    app = client.DispatchEx("Excel.Application")
    app.Interactive = False
    app.Visible = False
    Workbook = app.Workbooks.Open(os.path.abspath(target))

    filename, extension = os.path.abspath(target).split('.')

    try:
        Workbook.ActiveSheet.ExportAsFixedFormat(0, filename + '.pdf')
    except Exception as e:
        print("Failed to convert in PDF format.Please confirm environment meets all the requirements  and try again")
        print(str(e))
    finally:
        Workbook.Close()
        # app.Exit()

    # Delete excel version of invoice
    os.remove(target)

    # Close application
    window.destroy()


ypos += 30
# Generate Button
generate_button = tk.Button(text="Generate Invoice", font=("Arial", 10))
generate_button.place(x=310, y=ypos)
generate_button.bind("<Button-1>", generate)
window.mainloop()


