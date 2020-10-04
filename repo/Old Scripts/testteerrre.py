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
window.mainloop()
