import os
import shutil
import tkinter as tk
from datetime import datetime, date, timedelta
from tkinter import messagebox
from tkinter import ttk

import openpyxl
import pandas as pd
from tkcalendar import DateEntry
from win32com import client

window = tk.Tk()
window.geometry('420x670')
window.title("More Space")
window.wm_iconbitmap('logo.ico')

database = pd.read_excel("../database.xlsx", sheet_name=None, index_col=None)
SPACES = database['Spaces']

SPACES.set_index('Display Name', inplace=True)

ITEMS = database['Items']
ITEMS.set_index('Item', inplace=True)

# initial coordinates and some default spacing
ypos = 20
section_spacing = 40
header_spacing = 30
spacing = 20

# Admin Info
tk.Label(text='Admin Info', font=('Arial', 14)).place(x=20, y=ypos)
ypos += header_spacing

tk.Label(text='Name', font=('Arial', 10)).place(x=50, y=ypos)

admin = tk.Entry(fg="black", bg="white", width=33)
admin.place(x=120, y=ypos)
admin.insert(0, 'Omar Eddin')

# Customer Info
ypos += section_spacing
tk.Label(text='Customer Info', font=('Arial', 14)).place(x=20, y=ypos)
ypos += header_spacing

tk.Label(text='Name', font=('Arial', 10)).place(x=50, y=ypos)
customer = tk.Entry(fg="black", bg="white", width=33)
customer.place(x=120, y=ypos)

ypos += spacing + 1
tk.Label(text='Company', font=('Arial', 10)).place(x=50, y=ypos)
company = tk.Entry(fg="black", bg="white", width=33)
company.place(x=120, y=ypos)
company.insert(0, 'Individual')

ypos += spacing + 1
tk.Label(text='Credit', font=('Arial', 10)).place(x=50, y=ypos)
credit = tk.Entry(fg="black", bg="white", width=33)
credit.place(x=120, y=ypos)

# Space Reservation
ypos += section_spacing
tk.Label(text='Space Reservation', font=('Arial', 14)).place(x=20, y=ypos)
ypos += header_spacing

names = SPACES.index
boxes = [(name, tk.IntVar()) for name in names]

for i, box in enumerate(boxes):
    name, space = box
    checkbox = tk.Checkbutton(window, text=name, variable=space)
    checkbox.place(x=40 * i + 50, y=ypos)

ypos += 25
tk.Label(text="Dates", font=("Arial", 10)).place(x=50, y=ypos)
start_date = DateEntry(window, width=8, background='black', foreground='white', borderwidth=2)
start_date.place(x=120, y=ypos)

end_date = DateEntry(window, width=8, background='black', foreground='white', borderwidth=2)
end_date.place(x=193, y=ypos)

ypos += spacing + 5
tk.Label(text='Additional Makers', font=('Arial', 10)).place(x=50, y=ypos)
additional_makers = tk.Entry(fg="black", bg="white", width=11)
additional_makers.place(x=195, y=ypos)

ypos += spacing + 1
tk.Label(text='Tool Exclusions', font=('Arial', 10)).place(x=50, y=ypos)
tool_exclusions = tk.Entry(fg="black", bg="white", width=11)
tool_exclusions.place(x=195, y=ypos)

ypos += spacing + 1
tk.Label(text='Day Exclusions', font=('Arial', 10)).place(x=50, y=ypos)
day_exclusions = tk.Entry(fg="black", bg="white", width=11)
day_exclusions.place(x=195, y=ypos)

ypos += spacing + 1
tk.Label(text='Discount', font=('Arial', 10)).place(x=50, y=ypos)
discount = tk.Entry(fg="black", bg="white", width=11)
discount.place(x=195, y=ypos)

# Item Reservation
ypos += section_spacing
tk.Label(text='Item Reservation', font=('Arial', 14)).place(x=20, y=ypos)
ypos += header_spacing

tk.Label(text="Item", font=("Arial", 11)).place(x=100, y=ypos)
tk.Label(text="Rate", font=("Arial", 11)).place(x=238, y=ypos)
tk.Label(text="Quantity", font=("Arial", 11)).place(x=318, y=ypos)

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
    item_field.place(x=25, y=ypos)

    item_field.bind("<<ComboboxSelected>>", set_rate)

    rate_field = tk.Entry(fg="black", bg="white", width=15)
    rate_field.place(x=208, y=ypos)

    qty_field = tk.Entry(fg="black", bg="white", width=15)
    qty_field.place(x=301, y=ypos)

    table_fields[item_field] = (rate_field, qty_field)

    ypos += 18

# Notes
ypos += section_spacing - 18
tk.Label(text='Notes', font=('Arial', 14)).place(x=20, y=ypos)
ypos += header_spacing

notes = tk.Text(window, width=45, height=2)
notes.place(x=25, y=ypos)
notes.insert(1.0, 'Please make your payment through Paypal to info@makeplus.us')


# Generate Invoice
def generate(*args):
    target = False

    def separate(string_date, split_char='/'):
        string_date = string_date.split(split_char)
        m, d, y = [int(num) for num in string_date]
        return y, m, d

    def reformat(string_date, char='/'):
        date = string_date.split('/')
        for i, entry in enumerate(date):
            if len(entry) == 1:
                date[i] = '0' + entry

        return char.join(date)

    def count_days(start, end):
        delta_day = timedelta(days=1)

        days = {'mon': 0, 'tue': 1, 'wed': 2, 'thu': 3, 'fri': 4, 'sat': 5, 'sun': 6}

        dt = start
        day_count = 0

        while dt <= end:
            if dt.weekday() != days['sun']:
                day_count += 1
            dt += delta_day

        return day_count

    def reformat_number(num):

        num = str(num)

        while len(num) < 5:
            num = '0' + num
        return num

    def generate_col_ids(chars):
        res = []

        for num in range(24, 36):
            entry = []
            for char in chars:
                entry.append(char + str(num))
            res.append(entry)

        return res

    def convert_to_num(str, exception_msg, isFloat = True):

        test_str = str
        if isFloat:
            test_str = str.replace('.', '', 1)

        if test_str == '':
            return 0
        elif test_str.isnumeric():
            return float(str)
        else:
            raise Exception(exception_msg)

    try:
        # Read All inputs and validate reservation
        day_exclusion_count = convert_to_num(day_exclusions.get(), "Day Exclusions is a positive integer", isFloat=False)
        tool_exclusion_count = convert_to_num(tool_exclusions.get(), "Tool Exclusions is a positive integer", isFloat=False)

        invoice_number = reformat_number(len(os.listdir('../Invoices')) + 1)
        invoice_date = datetime.date(datetime.now()).strftime("%m/%d/%y")

        admin_name = admin.get()

        if admin_name == "":
            raise Exception("Insert Admin Name")

        customer_name = customer.get()

        if customer_name == "":
            raise Exception("Insert Customer Name")

        company_name = company.get()

        if company_name == "":
            raise Exception("Insert Company Name")

        credit_amount = convert_to_num(credit.get(), "Credit is a positive number")
        discount_amount = convert_to_num(discount.get(), "Discount is a positive number") / 100

        if discount_amount > 1:
            raise Exception("Discount is more than 100%")

        additional_makers_count = convert_to_num(additional_makers.get(), "Additional Makers is a positive integer", isFloat=False)

        num_of_days = count_days(date(*separate(start_date.get())), date(*separate(end_date.get())))

        if num_of_days < 1:
            raise Exception("Start Date is After End Date")

        num_of_days -= day_exclusion_count

        if num_of_days < 1:
            raise Exception("Too many Day exclusions")

        # Open Excel Sheet
        template = os.path.abspath('../Template.xlsx')
        # target = '../Invoices/' + datetime.date(datetime.now()).strftime("%y%m%d") + "_" + customer.get() + '.xlsx'
        target = '../Invoices/' + invoice_number + "_" + customer.get() + '.xlsx'


        # Create duplicate of template
        shutil.copyfile(template, target)

        placements = {
            'admin': 'A6',
            'name': 'A13',
            'company': 'A14',
            'invoice_date': 'G12',
            'invoice_number': 'G11',
            'due_date': 'G13',
            'start_date': 'B17',
            'end_date': 'B18',
            'table': (generate_col_ids(['A', 'C', 'D', 'F'])),
            'discount': 'F39',
            'credit': 'G42',
            'notes': 'A47'}

        workbook = openpyxl.load_workbook(target)
        worksheet = workbook['Sheet1']

        # add image
        img = openpyxl.drawing.image.Image('logo.png')
        img.width = 85
        img.height = 85
        img.anchor = 'A2'
        worksheet.add_image(img)

        # place info
        worksheet[placements['admin']] = admin_name
        worksheet[placements['name']] = customer_name
        worksheet[placements['company']] = company_name

        if credit_amount > 0:
            worksheet[placements['credit']] = credit_amount
            worksheet.row_dimensions[int(placements['credit'][1:])].hidden = False
        else:
            worksheet.row_dimensions[int(placements['credit'][1:])].hidden = True

        if discount_amount > 0:
            worksheet[placements['discount']] = discount_amount
            worksheet.row_dimensions[int(placements['discount'][1:])].hidden = False
        else:
            worksheet.row_dimensions[int(placements['discount'][1:])].hidden = True

        worksheet[placements['invoice_date']] = invoice_date
        worksheet[placements['invoice_number']] = invoice_number
        worksheet[placements['due_date']] = reformat(end_date.get())

        worksheet[placements['start_date']] = reformat(start_date.get())
        worksheet[placements['end_date']] = reformat(end_date.get())

        worksheet[placements['notes']] = notes.get("1.0", 'end-1c')

        table = []
        for name, checkbox in boxes[:-1]:  # Spaces
            if checkbox.get():
                table.append(
                    (SPACES['Space'][name], num_of_days, SPACES['Unit'][name], float(SPACES['Rate'][name]))
                )

        if len(table) == 0:
            raise Exception("No Spaces have been booked")

        if (additional_makers_count + 1) > len(table) * 4:
            raise Exception("Only Four Makers Allowed per Space")

        if boxes[-1][-1].get():  # Tools
            name = boxes[-1][0]

            num_of_tool_days = num_of_days - tool_exclusion_count
            if num_of_tool_days < 1:
                raise Exception("Too many tool exclusions")

            table.append(
                (SPACES['Space'][name], num_of_tool_days, SPACES['Unit'][name],
                 float(SPACES['Rate'][name]))
            )

        elif tool_exclusion_count:
            raise Exception("Tool exclusions with no tools!")

        if additional_makers_count > 0:
            unit = 'People'

            if additional_makers_count == 1:
                unit = 'Person'

            table.append(("Additional Maker(s)", additional_makers_count, unit, 10))

        for item, info in table_fields.items():
            if item.get() != '':

                _rate, _qty = info

                rate = convert_to_num(_rate.get(), "Invalid rate used")
                qty = convert_to_num(_qty.get(), "Invalid quantity Used", isFloat=False)

                if rate <= 0 or qty <= 0:
                    raise Exception("Rate and Quantity cant be zero")

                name = item.get()
                if name in ITEMS.index:
                    unit = ITEMS['Unit'][name]
                else:
                    unit = 'Item'

                table.append((name, qty, unit, rate))

        for placement, entry in zip(placements['table'], table):
            for cell, info in zip(placement, entry):
                worksheet[cell] = info

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
            print(
                "Failed to convert in PDF format.Please confirm environment meets all the requirements  and try again")
            print(str(e))
        finally:
            Workbook.Close()
            # app.Exit()

        # Delete excel version of invoice
        os.remove(target)
        os.startfile(filename + '.pdf')

        # Close application
        # window.destroy()

    except Exception as e:

        messagebox.showinfo("Error", message=e)

        # delete what has been created so far
        if target:
            os.remove(target)


def clear(*args):
    # Clear All Fields
    admin.delete(0, 'end')
    admin.insert(0, 'Omar Eddin')

    customer.delete(0, 'end')

    company.delete(0, 'end')
    company.insert(0, 'Individual')

    credit.delete(0, 'end')

    for name, space in boxes:
        space.set(0)

    additional_makers.delete(0, 'end')
    tool_exclusions.delete(0, 'end')
    day_exclusions.delete(0, 'end')
    discount.delete(0, 'end')

    for item_field, info_fields in table_fields.items():
        for field in [item_field, *info_fields]:
            field.delete(0, 'end')

    notes.delete(1.0, 'end')
    notes.insert(1.0, 'Please make your payment through Paypal to info@makeplus.us')

ypos += section_spacing + 10
# Generate Button
generate_button = tk.Button(text="Generate Invoice", font=("Arial", 10))
generate_button.place(x=160, y=ypos)
generate_button.bind("<Button-1>", generate)

# Clear Button
clear_button = tk.Button(text="Clear", font=("Arial", 10))
clear_button.place(x=25, y=ypos)
clear_button.bind("<Button-1>", clear)
window.mainloop()
