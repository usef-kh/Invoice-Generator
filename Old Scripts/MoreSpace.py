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


class MoreSpace:

    def __init__(self, master):
        self.master = master
        self.master.geometry('420x670')
        self.master.title("Invoice Generator - More Space")
        self.master.wm_iconbitmap('logo.ico')

        self.load_database()
        self.build_window()

        self.master.mainloop()

    def load_database(self):
        self.database = pd.read_excel("../database.xlsx", sheet_name=None, index_col=None)
        self.SPACES = self.database['MoreSpace Spaces']
        self.SPACES.set_index('Display Name', inplace=True)

        self.ITEMS = self.database['MoreSpace Items']
        self.ITEMS.set_index('Item', inplace=True)

    def build_window(self):

        self.ypos = 20
        self.place_admin_fields()
        self.place_customer_fields()
        self.place_space_reservation_fields()
        self.place_item_reservation_fields()
        self.place_notes()

        self.ypos += 10
        # Generate Button
        self.generate_button = tk.Button(text="Generate Invoice", font=("Arial", 10))
        self.generate_button.place(x=160, y=self.ypos)
        self.generate_button.bind("<Button-1>", self.generate)

        # Clear Button
        self.clear_button = tk.Button(text="Clear", font=("Arial", 10))
        self.clear_button.place(x=25, y=self.ypos)
        self.clear_button.bind("<Button-1>", self.clear)

    def place_admin_fields(self):
        self.create_header('Admin Info')

        self.admin = self.create_entry('Name', default_value='Omar Eddin', change_ypos=False)
        self.ypos += 40

    def place_customer_fields(self):
        self.create_header('Customer Info')

        self.customer = self.create_entry('Name')
        self.company = self.create_entry('Company', default_value='Individual')
        self.credit = self.create_entry('Credit', change_ypos=False)
        self.ypos += 40

    def place_space_reservation_fields(self):
        self.create_header('Space Reservation')

        self.boxes = [(name, tk.IntVar()) for name in self.SPACES.index]
        for i, box in enumerate(self.boxes):
            name, space = box
            checkbox = tk.Checkbutton(self.master, text=name, variable=space)
            checkbox.place(x=40 * i + 50, y=self.ypos)

        self.ypos += 25
        tk.Label(text="Dates", font=("Arial", 10)).place(x=50, y=self.ypos)
        self.start_date = DateEntry(self.master, width=8, background='black', foreground='white', borderwidth=2)
        self.start_date.place(x=120, y=self.ypos)

        self.end_date = DateEntry(self.master, width=8, background='black', foreground='white', borderwidth=2)
        self.end_date.place(x=193, y=self.ypos)
        self.ypos += 25

        self.additional_makers = self.create_entry('Additional Makers', xpos=195, width=11)
        self.tool_exclusions = self.create_entry('Tool Exclusions', xpos=195, width=11)
        self.day_exclusions = self.create_entry('Day Exclusions', xpos=195, width=11)
        self.discount = self.create_spinbox('Discount', nums=list(range(0, 101, 5)), xpos=195)
        self.ypos += 40

    def place_item_reservation_fields(self):
        self.create_header('Item Reservation')

        tk.Label(self.master, text="Item", font=("Arial", 11)).place(x=100, y=self.ypos)
        tk.Label(self.master, text="Rate", font=("Arial", 11)).place(x=238, y=self.ypos)
        tk.Label(self.master, text="Quantity", font=("Arial", 11)).place(x=318, y=self.ypos)
        self.ypos += 20

        self.table_fields = dict()

        def set_rate(event):
            table_entry = self.table_fields[event.widget]
            choice = event.widget.get()

            table_entry[0].delete(0, "end")
            table_entry[0].insert(0, self.ITEMS['Rate'][choice])

            table_entry[1].delete(0, "end")
            table_entry[1].insert(0, 1)

        for i in range(5):
            item_field = ttk.Combobox(self.master, values=list(self.ITEMS.index), width=27)
            item_field.place(x=25, y=self.ypos)

            item_field.bind("<<ComboboxSelected>>", set_rate)

            rate_field = tk.Entry(fg="black", bg="white", width=15)
            rate_field.place(x=208, y=self.ypos)

            qty_field = tk.Entry(fg="black", bg="white", width=15)
            qty_field.place(x=301, y=self.ypos)

            self.table_fields[item_field] = (rate_field, qty_field)

            self.ypos += 18

        self.ypos += 40 - 18

    def place_notes(self):
        self.create_header('Notes')

        self.notes = tk.Text(self.master, width=45, height=2)
        self.notes.place(x=25, y=self.ypos)
        self.notes.insert(1.0, 'Please make your payment through Paypal to info@makeplus.us')
        self.ypos += 40

    def create_header(self, header_text):
        tk.Label(self.master, text=header_text, font=('Arial', 14)).place(x=20, y=self.ypos)
        self.ypos += 30

    def create_entry(self, description_text, xpos=120, default_value=None, width=33, change_ypos=True):
        tk.Label(self.master, text=description_text, font=('Arial', 10)).place(x=50, y=self.ypos)

        entry = tk.Entry(self.master, fg="black", bg="white", width=width)
        entry.place(x=xpos, y=self.ypos)

        if default_value != None:
            entry.insert(0, default_value)

        if change_ypos:
            self.ypos += 21

        return entry

    def create_spinbox(self, description_text, nums=[], xpos=120):
        tk.Label(self.master, text=description_text, font=('Arial', 10)).place(x=50, y=self.ypos)

        spinbox = tk.Spinbox(values=nums, wrap=True, width=4)
        spinbox.place(x=xpos, y=self.ypos)

        return spinbox

    def generate(self, *args):

        def convert_to_num(str, isFloat=True):
            test_str = str
            if isFloat:
                test_str = str.replace('.', '', 1)

            if test_str == '':
                return 0
            elif test_str.isnumeric():
                return float(str)
            else:
                return -1   # Error

        def separate(string_date, split_char='/'):
            string_date = string_date.split(split_char)
            m, d, y = [int(num) for num in string_date]
            return y, m, d

        def count_days(start, end):

            start, end = date(*separate(start)), date(*separate(end))
            delta_day = timedelta(days=1)

            days = {'mon': 0, 'tue': 1, 'wed': 2, 'thu': 3, 'fri': 4, 'sat': 5, 'sun': 6}

            dt = start
            day_count = 0

            while dt <= end:
                if dt.weekday() != days['sun']:
                    day_count += 1
                dt += delta_day

            return day_count

        def reformat(string_date, char='/'):
            date = string_date.split('/')
            for i, entry in enumerate(date):
                if len(entry) == 1:
                    date[i] = '0' + entry

            return char.join(date)

        def generate_col_ids(chars):
            res = []

            for num in range(24, 36):
                entry = []
                for char in chars:
                    entry.append(char + str(num))
                res.append(entry)

            return res

        admin, customer, company = self.admin.get(), self.customer.get(), self.company.get()
        credit = convert_to_num(self.credit.get())

        additional_makers = convert_to_num(self.additional_makers.get(), isFloat=False)
        day_exclusions = convert_to_num(self.day_exclusions.get(), isFloat=False)
        tool_exclusions = convert_to_num(self.tool_exclusions.get(), isFloat=False)
        discount = convert_to_num(self.discount.get()) / 100

        current_file_count = str(len(os.listdir('../../Invoices/MoreSpace/')) + 1)
        zeros = ''.join(['0'] * (5 - len(current_file_count.split())))
        invoice_number = zeros + current_file_count
        invoice_date = datetime.date(datetime.now()).strftime("%m/%d/%y")

        num_of_days = count_days(self.start_date.get(), self.end_date.get())
        num_of_billable_days = num_of_days - day_exclusions

        booked_spaces = []
        for box_name, checkbox in self.boxes[:-1]:  # Only Spaces
            if checkbox.get():
                name = self.SPACES['Space'][box_name]
                unit = self.SPACES['Unit'][box_name]
                rate = float(self.SPACES['Rate'][box_name])

                booked_spaces += [(name, num_of_billable_days, unit, rate)]

        billables = [space for space in booked_spaces]
        num_of_billable_tool_days = 0
        if self.boxes[-1][-1].get():
            box_name = self.boxes[-1][0]
            num_of_billable_tool_days = num_of_billable_days - tool_exclusions

            name = self.SPACES['Space'][box_name]
            unit = self.SPACES['Unit'][box_name]
            rate = float(self.SPACES['Rate'][box_name])

            billables += [(name, num_of_billable_tool_days, unit, rate)]

        if additional_makers > 0:
            unit = 'People'

            if additional_makers == 1:
                unit = 'Person'

            billables += [("Additional Maker(s)", additional_makers, unit, 10)]

        reserved_items = []
        for item, info in self.table_fields.items():
            if item.get() != '':

                _rate, _qty = info

                rate = convert_to_num(_rate.get())
                qty = convert_to_num(_qty.get(), isFloat=False)

                if rate <= 0 or qty <= 0:
                    raise Exception("Invalid rate or quanitity used")

                name = item.get()
                if name in self.ITEMS.index:
                    unit = self.ITEMS['Unit'][name]
                else:
                    unit = 'Item'

                reserved_items += [(name, qty, unit, rate)]

        billables += reserved_items

        # Validate All inputspyin
        try:
            if admin == "":
                raise Exception("Insert Admin Name")

            if customer == "":
                raise Exception("Insert Customer Name")

            if company == "":
                raise Exception("Insert Company Name")

            if credit < 0:
                raise Exception("Credit is a positive number")

            if day_exclusions < 0:
                raise Exception("Day Exclusions is a positive integer")

            if tool_exclusions < 0:
                raise Exception("Tool Exclusions is a positive integer")

            if discount < 0 or discount > 1:
                raise Exception("Discount is a number between 0 and 100")

            if booked_spaces:
                if num_of_days < 1:
                    raise Exception("Start Date is after End Date")

                if num_of_billable_days < 1:
                    raise Exception("Too many day exclusions")

                if additional_makers < 0:
                    raise Exception("Additional Makers is a positive integer")

                elif (additional_makers + 1) > len(booked_spaces) * 4:
                    raise Exception("Only Four Makers Allowed per Space")

            elif num_of_billable_days > 1 or additional_makers > 0:
                raise Exception("No spaces have been booked!")

            if num_of_billable_tool_days < 0:
                raise Exception("Too many tool Exclusions")

            elif not self.boxes[-1][-1].get() and num_of_billable_tool_days > 0:
                raise Exception("Tool exclusion with no tools!")

            for name, qty, unit, rate in reserved_items:
                if qty <= 0:
                    raise Exception("Invalid quantity used in item reservation")
                if rate <= 0:
                    raise Exception("Invalid rate used in item reservation")

        except Exception as e:
            messagebox.showinfo("Error", message=e)
            return

        # Open Excel Sheet
        template = os.path.abspath('../../Templates.xlsx')
        target = '../Invoices/MoreSpace/' + invoice_number + "_" + customer + '.xlsx'

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
            'notes': 'A47'
        }

        workbook = openpyxl.load_workbook(target)
        worksheet = workbook['MoreSpace']

        # add image
        img = openpyxl.drawing.image.Image('logo.png')
        img.width = 85
        img.height = 85
        img.anchor = 'A2'
        worksheet.add_image(img)

        # place info
        worksheet[placements['admin']] = admin
        worksheet[placements['name']] = customer
        worksheet[placements['company']] = company

        if credit > 0:
            worksheet[placements['credit']] = credit
            worksheet.row_dimensions[int(placements['credit'][1:])].hidden = False
        else:
            worksheet.row_dimensions[int(placements['credit'][1:])].hidden = True

        if discount > 0:
            worksheet[placements['discount']] = discount
            worksheet.row_dimensions[int(placements['discount'][1:])].hidden = False
        else:
            worksheet.row_dimensions[int(placements['discount'][1:])].hidden = True

        worksheet[placements['invoice_date']] = invoice_date
        worksheet[placements['invoice_number']] = invoice_number
        worksheet[placements['due_date']] = reformat(self.end_date.get())

        worksheet[placements['start_date']] = reformat(self.start_date.get())
        worksheet[placements['end_date']] = reformat(self.end_date.get())

        worksheet[placements['notes']] = self.notes.get("1.0", 'end-1c')

        for placement, entry in zip(placements['table'], billables):
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

        Workbook.WorkSheets('MoreSpace').Select()
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

    def clear(self, *args):
        self.admin.delete(0, 'end')
        self.admin.insert(0, 'Omar Eddin')

        self.customer.delete(0, 'end')

        self.company.delete(0, 'end')
        self.company.insert(0, 'Individual')

        self.credit.delete(0, 'end')

        for name, space in self.boxes:
            space.set(0)

        self.additional_makers.delete(0, 'end')
        self.tool_exclusions.delete(0, 'end')
        self.day_exclusions.delete(0, 'end')
        self.discount.delete(0, 'end')

        for item_field, info_fields in self.table_fields.items():
            for field in [item_field, *info_fields]:
                field.delete(0, 'end')

        self.notes.delete(1.0, 'end')
        self.notes.insert(1.0, 'Please make your payment through Paypal to info@makeplus.us')


if __name__ == '__main__':
    window = tk.Tk()
    MoreSpace(window)
