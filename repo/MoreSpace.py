import os
import shutil
import tkinter as tk
from datetime import datetime, date, timedelta
from tkinter import messagebox
from tkinter import ttk

import openpyxl
from Module import Module
from tkcalendar import DateEntry
from win32com import client


class MoreSpace(Module):

    def __init__(self, root):
        super(MoreSpace, self).__init__(root)
        self.root.title("Invoice Generator - More Space")
        self.root.geometry('450x625')

    def load_database(self):
        super(MoreSpace, self).load_database()
        self.SPACES = self.database['MoreSpace Spaces']
        self.SPACES.set_index('Display Name', inplace=True)

        self.ITEMS = self.database['MoreSpace Items']
        self.ITEMS.set_index('Item', inplace=True)

    def build_window(self):
        super(MoreSpace, self).place_admin_fields()
        super(MoreSpace, self).place_customer_fields()
        self.place_space_reservation_fields()
        self.place_item_reservation_fields()
        super(MoreSpace, self).place_notes()
        self.clear()

    def place_space_reservation_fields(self):
        frame = tk.LabelFrame(self.master, text='Space Reservation', font=('Arial', 14), padx=10, pady=10,
                              borderwidth=0, highlightthickness=0)

        spacesframe = tk.LabelFrame(frame, borderwidth=0, highlightthickness=0)
        self.boxes = [(name, tk.IntVar()) for name in self.SPACES.index]
        for i, box in enumerate(self.boxes):
            name, space = box
            checkbox = tk.Checkbutton(spacesframe, text=name, variable=space)
            checkbox.pack(side=tk.LEFT)
        spacesframe.pack(padx=20, anchor=tk.W)

        datesframe = tk.LabelFrame(frame, borderwidth=0, highlightthickness=0)
        tk.Label(datesframe, text='Dates', font=('Arial', 10), padx=10, width=5, anchor=tk.W).grid(row=0, column=0)

        self.start_date = DateEntry(datesframe, width=8, background='black', foreground='white', borderwidth=2)
        self.start_date.grid(row=0, column=1, padx=9)

        self.end_date = DateEntry(datesframe, width=8, background='black', foreground='white', borderwidth=2)
        self.end_date.grid(row=0, column=2, padx=9)
        datesframe.pack(padx=10, pady=5, anchor=tk.W)

        detailsframe = tk.LabelFrame(frame, borderwidth=0, highlightthickness=0)

        entries = []
        for i, text in enumerate(['Tool Exclusions', 'Day Exclusions']):
            tk.Label(detailsframe, text=text, font=('Arial', 10), padx=10, width=12, anchor=tk.W).grid(row=i, column=0,
                                                                                                       padx=10)
            entry = tk.Spinbox(detailsframe, from_=0, to=100, wrap=True, width=4, justify=tk.LEFT)
            entry.grid(row=i, column=1, padx=10, sticky=tk.W)
            entries.append(entry)

        self.tool_exclusions, self.day_exclusions = entries

        tk.Label(detailsframe, text='Additional Makers', font=('Arial', 10), padx=10, width=12, anchor=tk.W).grid(row=0,
                                                                                                                  column=2,
                                                                                                                  padx=10)
        self.additional_makers = tk.Spinbox(detailsframe, from_=0, to=100, wrap=True, width=4, justify=tk.LEFT)
        self.additional_makers.grid(row=0, column=3, padx=10, sticky=tk.W)

        tk.Label(detailsframe, text='Discount', font=('Arial', 10), padx=10, width=12, anchor=tk.W).grid(row=1,
                                                                                                         column=2,
                                                                                                         padx=10)

        self.discount = tk.Spinbox(detailsframe, from_=0, to=101, increment=5, wrap=True, width=4, justify=tk.LEFT)
        self.discount.grid(row=1, column=3, padx=10, sticky=tk.W)
        detailsframe.pack(anchor=tk.W)

        frame.pack(padx=20, pady=2, anchor=tk.W)

    def place_item_reservation_fields(self):
        frame = tk.LabelFrame(self.master, text='Item Reservation', width=10, font=('Arial', 14), padx=10, pady=10,
                              borderwidth=0, highlightthickness=0)
        for i, text in enumerate(['Item', 'Rate', 'Quantity']):
            tk.Label(frame, text=text, font=('Arial', 12), padx=10, width=6).grid(row=0, column=i, padx=10)

        self.table_fields = dict()

        def set_rate(event):
            table_entry = self.table_fields[event.widget]
            choice = event.widget.get()

            table_entry[0].delete(0, "end")
            table_entry[0].insert(0, self.ITEMS['Rate'][choice])

            table_entry[1].delete(0, "end")
            table_entry[1].insert(0, 1)

        for i in range(5):
            item_field = ttk.Combobox(frame, values=list(self.ITEMS.index), font=('Arial', 9), width=25)
            # item_field = tk.Entry(frame, fg="black", bg="white")
            item_field.grid(row=i + 1, column=0, sticky="NSEW")

            item_field.bind("<<ComboboxSelected>>", set_rate)

            rate_field = tk.Entry(frame, fg="black", bg="white", font=('Arial', 11), width=3)
            rate_field.grid(row=i + 1, column=1, sticky="NSEW")

            qty_field = tk.Entry(frame, fg="black", bg="white", font=('Arial', 11), width=4)
            qty_field.grid(row=i + 1, column=2, sticky="NSEW")
            self.table_fields[item_field] = (rate_field, qty_field)

        frame.pack(padx=20, pady=2, anchor=tk.W)

    def place_notes(self):
        frame = tk.LabelFrame(self.master, text='Notes', font=('Arial', 14), padx=10, pady=10, borderwidth=0,
                              highlightthickness=0)
        self.notes = tk.Text(frame, width=49, height=2)
        self.notes.pack()

        frame.pack(padx=20, pady=2, anchor=tk.W)

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
                return -1  # Error

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

        current_file_count = str(len(os.listdir('../Invoices/MoreSpace/')) + 1)
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
        template = os.path.abspath('../Templates.xlsx')
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
        self.target_xlsx = target
        self.target_pdf = filename + '.pdf'
        os.remove(self.target_xlsx)

        os.startfile(self.target_pdf)

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
    module = MoreSpace(window)
    module.start()