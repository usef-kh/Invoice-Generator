import os
import shutil
import tkinter as tk
from datetime import datetime
from tkinter import messagebox
from tkinter import ttk

import openpyxl
from Module import Module
from tkcalendar import DateEntry
from win32com import client


class MediaSpace(Module):
    def __init__(self, root):
        super(MediaSpace, self).__init__(root)
        self.root.title("Invoice Generator - Media Space")
        self.root.geometry('465x650')

    def load_database(self):
        super(MediaSpace, self).load_database()
        self.SPACES = self.database['Media Spaces']
        self.SPACES.set_index('Space', inplace=True)

        self.ITEMS = self.database['Media Items']
        self.ITEMS.set_index('Item', inplace=True)

    def build_window(self):
        super(MediaSpace, self).place_admin_fields()
        super(MediaSpace, self).place_customer_fields()
        self.place_space_reservation_details()
        super(MediaSpace, self).place_item_reservation_fields()
        super(MediaSpace, self).place_notes()
        self.clear()

    def place_space_reservation_details(self):
        frame = tk.LabelFrame(self.master, text='Reservation Details', font=('Arial', 14), padx=10, pady=10,
                              borderwidth=0, highlightthickness=0)

        for i, text in enumerate(['Selection', 'Day', 'Timings']):
            tk.Label(frame, text=text, font=('Arial', 12), padx=10, width=6).grid(row=0, column=i, padx=10)

        options = []
        for i in range(12):
            str_num = str(i)

            if len(str_num) == 1:
                str_num = '0' + str_num

            options.append(str_num + ':' + '00 am')
            options.append(str_num + ':' + '30 am')

        options.append('12' + ':' + '00 pm')
        options.append('12' + ':' + '30 pm')

        for i in range(1, 12):
            str_num = str(i)

            if len(str_num) == 1:
                str_num = '0' + str_num

            options.append(str_num + ':' + '00 pm')
            options.append(str_num + ':' + '30 pm')

        self.reservation_details = dict()
        for i in range(3):
            item_field = ttk.Combobox(frame, values=list(self.SPACES.index), width=25)
            item_field.grid(row=i + 1, column=0)

            date_field = DateEntry(frame, width=8, background='black', foreground='white', borderwidth=2)
            date_field.grid(row=i + 1, column=1)

            timing_frame = tk.LabelFrame(frame, borderwidth=0, highlightthickness=0)

            start_time = ttk.Combobox(timing_frame, values=options, width=8)
            end_time = ttk.Combobox(timing_frame, values=options, width=8)

            start_time.grid(row=0, column=1)
            end_time.grid(row=0, column=2)
            timing_frame.grid(row=i + 1, column=2)
            timings = (start_time, end_time)
            self.reservation_details[item_field] = (date_field, timings)

        discount_frame = tk.LabelFrame(frame, borderwidth=0, highlightthickness=0)
        tk.Label(discount_frame, text="Discount", font=('Arial', 10), pady=10, width=6).pack(side=tk.LEFT)
        self.discount = tk.Spinbox(discount_frame, from_=0, to=101, increment=5, wrap=True, width=3)
        self.discount.pack(side=tk.RIGHT, padx=20)

        discount_frame.grid(row=i + 2, column=0, sticky=tk.W)

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
                return -1

        def reformat(string_date, char='/'):
            date = string_date.split('/')
            for i, entry in enumerate(date):
                if len(entry) == 1:
                    date[i] = '0' + entry

            return char.join(date)

        def generate_col_ids(chars):
            res = []

            for num in range(25, 36):
                entry = []
                for char in chars:
                    entry.append(char + str(num))
                res.append(entry)

            return res

        admin, customer, company = self.admin.get(), self.customer.get(), self.company.get()
        credit = convert_to_num(self.credit.get())

        discount = convert_to_num(self.discount.get()) / 100

        current_file_count = str(len(os.listdir('../Invoices/MediaSpace/')) + 1)
        zeros = ''.join(['0'] * (5 - len(current_file_count.split())))
        invoice_number = zeros + current_file_count

        invoice_date = datetime.date(datetime.now()).strftime("%m/%d/%y")

        billables, reservation_descriptions = [], []
        for space, info in self.reservation_details.items():
            if space.get():
                name = space.get()
                _date, timings = info
                date = reformat(_date.get())

                start, end = [timing.get() for timing in timings]

                start_time, start_meridiem = start.split(' ')
                end_time, end_meridiem = end.split(' ')

                start_hr, start_min = [int(num) for num in start_time.split(':')]
                end_hr, end_min = [int(num) for num in end_time.split(':')]

                if start_meridiem == 'pm':
                    start_hr += 12

                if end_meridiem == 'pm':
                    end_hr += 12

                duration = end_hr - start_hr + end_min / 60 - start_min / 60

                time = start + ' - ' + end

                rate = float(self.SPACES['Rate'][name])
                unit = self.SPACES['Unit'][name]

                billables += [(name, duration, unit, rate)]
                reservation_descriptions += [(date, time)]

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

        try:
            if admin == "":
                raise Exception("Insert Admin Name")

            if customer == "":
                raise Exception("Insert Customer Name")

            if company == "":
                raise Exception("Insert Company Name")

            if credit < 0:
                raise Exception("Credit is a positive number")

            if discount < 0 or discount > 1:
                raise Exception("Discount is a number between 0 and 100")

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
        target = '../Invoices/MediaSpace/' + invoice_number + "_" + customer + '.xlsx'

        # Create duplicate of template
        shutil.copyfile(template, target)

        placements = {
            'admin': 'A6',
            'name': 'A13',
            'company': 'A14',
            'invoice_date': 'G12',
            'invoice_number': 'G11',
            'due_date': 'G13',
            'reservation_details': [['A18', 'B18'], ['A19', 'B19'], ['A20', 'B20']],
            'table': (generate_col_ids(['A', 'C', 'D', 'F'])),
            'discount': 'F40',
            'credit': 'G43',
            'notes': 'A48'
        }

        workbook = openpyxl.load_workbook(target)
        worksheet = workbook['MediaSpace']

        # add image
        img = openpyxl.drawing.image.Image('Graphics/logo.png')
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
        worksheet[placements['due_date']] = reservation_descriptions[0][0]

        worksheet[placements['notes']] = self.notes.get("1.0", 'end-1c')

        for i, entry in enumerate(zip(placements['reservation_details'], reservation_descriptions)):
            placement, description = entry
            for cell, info in zip(placement, description):
                worksheet[cell] = info

        for placement in placements['reservation_details'][i + 1:]:
            row = placement[0][1:]
            worksheet.row_dimensions[int(row)].hidden = True

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
        Workbook.WorkSheets('MediaSpace').Select()
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

        for selection, info_fields in self.reservation_details.items():
            # day, start, end =
            for field in [selection, info_fields[0], *info_fields[1]]:
                field.delete(0, 'end')

        # self.discount.delete(0, 'end')

        for item_field, info_fields in self.table_fields.items():
            for field in [item_field, *info_fields]:
                field.delete(0, 'end')

        self.notes.delete(1.0, 'end')
        self.notes.insert(1.0, 'Please make your payment through Paypal to info@makeplus.us')


if __name__ == '__main__':
    window = tk.Tk()
    module = MediaSpace(window)
    module.start()
