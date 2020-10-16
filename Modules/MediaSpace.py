import os
import tkinter as tk
from datetime import datetime, date
from tkinter import ttk

from tkcalendar import DateEntry

from Database.Sheet import Sheet
from Modules.Module import Module


class MediaSpace(Module):
    def __init__(self, root):
        super(MediaSpace, self).__init__(root)
        self.root.title("Invoice Generator - Media Space")
        self.module_name = 'MediaSpace'
        self.root.geometry('460x610')

    def load_database(self):
        self.SPACES = Sheet('mediaspace_spaces')
        self.ITEMS = Sheet('mediaspace_items')
        self.EXTRAS = Sheet('mediaspace_extras')

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
            item_field = ttk.Combobox(frame, values=self.SPACES.list_items(), width=22)
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

    def read_inputs(self, *args):

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

        def separate(string_date, split_char='/'):
            string_date = string_date.split(split_char)
            m, d, y = [int(num) for num in string_date]
            return y, m, d

        admin, customer, company = self.admin.get(), self.customer.get(), self.company.get()
        credit = convert_to_num(self.credit.get())

        discount = convert_to_num(self.discount.get()) / 100

        current_file_count = str(len(os.listdir('Invoices/MediaSpace/')) + 1)
        zeros = ''.join(['0'] * (5 - len(current_file_count.split())))
        invoice_number = zeros + current_file_count

        invoice_date = datetime.date(datetime.now()).strftime("%m/%d/%y")
        days = {'mon': 0, 'tue': 1, 'wed': 2, 'thu': 3, 'fri': 4, 'sat': 5, 'sun': 6}

        billables, reservation_descriptions = [], []
        for space, info in self.reservation_details.items():
            if space.get():
                name = space.get()
                _date, timings = info

                start, end = [timing.get() for timing in timings]

                # seperate time from am/pm
                start_time, start_meridiem = start.split(' ')
                end_time, end_meridiem = end.split(' ')

                start_hr, start_min = [int(num) for num in start_time.split(':')]
                end_hr, end_min = [int(num) for num in end_time.split(':')]

                # reformat timing to 24-hr format for calculations
                if start_meridiem == 'pm' and start_hr != 12:
                    start_hr += 12

                if end_meridiem == 'pm' and end_hr != 12:
                    end_hr += 12

                duration = end_hr - start_hr + end_min / 60 - start_min / 60
                time = start + ' - ' + end

                # Extract quantities
                _, weekday_rate, weekend_rate, unit = self.SPACES[name]

                # Choose rate based on weekday
                display_name = name
                dt = date(*separate(_date.get()))
                if dt.weekday() == days['sun'] or dt.weekday() == days['sat']:
                    rate = weekend_rate
                    display_name = name + ' - Weekend'
                else:
                    rate = weekday_rate

                booking_date = reformat(_date.get())

                # Save info
                billables += [(display_name, duration, unit, float(rate))]
                reservation_descriptions += [(booking_date, time)]

        reserved_items = []
        for item, info in self.table_fields.items():
            if item.get() != '':

                _rate, _qty = info

                rate = convert_to_num(_rate.get())
                qty = convert_to_num(_qty.get(), isFloat=False)

                if rate <= 0 or qty <= 0:
                    raise Exception("Invalid rate or quanitity used")

                item_name = item.get()
                if item_name in self.ITEMS:
                    _, _, unit = self.ITEMS[item_name]
                else:
                    unit = "Item"

                reserved_items += [(item_name, qty, unit, rate)]

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
            return e

        target = 'Invoices/MediaSpace/' + invoice_number + "_" + customer

        cells = {
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

        # place info
        entries = [(cells['admin'], admin),
                   (cells['name'], customer),
                   (cells['company'], company),
                   (cells['invoice_date'], invoice_date),
                   (cells['invoice_number'], invoice_number),
                   (cells['due_date'], reservation_descriptions[0][0]),
                   (cells['notes'], self.notes.get("1.0", 'end-1c'))]

        for placement, entry in zip(cells['table'], billables):
            for cell, info in zip(placement, entry):
                entries.append((cell, info))

        for i, entry in enumerate(zip(cells['reservation_details'], reservation_descriptions)):
            placement, description = entry
            for cell, info in zip(placement, description):
                entries.append((cell, info))

        hidden_entries = []
        for cell1, cell2 in cells['reservation_details'][i + 1:]:
            hidden_entries.append((cell1, 0))

        hidden_entries.append((cells['credit'], credit))
        hidden_entries.append((cells['discount'], discount))

        return target, entries, hidden_entries

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
