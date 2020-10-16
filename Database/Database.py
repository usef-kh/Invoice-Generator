import tkinter as tk
from Database.Sheet import GUISheet
from tkinter import ttk


class DatabaseModule:
    def __init__(self, root):
        self.root = root
        self.sheets = []

    def close(self):
        for sheet in self.sheets:
            sheet.close()


class Customers(DatabaseModule):

    def __init__(self, root):
        self.root = root
        self.module_name = "Customers"

        sheet = GUISheet('customers', root, db='Database.db')
        self.sheets = [sheet]


class MoreSpace(DatabaseModule):

    def __init__(self, root):
        self.root = root
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
        sheet_names = ['morespace_spaces', 'morespace_items', 'morespace_extras']
        self.sheets = []
        for tab, name in zip(tabs.children.values(), sheet_names):
            sheet = GUISheet(name, tab, db='Database.db')
            self.sheets.append(sheet)


class MediaSpace(DatabaseModule):

    def __init__(self, root):
        self.root = root
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
        sheet_names = ['mediaspace_spaces', 'mediaspace_items', 'mediaspace_extras']
        self.sheets = []
        for tab, name in zip(tabs.children.values(), sheet_names):
            sheet = GUISheet(name, tab, db='Database.db')
            self.sheets.append(sheet)


class Database:

    def __init__(self, root):
        self.root = root
        self.root.geometry('630x335')

        self.master = tk.LabelFrame(self.root, borderwidth=0, highlightthickness=0)
        self.master.pack()
        tabs = ttk.Notebook(self.master)
        tabs.pack(expand=True, fill="both")

        module_classes = [Customers, MoreSpace, MediaSpace]
        self.modules = []
        for module in module_classes:
            tab = ttk.Frame(tabs)
            instance = module(tab)

            tabs.add(tab, text=instance.module_name)
            self.modules.append(instance)

    def start(self):
        self.root.mainloop()

    def close(self):
        for module in self.modules:
            module.close()

        self.master.destroy()


if __name__ == '__main__':
    root = tk.Tk()
    Database(root)
