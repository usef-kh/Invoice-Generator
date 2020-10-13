import tkinter as tk
from tkinter import ttk

from Database.Sheet import GUISheet


class DatabaseModule:
    def __init__(self, root):
        self.root = root


class Customers(DatabaseModule):

    def __init__(self, root):
        self.root = root
        self.module_name = "Customers"

        GUISheet('customers', root, db='Database.db')


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
        all_sheets = ['morespace_spaces', 'morespace_items', 'morespace_extras']

        for tab, sheet in zip(tabs.children.values(), all_sheets):
            GUISheet(sheet, tab, db='Database.db')


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
        all_sheets = ['mediaspace_spaces', 'mediaspace_items', 'mediaspace_extras']

        for tab, sheet in zip(tabs.children.values(), all_sheets):
            GUISheet(sheet, tab, db='Database.db')

class Database:

    def __init__(self, root):
        self.root = root
        self.root.geometry('630x335')

        self.master = tk.LabelFrame(self.root, borderwidth=0, highlightthickness=0)
        self.master.pack()
        tabs = ttk.Notebook(self.master)
        tabs.pack(expand=True, fill="both")

        modules = [Customers, MoreSpace, MediaSpace]
        for module in modules:
            tab = ttk.Frame(tabs)
            instance = module(tab)

            tabs.add(tab, text=instance.module_name)

    def start(self):
        self.root.mainloop()


if __name__ == '__main__':
    root = tk.Tk()
    Database(root)