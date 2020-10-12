import sqlite3
import tkinter as tk
from tkinter import ttk, messagebox
from Sheet import GUISheet


class DatabaseModule:
    def __init__(self, root):
        self.root = root


class Customers(DatabaseModule):

    def __init__(self, root):
        super(Customers, self).__init__(root)
        self.module_name = "Customers"

        GUISheet('customers', root)


class MoreSpace(DatabaseModule):

    def __init__(self, root):
        super(MoreSpace, self).__init__(root)
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
            GUISheet(sheet, tab)


class MediaSpace(DatabaseModule):

    def __init__(self, root):
        super(MediaSpace, self).__init__(root)
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
            GUISheet(sheet, tab)


if __name__ == '__main__':
    root = tk.Tk()
    tabs = ttk.Notebook(root)
    tabs.pack(expand=True, fill="both")

    modules = [Customers, MoreSpace, MediaSpace]
    for module in modules:
        tab = ttk.Frame(tabs)
        instance = module(tab)

        tabs.add(tab, text=instance.module_name)

    root.mainloop()
