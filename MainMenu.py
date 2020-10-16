import tkinter as tk
from Database.Database import Database
from Generator.Generator import Generator
from tkinter import ttk


class MainMenu:

    def __init__(self, root):
        self.root = root

        self.root.title("MakePlus Automation")
        self.root.wm_iconbitmap('Graphics/logo.ico')

        self.modules = ["Invoice Generator", "Inventory", "Analytics"]

        self.master = tk.LabelFrame(self.root, borderwidth=0, highlightthickness=0)
        self.current_module = None

    def start(self):
        self.new()
        self.create_menubar()
        self.root.geometry('280x275')
        self.root.mainloop()

    def new(self):
        self.reset()

        self.root.geometry('280x255')
        self.root.title("MakePlus Automation")

        tk.Label(self.master, text='MakePlus Automation', font=('Arial', 14)).pack(pady=20)

        for name in self.modules:
            button = ttk.Button(self.master, text=name, width=25,
                                command=lambda module=name: self.create(module))
            button.pack(pady=5)

        self.master.pack()

    def create(self, module_name):
        self.reset()

        modules = {'Invoice Generator': Generator,
                   'Inventory': Database,
                   'Analytics': Database}

        module = modules[module_name]

        my_module = module(root, self.menubar)
        self.current_module = my_module
        self.master = self.current_module.master

        my_module.start()

    def create_menubar(self):

        self.menubar = tk.Menu(self.root, tearoff=False)
        self.root.config(menu=self.menubar)

        menu_menubar = tk.Menu(self.menubar, tearoff=False)

        for name in self.modules:
            menu_menubar.add_command(label=name, command=lambda module=name: self.create(module))

        menu_menubar.add_separator()
        menu_menubar.add_command(label='Main Menu', command=self.new)

        self.menubar.add_cascade(label='Menu', menu=menu_menubar)

    def reset(self):
        self.master.destroy()
        self.master = tk.LabelFrame(self.root, borderwidth=0, highlightthickness=0)

        if self.current_module:
            self.current_module.close()
            self.current_module = None


if __name__ == '__main__':
    root = tk.Tk()
    mainmenu = MainMenu(root)
    mainmenu.start()
