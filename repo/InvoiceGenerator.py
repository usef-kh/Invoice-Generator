import tkinter as tk
from MoreSpace import MoreSpace
from MediaSpace import MediaSpace
import os


class InvoiceGenerator:

    def __init__(self):
        self.root = tk.Tk()
        self.root.wm_iconbitmap('logo.ico')
        self.modules = ["More Space", "Media Space", "Fabrication Services", "CNC Milling"]

        self.create_menubar()

        self.current_frame = None
        self.current_module = None

        self.new()
        self.root.mainloop()

    def new(self):
        if self.current_frame:
            print('hi')
            self.current_frame.destroy()
            self.current_module = None

        self.root.geometry('300x275')
        self.root.title("Invoice Generator")

        self.current_frame = tk.LabelFrame(self.root, borderwidth=0, highlightthickness=0)

        tk.Label(self.current_frame, text='Invoice Generator', font=('Arial', 14)).pack(pady=20)

        for name in self.modules:
            button = tk.Button(self.current_frame, text=name, width=20, command=lambda module=name: self.create(module))
            button.pack(pady=5)

        self.current_frame.pack()

    def create(self, module_name):
        self.current_frame.destroy()

        self.current_frame = tk.LabelFrame(self.root)
        modules = {'More Space': MoreSpace,
                   'Media Space': MediaSpace,
                   'Fabrication Services': MoreSpace,
                   'CNC Milling': MoreSpace}

        module = modules[module_name]

        my_module = module(self.root)
        self.current_module = my_module
        self.current_frame = self.current_module.master

        my_module.start()

    def create_menubar(self):

        self.menubar = tk.Menu(self.root, tearoff=False)
        self.root.config(menu=self.menubar)

        file_menubar = tk.Menu(self.menubar, tearoff=False)

        new_menubar = tk.Menu(file_menubar, tearoff=False)
        for name in self.modules:
            new_menubar.add_command(label=name, command=lambda module=name: self.create(module))

        self.menubar.add_cascade(label='File', menu=file_menubar)

        file_menubar.add_cascade(label='New... ', menu=new_menubar)

        file_menubar.add_command(label='Clear', command=self.clear)
        file_menubar.add_command(label='Preview', command=self.preview)

        file_menubar.add_command(label='Generate', command=self.generate)

        file_menubar.add_separator()
        file_menubar.add_command(label='Exit', command=self.root.destroy)

        self.menubar.add_command(label='Mode', command=self.change_mode)

    def clear(self):
        if self.current_module:
            self.current_module.clear()

    def generate(self):
        if self.current_module:
            self.current_module.generate()

    def preview(self):
        if self.current_module:
            self.current_module.preview()

    def change_mode(self):
        pass


if __name__ == '__main__':
    InvoiceGenerator()
