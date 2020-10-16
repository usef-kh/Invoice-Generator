import tkinter as tk
from Generator.MediaSpace import MediaSpace
from Generator.MoreSpace import MoreSpace
from tkinter import ttk


class Generator:

    def __init__(self, root, menubar=None):
        self.root = root
        self.root.wm_iconbitmap('Graphics/logo.ico')

        self.modules = ["More Space", "Media Space", "Fabrication Services", "CNC Milling"]

        self.master = tk.LabelFrame(self.root, borderwidth=0, highlightthickness=0)
        self.current_module = None

        self.menubar = menubar or tk.Menu(self.root, tearoff=False)

    def start(self):
        self.new()
        self.create_menubar()
        self.root.mainloop()

    def reset_master(self):
        self.master.destroy()
        self.master = tk.LabelFrame(self.root, borderwidth=0, highlightthickness=0)
        self.master.pack()

    def new(self):
        self.reset_master()
        if self.current_module:
            self.current_module.close()
            self.current_module = None

        self.root.geometry('280x255')
        self.root.title("Invoice Generator")

        tk.Label(self.master, text='Invoice Generator', font=('Arial', 14)).pack(pady=20)

        for name in self.modules:
            button = ttk.Button(self.master, text=name, width=25,
                                command=lambda module=name: self.create(module))
            button.pack(pady=5)

    def create(self, module_name):
        self.master.destroy()

        modules = {'More Space': MoreSpace,
                   'Media Space': MediaSpace,
                   'Fabrication Services': MoreSpace,
                   'CNC Milling': MoreSpace}

        module = modules[module_name]

        my_module = module(self.root)
        self.current_module = my_module
        self.master = self.current_module.master

        my_module.start()

    def create_menubar(self):

        self.root.config(menu=self.menubar)

        file_menubar = tk.Menu(self.menubar, tearoff=False)

        new_menubar = tk.Menu(file_menubar, tearoff=False)
        for name in self.modules:
            new_menubar.add_command(label=name, command=lambda module=name: self.create(module))

        self.menubar.add_cascade(label='File', menu=file_menubar)
        file_menubar.add_cascade(label='New... ', menu=new_menubar)
        file_menubar.add_command(label='Clear', command=self.clear)
        file_menubar.add_command(label='Preview', command=self.preview)
        file_menubar.add_separator()
        file_menubar.add_command(label='Save', command=self.save)

        save_as_menubar = tk.Menu(file_menubar, tearoff=False)
        for option in ['PDF', 'Excel']:
            save_as_menubar.add_command(label=option, command=lambda choice=option: self.save(choice))

        file_menubar.add_cascade(label='Save As...', menu=save_as_menubar)

        file_menubar.add_separator()
        file_menubar.add_command(label='Exit', command=self.close)

        self.menubar.add_command(label='Mode', command=self.change_mode)

    def clear(self):
        if self.current_module:
            self.current_module.clear()

    def save(self, choice='PDF'):
        if self.current_module:
            self.current_module.save(choice)

    def preview(self):
        if self.current_module:
            self.current_module.preview()

    def change_mode(self):
        pass

    def close(self):
        for item in ['File', 'Mode']:
            self.menubar.delete(item)

        if self.current_module:
            self.current_module.close()

        self.master.destroy()


if __name__ == '__main__':
    root = tk.Tk()
    Generator(root)
