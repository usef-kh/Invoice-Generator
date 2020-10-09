import tkinter as tk

from Modules.MediaSpace import MediaSpace
from Modules.MoreSpace import MoreSpace


class Generator:

    def __init__(self):
        self.root = tk.Tk()
        self.root.wm_iconbitmap('Graphics/logo.ico')
        self.modules = ["More Space", "Media Space", "Fabrication Services", "CNC Milling"]

        self.current_frame = None
        self.current_module = None

        self.new()
        self.create_menubar()
        self.root.mainloop()

    def new(self):
        if self.current_frame:
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
        self.menubar.add_command(label='Main Menu', command=self.new)
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
        file_menubar.add_command(label='Exit', command=self.root.destroy)

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

        password_window = tk.Tk()
        password_window.mainloop()
        self.create('More Space')


if __name__ == '__main__':
    Generator()
