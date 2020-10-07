# import tkinter as tk
#
#
#
# class ExampleApp(tk.Frame):
#     def __init__(self, master):
#         tk.Frame.__init__(self, master)
#         self.master = master
#         self.some_frame = None
#
#         tk.Button(self.master, text="Create new frame with widgets!", command=self.create_stuff).pack()
#
#     def create_stuff(self):
#         if self.some_frame == None:
#             self.some_frame = tk.Frame(self.master)
#             self.some_frame.pack()
#
#             for i in range(5):
#                 tk.Label(self.some_frame, text="This is label {}!".format(i + 1)).pack()
#
#             tk.Button(self.some_frame, text="Destroy all widgets in this frame!",
#                       command=self.destroy_some_frame).pack()
#
#     def destroy_some_frame(self):
#         self.some_frame.destroy()
#         self.some_frame = None
#
#



import tkinter as tk

# class App(tk.Frame):
#     def __init__(self, parent):
#         super().__init__(parent)
#         self.hourstr=tk.StringVar(self,'10')
#         self.hour = tk.Spinbox(self,from_=0,to=23,wrap=True,textvariable=self.hourstr,width=2,state="readonly")
#         self.minstr=tk.StringVar(self,'30')
#         self.minstr.trace("w",self.trace_var)
#         self.last_value = ""
#         self.min = tk.Spinbox(self,from_=0,to=59,wrap=True,textvariable=self.minstr,width=2,state="readonly")
#         self.hour.grid()
#         self.min.grid(row=0,column=1)
#
#     def trace_var(self,*args):
#         if self.last_value == "59" and self.minstr.get() == "0":
#             self.hourstr.set(int(self.hourstr.get())+1 if self.hourstr.get() !="23" else 0)
#         self.last_value = self.minstr.get()

root = tk.Tk()
w = tk.Spinbox(values=list(range(0,101,5)), wrap=True, width=4)
w.pack()
# App(root).pack()
root.mainloop()


# import tkinter as tk
#
# root = tk.Tk()
#
# on_image = tk.PhotoImage(width=48, height=24)
# off_image = tk.PhotoImage(width=48, height=24)
# on_image.put(("green",), to=(0, 0, 23, 23))
# off_image.put(("0xEDC31B"), to=(24, 0, 47, 23))
#
# var1 = tk.IntVar(value=1)
# var2 = tk.IntVar(value=0)
# cb1 = tk.Checkbutton(root, image=off_image, selectimage=on_image, indicatoron=False,
#                      onvalue=1, offvalue=0, variable=var1)
# cb2 = tk.Checkbutton(root, image=off_image, selectimage=on_image, indicatoron=False,
#                      onvalue=1, offvalue=0, variable=var2)
#
# cb1.pack(padx=20, pady=10)
# cb2.pack(padx=20, pady=10)
#
# root.mainloop()
