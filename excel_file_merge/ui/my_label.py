from tkinter import *
from tkinter.ttk import *


class MyLabel():
    def __init__(self, parent, text=" ", width=0, anchor="w"):
        self.text_var = StringVar()
        self.text_var.set(text)
        print(text)
        print(self.text_var.get())


        if width > 0:
            self.label = Label(parent, textvariable=self.text_var, width=width, anchor=anchor)
        else:
            self.label = Label(parent, textvariable=self.text_var, anchor=anchor)
        
    #直接调用label的布局方法
    def grid(self, row=0, column=0):
        self.label.grid(row=row, column=column)

    def text(self, text):
        self.text_var.set(text)