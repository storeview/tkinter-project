from tkinter import *
from tkinter.ttk import *


#configs
font = ("Arial Bold", 14)


root = Tk()
root.title("这是标题")
root.geometry("300x400")


lbl = Label(root, text="Label1")

lbl.grid(column=0, row=0)

btn = Button(root, text="Click Me")
btn.grid(column=1, row=0)

button = Button(root, text='Okay')

def myaction():
    print(1)

action = Button(root, text="Action", default="active", command=myaction)
root.bind('<Return>', lambda e: action.invoke())
action.grid(column=2, row=0)


countryvar = StringVar() 
var_list = ["111", "222", "333"]
country = Combobox(root, textvariable=countryvar, values=var_list)
country
country.grid(column=3, row=0)

root.mainloop()