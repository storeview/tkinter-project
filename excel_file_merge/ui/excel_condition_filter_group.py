from tkinter import *
from tkinter.ttk import *

class ExcelConditionFilterGroup():
    #初始化
    def __init__(self, parent):
        self.frame = Frame(parent)
        self.num = 0
    
    #直接调用frame的布局方法
    def grid(self, row=0, column=0):
        self.frame.grid(row=row, column=column)

    #添加一个条件筛选
    def add(self, column_titles, selected_columns):
        #删除按钮
        delete_condition_button = Button(self.frame, text="X", width=2, 
                                         class_= str(self.num))
        delete_condition_button.bind("<Button-1>", self.delete_condition_callback)
        
        #列名下拉框
        values = [column_titles[i] for i,v in enumerate(selected_columns) if v != 0 and i != len(column_titles)]
        column_combobox = Combobox(self.frame, 
                          values=values)
        column_combobox.current(0)

        #条件：包含、大于、小于、等于
        condition_combobox = Combobox(self.frame, values=["包含", "大于", "小于", "等于"])
        condition_combobox.current(0)

        #输入框
        val_entry = Entry(self.frame)

        #添加布局
        delete_condition_button.grid(row=self.num, column=0)
        column_combobox.grid(row=self.num, column=1)
        condition_combobox.grid(row=self.num, column=2)
        val_entry.grid(row=self.num, column=3)

        #总数加一
        self.num += 1

    #删除一个条件过滤
    def delete_condition_callback(self, event):
        index = event.widget.config("class")[4]
        name = event.widget._name
        children = self.frame.children
        val = list(children.keys())
        index = val.index(name)
        children[val[index+3]].destroy()
        children[val[index+2]].destroy()
        children[val[index+1]].destroy()
        children[val[index]].destroy()

    #获得当前设置的条件
    def get_conditions(self):
        children = self.frame.children
        conditions = []
        for child in children:
            pop = (str(child)).split("!").pop()
            if len(pop.split("combobox")) > 1:
                conditions.append(children[str(child)].current())
            elif len(pop.split("entry")) > 1:
                conditions.append(children[str(child)].get())
        return conditions