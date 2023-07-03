from tkinter import *
from tkinter.ttk import *
from tkinter import Checkbutton

class ExcelColumnExportGroup():
    #初始化
    def __init__(self, parent):
        self.frame = Frame(parent)
        self.g_selected_columns = []
    
    #直接调用frame的布局方法
    def grid(self, row=0, column=0):
        self.frame.grid(row=row, column=column)

    #添加一个条件筛选
    def update(self, column_titles, selected_columns):
        if len(self.g_selected_columns) == 0:
            self.g_selected_columns = [IntVar() for i in range(len(column_titles)+1)]
        children = self.frame.children
        val = list(children.keys())
        for v in val:
            children[v].destroy()
        #全选/全不选
        select_all_checkbutton_status = IntVar()
        select_all_checkbutton_status.set(selected_columns[len(column_titles)])
        self.g_selected_columns[len(column_titles)] = select_all_checkbutton_status
        select_all_checkbutton = Checkbutton(self.frame, text="全选", width=15, anchor="w", 
                                    onvalue=1, offvalue=0, variable=self.g_selected_columns[len(column_titles)],
                                    command=lambda: self.select_all(column_titles, selected_columns))
        select_all_checkbutton.grid(row=0, column=0)

        for index, column_title in enumerate(column_titles):
            column_checkbox = Checkbutton(self.frame, text=column_title, width=15, anchor="w", 
                                            onvalue=1, offvalue=0, variable=self.g_selected_columns[index],
                                            command=lambda: self.state_changed(selected_columns))
            if selected_columns[index] == 1:
                self.g_selected_columns[index].set(1)
            else:
                self.g_selected_columns[index].set(0)
            index += 1
            column_checkbox.grid(row=int(int(index)/4), column=int(int(index)%4))


    #更新选择的字段
    def state_changed(self, selected_columns):
        print(selected_columns)
        for i in range(len(selected_columns)):
            checkbutton_state = self.g_selected_columns[i].get()
            selected_columns[i] = checkbutton_state

    #全选字段
    def select_all(self, column_titles, selected_columns):
        select_all_checkbutton_status = self.g_selected_columns[len(self.g_selected_columns)-1]
        #取消全选
        if select_all_checkbutton_status.get() == 0:
            for i in range(len(selected_columns)):
                selected_columns[i] = 0
            selected_columns[len(selected_columns)-1] = 0
        else:
            for i in range(len(selected_columns)):
                selected_columns[i] = 1
            selected_columns[len(selected_columns)-1] = 1
        self.update(column_titles, selected_columns)