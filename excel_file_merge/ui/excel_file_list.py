from tkinter import *
from tkinter.ttk import *

class ExcelFileListFrame():
    """Excel文件列表Frame"""
    def __init__(self, parent=None):
        self.frame = Frame(parent)
        self.delete_buttons = []
    
    def grid(self, row, column):
        self.frame.grid(row=row, column=column)

    def update(self, files):
        for index, file in enumerate(files):
            delete_button = Button(self.frame, text="X", width=2, class_= str(index))
            self.delete_buttons.append(delete_button)
            
            split_len = len(file["filename"].split("/"))
            filenames = Label(self.frame, text=file["filename"].split("/")[split_len-1], width=40, anchor="w")
            sheetnames_combobox = Combobox(self.frame, values=file["sheetnames"])
            sheetnames_combobox.current(file["selected"])
            file_detail = Label(self.frame, text=f"数据行：{file['rows']}，数据列：{file['columns']}")
            delete_button.grid(row=index, column=0)
            filenames.grid(row=index, column=1)
            sheetnames_combobox.grid(row=index, column=2)
            file_detail.grid(row=index, column=3)
            sheetnames_combobox.bind("<<ComboboxSelected>>", self.selection_changed)

    def destroy(self, index):
        children = self.frame.children
        val = list(children.keys())
        for v in val:
            children[v].destroy()
        self.delete_buttons = []

    def selection_changed(self, event):
        print("selection changed...")


    
    def get_delete_buttons(self):
        return self.delete_buttons

            



if __name__ == "__main__":
    excel_file_list = ExcelFileListFrame()
    files = [{'filename': 'D:/Windows数据移动到此文件夹/Desktop/测试用例导出.xlsx', 'sheetnames': ['Sheet1', 'Sheet2', 'Sheet3'], 'selected': 0, 'rows': 40, 'columns': 26}, {'filename': 'D:/Windows数据移动到此文件夹/Desktop/测试用例导出 - 副本.xlsx', 'sheetnames': ['Sheet1', '标签页2'], 'selected': 0, 'rows': 40, 'columns': 
24}]
    excel_file_list.update(files)