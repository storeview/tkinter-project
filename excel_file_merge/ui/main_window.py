from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from controller import controller
from tkinter import Checkbutton
from ui.excel_file_list import ExcelFileListFrame
from ui.excel_condition_filter_group import ExcelConditionFilterGroup
from ui.excel_column_export_group import ExcelColumnExportGroup
from ui.my_label import MyLabel

window_width = 800
window_height = 600

ret = []
columns = []
labels = []

def run():

#设置主窗口的标题、大小
   window = Tk()

   global ret
   global columns

   #居中显示
   win_width = window.winfo_screenwidth()
   win_height = window.winfo_screenheight()
   location_x = int(win_width/2-window_width/2)
   location_y = int(win_height/2-window_height/2)

   window.title("Excel文件合并处理")
   window.geometry(f"{window_width}x{window_height}+{location_x}+{location_y}")

   notebook = Notebook(window)

   #取消notebook的矩形虚线
   style = Style()
   style.layout("Tab", [('Notebook.tab', {'sticky': 'nswe', 'children':
      [('Notebook.padding', {'side': 'top', 'sticky': 'nswe', 'children':
         [('Notebook.label', {'side': 'top', 'sticky': ''})],
      })],
   })]
   )
   style.configure("Tab", focuscolor=style.configure(".")["background"])

   #文件导入
   import_file_frame = Frame(notebook, padding=10)
   notebook.add(import_file_frame, text=" 文件导入 ")

   excel_file_list = ExcelFileListFrame(import_file_frame)

   def import_file_callback():
      global ret
      filenames = filedialog.askopenfilenames(title="添加文件", filetypes=(("Excel 工作簿", "*.xlsx"),("Excel 97-2003 工作簿", "*.xls")))
      ret = controller.add_file(filenames)
      excel_file_list.update(ret)
      delete_buttons = excel_file_list.get_delete_buttons()
      if len(ret) != 0:
         next_step_button.config(state="normal")
         for delete_button in delete_buttons:
            delete_button.bind("<Button-1>", delete_file_callback)
         next_step_label.text("可执行『下一步』")
      else:
         next_step_button.config(state="disabled")
         next_step_label.text("请先添加需要处理的文件，再执行下一步")
   
   def next_step_callback():
      notebook.tab(export_file_frame, state="normal")
      notebook.select(export_file_frame)
      global columns
      update_export_file_list(export_file_list, ret, export_file_label)
      columns = controller.get_columns()
      export_columns_group.update(columns[0], columns[1])

   def delete_file_callback(event):
      global ret
      index = event.widget.config("class")[4]
      ret = controller.delete_file(int(index))
      excel_file_list.destroy(int(index))
      excel_file_list.update(ret)
      delete_buttons = excel_file_list.get_delete_buttons()
      notebook.tab(export_file_frame, state="disable")
      if len(ret) != 0:
         for delete_button in delete_buttons:
            delete_button.bind("<Button-1>", delete_file_callback)
         next_step_button.config(state="normal")
         next_step_label.text("可执行『下一步』")
      else:
         next_step_button.config(state="disabled")
         next_step_label.text("请先添加需要处理的文件，再执行下一步")

   import_file_button = Button(import_file_frame, text="添加文件", command=import_file_callback, cursor="plus")
   import_file_button.config()
   import_file_button.grid(row=0, column=0)
   excel_file_list.grid(row=1, column=1)

   next_step_button = Button(import_file_frame, text="下一步", command=next_step_callback, state="disabled")
   next_step_button.grid(row=3, column=0)
   next_step_label = MyLabel(import_file_frame, text="请先添加需要处理的文件，再执行下一步")
   next_step_label.grid(row=4, column=1)

   #功能：文件导出
   def export_file():
      filetypes=(("Excel 工作簿", "*.xlsx"),("Excel 97-2003 工作簿", "*.xls"))
      filename = filedialog.asksaveasfilename(title="导出文件", filetypes=filetypes,
                                   initialfile="导出", defaultextension="xlsx")

      ret = controller.export_file(filename, columns[1], condition_filter_group.get_conditions())
      export_msg_label.text(ret)

   #文件导出
   export_file_frame = Frame(notebook, padding=10)
   notebook.add(export_file_frame, text="    导出    ", state="disabled")

   export_file_label = MyLabel(export_file_frame, text="已选择的文件")
   export_file_label.grid(row=0, column=0)
   export_file_list = Frame(export_file_frame)
   export_file_list.grid(row=1, column=1)

   #【选择需要导出的表格列】
   export_columns_label = MyLabel(export_file_frame, text="选择需要导出的表格列", width=25, anchor="w")
   export_columns_label.grid(row=2, column=0)
   export_columns_group = ExcelColumnExportGroup(export_file_frame)
   export_columns_group.grid(row=3, column=1)


   #【添加筛选条件】
   filter_condition_label = Button(export_file_frame, text="添加筛选条件", width=25,
                                    command=lambda: condition_filter_group.add(columns[0], columns[1]))
   filter_condition_label.grid(row=4, column=0)
   condition_filter_group = ExcelConditionFilterGroup(export_file_frame)
   condition_filter_group.grid(row=6, column=1)


   #文件导出按钮
   export_file_button = Button(export_file_frame, text="    开始导出    ", command=export_file)
   export_file_button.grid(row=7, column=0)

   export_msg_label = MyLabel(export_file_frame, width=25, text="")
   export_msg_label.grid(row=8, column=1)

   notebook.pack(padx=10, pady=10, fill=BOTH, expand=True)

   window.mainloop()


def update_export_file_list(export_file_list, file_list, export_file_label):
   total = len(file_list)
   children = export_file_list.children
   val = list(children.keys())
   export_file_label.text(f"已选择的文件（{total}）")
   #for v in val:
   #   children[v].destroy()
   for index, file in enumerate(file_list):
      text = file["filename"].split("/")[-1:][0]
      file_name_label1 = MyLabel(export_file_list, text=text, width=40, anchor="w")
      #必须加上一个这一行，否则循环中的变量会被垃圾回收机制清除掉
      labels.append(file_name_label1)
      file_name_label1.grid(row=int(index/2), column=1+int(index%2))

if __name__ == "__main__":
   run()