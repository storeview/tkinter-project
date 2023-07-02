from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from controller import controller
from tkinter import Checkbutton
from ui.excel_file_list import ExcelFileListFrame

window_width = 800
window_height = 600

ret = []
columns = []
g_selected_columns = []
condition_num = 0
conditions = []

def run():

#设置主窗口的标题、大小
   window = Tk()

   global ret
   global columns
   global condition_num
   global conditions

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
         next_step_label_str.set("可执行『下一步』")
      else:
         next_step_button.config(state="disabled")
         next_step_label_str.set("请先添加需要处理的文件，再执行下一步")
   
   def next_step_callback():
      notebook.tab(export_file_frame, state="normal")
      notebook.select(export_file_frame)
      global columns
      print("update_export_file_list")
      update_export_file_list(export_file_list, ret, export_file_label_str)
      columns = controller.get_columns()
      update_export_columns_list(export_columns_list, columns)

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
         next_step_label_str.set("可执行『下一步』")
      else:
         next_step_button.config(state="disabled")
         next_step_label_str.set("请先添加需要处理的文件，再执行下一步")

   import_file_button = Button(import_file_frame, text="添加文件", command=import_file_callback, cursor="plus")
   import_file_button.config()
   import_file_button.grid(row=0, column=0)
   excel_file_list.grid(row=1, column=1)

   next_step_label_str = StringVar()
   next_step_label_str.set("请先添加需要处理的文件，再执行下一步")
   next_step_label = Label(import_file_frame, textvariable=next_step_label_str)
   next_step_button = Button(import_file_frame, text="下一步", command=next_step_callback, state="disabled")
   next_step_button.grid(row=3, column=0)
   next_step_label.grid(row=4, column=1)

   #功能：文件导出
   def export_file():
      filename = filedialog.asksaveasfilename(title="导出文件", 
                                   filetypes=(("Excel 工作簿", "*.xlsx"),("Excel 97-2003 工作簿", "*.xls")),
                                   initialfile="导出", defaultextension="xlsx")
      print(filter_condition_list)
      children = filter_condition_list.children
      seperate_conditions = []
      for child in children:
         pop = (str(child)).split("!").pop()
         if len(pop.split("combobox")) > 1:
            seperate_conditions.append(children[str(child)].current())
         elif len(pop.split("entry")) > 1:
            seperate_conditions.append(children[str(child)].get())

      print(seperate_conditions)
      ret = controller.export_file(filename, columns[1], seperate_conditions)
      export_file_ret_msg.set(ret)
      print(ret)



   #文件导出
   export_file_frame = Frame(notebook, padding=10)
   notebook.add(export_file_frame, text="    导出    ", state="disabled")

   export_file_label_str = StringVar()
   export_file_label_str.set("已选择的文件")
   export_file_label = Label(export_file_frame, textvariable=export_file_label_str, width=25, anchor="w")
   export_file_label.grid(row=0, column=0)
   export_file_list = Frame(export_file_frame)
   export_file_list.grid(row=1, column=1)

   export_columns_frame = Frame(export_file_frame)
   export_columns_frame.grid(row=2, column=0)
   export_columns_label = Label(export_columns_frame, text="选择需要导出的表格列", width=25, anchor="w")
   export_columns_label.grid(row=0, column=0)

   export_columns_list = Frame(export_file_frame)
   export_columns_list.grid(row=3, column=1)

   filter_condition_frame = Frame(export_file_frame)
   filter_condition_frame.grid(row=5, column=0)

   filter_condition_label = Button(filter_condition_frame, text="添加筛选条件", width=25,
                                    command=lambda: add_condition(filter_condition_list, columns))
   filter_condition_label.grid(row=0, column=0)
   filter_condition_list = Frame(export_file_frame)
   filter_condition_list.grid(row=6, column=1)


   #文件导出按钮
   export_file_button = Button(export_file_frame, text="    开始导出    ", command=export_file)
   export_file_button.grid(row=7, column=0)

   export_file_ret_msg = StringVar()
   export_file_ret_msg.set(" ")
   filter_condition_label = Label(export_file_frame, text="", width=25, textvariable=export_file_ret_msg)
   filter_condition_label.grid(row=8, column=1)


   notebook.pack(padx=10, pady=10, fill=BOTH, expand=True)




   def add_condition(filter_condition_list, columns):
      global condition_num
      column_titles = columns[0]
      selected_columns = columns[1]

      delete_condition_button = Button(filter_condition_list, text="X", width=2,  class_= str(condition_num))
      delete_condition_button.bind("<Button-1>", delete_condition_callback)
      column = Combobox(filter_condition_list, values=[column_titles[index] for index,val in enumerate(selected_columns) if val != 0 and index != len(column_titles)])
      column.current(0)
      condition = Combobox(filter_condition_list, values=["包含", "大于", "小于", "等于"])
      condition.current(0)
      val = Entry(filter_condition_list)

      delete_condition_button.grid(row=condition_num, column=0)
      column.grid(row=condition_num, column=1)
      condition.grid(row=condition_num, column=2)
      val.grid(row=condition_num, column=3)

      condition_num += 1


   def delete_condition_callback(event):
      index = event.widget.config("class")[4]
      name = event.widget._name
      children = filter_condition_list.children
      val = list(children.keys())
      index = val.index(name)
      children[val[index+3]].destroy()
      children[val[index+2]].destroy()
      children[val[index+1]].destroy()
      children[val[index]].destroy()
   window.mainloop()

def state_changed(export_columns_list):
   #更新已选择的字段
   global columns
   global g_selected_columns

   selected_columns = columns[1]
   print(selected_columns)
   for i in range(len(selected_columns)):
      checkbutton_state = g_selected_columns[i].get()
      selected_columns[i] = checkbutton_state
   print(selected_columns)
   print()

def select_all(export_columns_list):
   global columns
   global g_selected_columns
   selected_columns = columns[1]
   select_all_checkbutton_status = g_selected_columns[len(g_selected_columns)-1]
   #取消全选
   if select_all_checkbutton_status.get() == 0:
      for i in range(len(selected_columns)):
         selected_columns[i] = 0
      selected_columns[len(selected_columns)-1] = 0
   else:
      for i in range(len(selected_columns)):
         selected_columns[i] = 1
      selected_columns[len(selected_columns)-1] = 1
   update_export_columns_list(export_columns_list, columns)
   





def update_export_columns_list(export_columns_list, columns):
   children = export_columns_list.children
   val = list(children.keys())
   for v in val:
      children[v].destroy()
   column_titles = columns[0]
   selected_columns = columns[1]
   global g_selected_columns
   g_selected_columns = [IntVar() for i in range(len(column_titles)+1)]

   select_all_checkbutton_status = IntVar()
   select_all_checkbutton_status.set(selected_columns[len(column_titles)])
   g_selected_columns[len(column_titles)] = select_all_checkbutton_status

   select_all_checkbutton = Checkbutton(export_columns_list, text="全选", width=15, anchor="w", 
                                    onvalue=1, offvalue=0, variable=g_selected_columns[len(column_titles)],
                                    command=lambda: select_all(export_columns_list))
   
   select_all_checkbutton.grid(row=0, column=0)

   for index, column_title in enumerate(column_titles):
      column_checkbox = Checkbutton(export_columns_list, text=column_title, width=15, anchor="w", 
                                    onvalue=1, offvalue=0, variable=g_selected_columns[index],
                                    command=lambda: state_changed(export_columns_list))
      if selected_columns[index] == 1:
         g_selected_columns[index].set(1)
      else:
         g_selected_columns[index].set(0)
      index += 1
      column_checkbox.grid(row=int(int(index)/4), column=int(int(index)%4))

def update_export_file_list(export_file_list, file_list, export_file_label_str):
   total = len(file_list)
   children = export_file_list.children
   val = list(children.keys())
   export_file_label_str.set(f"已选择的文件（{total}）")
   for v in val:
      children[v].destroy()
   for index, file in enumerate(file_list):
      split_len = len(file["filename"].split("/"))
      file_name_label = Label(export_file_list, text=file["filename"].split("/")[split_len-1], width=40, anchor="w")
      file_name_label.grid(row=int(index/2), column=1+int(index%2))

if __name__ == "__main__":
   run()