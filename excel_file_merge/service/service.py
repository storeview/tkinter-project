from openpyxl import load_workbook
from openpyxl import Workbook

datas = []
selected_workbooks = []
selected_columns = []
column_titles = []
columns_name = []

def add_file(filenames):
    for filename in filenames:
        if filename in selected_workbooks:
            print("filename ALREADY selected")
        else:
            wb = load_workbook(filename=filename)
            sh = wb.get_sheet_by_name(wb.sheetnames[0])
            datas.append({"filename": filename, "sheetnames": wb.sheetnames, "selected": 0, 
                              "rows": sh.max_row, "columns": sh.max_column})
            selected_workbooks.append(filename)
    return datas

def delete_file(index):
    print(f"removed: {datas[index]}")
    selected_workbooks.remove(datas[index]["filename"])
    datas.remove(datas[index])
    return datas

def get_columns():
    global column_titles
    global columns_name
    first_one = datas[0]
    filename = first_one["filename"]
    selected_sheet_index = first_one["selected"]
    max_column = first_one["columns"]
    wb = load_workbook(filename=filename)
    sh = wb.get_sheet_by_name(wb.sheetnames[selected_sheet_index])
    column_titles = [sh.cell(1, i).value for i in range(1, max_column+1)]
    columns_name = [sh.cell(2, i).value for i in range(1, max_column+1)]
    #+1的是代表【全选】
    selected_columns = [1 for i in range(len(column_titles)+1)]
    print( [column_titles, selected_columns])

    return [column_titles, selected_columns]


def export_file(filename, selected_column, seperate_conditions):
    #filename：导出文件的名称
    #selected_column：选择导出的列
    #seperate_conditions：条件列中的过滤项

    #1.创建一个Excel表格
    new_excel = Workbook()
    sheet = new_excel.active
    sheet.title = "Sheet1"
    for i in range (len(selected_column)-1):
        if selected_column[i] == 1:
            sheet.cell(1, i+1).value = column_titles[i]
            sheet.cell(2, i+1).value = columns_name[i]
    """
    export_columns = []
    for i in range(len(selected_column)-1):
        #选择导出的列
        if selected_column[i] == 1:
            export_columns.append(column_titles[i])
    print(f"被选择的列：{export_columns}")        
    """

    index = 3
    for excel in datas:
        wb = load_workbook(filename=excel["filename"])
        sh = wb.get_sheet_by_name(wb.sheetnames[excel["selected"]])
        rows = list(sh.rows)
        for row in rows[2:]:
            for i in range (len(selected_column)-1):
                if selected_column[i] == 1:
                    sheet.cell(index, i+1).value = row[i].value
            index += 1

    new_excel.save(filename)

    return "导出成功！"


    

if __name__ == "__main__":
    filenames = ('D:/Windows数据移动到此文件夹/Desktop/测试用例导出.xlsx', 'D:/Windows数据移动到此文件夹/Desktop/测试用例导出 - 副本.xlsx')
    add_file(filenames)
    get_columns()