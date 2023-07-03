from openpyxl import load_workbook
from openpyxl import Workbook

datas = []
imported_workbooks = []
selected_columns = []
column_titles = []
columns_name = []


def add_file(filenames):
    for filename in filenames:
        if filename in imported_workbooks:
            print("filename ALREADY selected")
        else:
            ret = analysis_excel(filename, get_file_type(filename))
            datas.append({"filename": ret[0], "sheetnames": ret[1], "selected": ret[2], 
                              "rows": ret[3], "columns": ret[4]})
            imported_workbooks.append(filename)
    return datas


def delete_file(index):
    print(f"removed: {datas[index]}")
    imported_workbooks.remove(datas[index]["filename"])
    datas.remove(datas[index])
    return datas


def get_columns():
    first_one = datas[0]
    return   get_columns_by_filetype(first_one, get_file_type(first_one["filename"]))


def export_file(filename, selected_column, seperate_conditions):
    return export_file_by_filetype(filename, selected_column, seperate_conditions, get_file_type(filename))


def get_file_type(filename):
    if len(filename.split(".xlsx")) > 1:
        return "xlsx"
    elif len(filename.split(".xls")) > 1:
         return "xls"
    return ""


def analysis_excel(filename, filetype):
    wb = None
    sh = None
    if filetype == "xls":
        print("xls")
    elif filetype == "xlsx":
        wb = load_workbook(filename=filename)
        sh = wb.get_sheet_by_name(wb.sheetnames[0])
    return [filename, wb.sheetnames, 0, sh.max_row, sh.max_column]


def get_columns_by_filetype(first_one, filetype):
    global column_titles
    global columns_name

    if filetype == "xlsx":
        selected_sheet_index = first_one["selected"]
        max_column = first_one["columns"]
        wb = load_workbook(filename=first_one["filename"])
        sh = wb.get_sheet_by_name(wb.sheetnames[selected_sheet_index])
        column_titles = [sh.cell(1, i).value for i in range(1, max_column+1)]
        columns_name = [sh.cell(2, i).value for i in range(1, max_column+1)]
        #+1的是代表【全选】
        selected_columns = [1 for i in range(len(column_titles)+1)]

        return [column_titles, selected_columns]
    else:
        return []


def export_file_by_filetype(filename, selected_column, seperate_conditions, filetype):
    #filename：导出文件的名称
    #selected_column：选择导出的列
    #seperate_conditions：条件列中的过滤项
    if filetype == "xlsx":
        new_excel = Workbook()
        sheet = new_excel.active
        sheet.title = "Sheet1"
        for i in range (len(selected_column)-1):
            if selected_column[i] == 1:
                sheet.cell(1, i+1).value = column_titles[i]
                sheet.cell(2, i+1).value = columns_name[i]
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
    
    elif filetype == "xls":
        return "导出成功！"

if __name__ == "__main__":
    filenames = ('D:/Windows数据移动到此文件夹/Desktop/测试用例导出.xlsx', 'D:/Windows数据移动到此文件夹/Desktop/测试用例导出 - 副本.xlsx')
    add_file(filenames)
    get_columns()