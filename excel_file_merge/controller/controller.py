from service import service

def add_file(filename):
    print("add file...")

def add_file(filenames):
    return service.add_file(filenames)

def delete_file(index):
    return service.delete_file(index)

def get_columns():
    return service.get_columns()

def export_file(filename, selected_column, seperate_conditions):
    return service.export_file(filename, selected_column, seperate_conditions)

if __name__ == "__main__":
    filenames = ('D:/Windows数据移动到此文件夹/Desktop/测试用例导出.xlsx', 'D:/Windows数据移动到此文件夹/Desktop/测试用例导出 - 副本.xlsx')
    add_file(filenames)