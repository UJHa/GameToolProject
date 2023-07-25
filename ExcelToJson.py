import xlrd3
from collections import OrderedDict
import json

excel_path = 'C:\\Users\\user\\Documents\\Github\\zxcvProject\\DesignData\\AnimationInfo.xlsx'
wb = xlrd3.open_workbook(excel_path)
sh = wb.sheet_by_index(0)

column_name_list = []

column_name_list = sh.row_values(0)
column_name_list[0] = "id"
print(column_name_list)

data_list = []

for rownum in range(1, sh.nrows):
    data = OrderedDict()
    row_values = sh.row_values(rownum)
    for idx in range(0, len(column_name_list)):
        columnName = column_name_list[idx]
        rowIndex = idx
        data[columnName] = row_values[rowIndex]
        print(f"columnName({columnName})\trowIndex({rowIndex})\trow_value({row_values[rowIndex]})")
    data_list.append(data)

j = json.dumps(data_list, ensure_ascii=False)

filePath = '\\'.join(excel_path.split("\\")[:-1])
fileName = excel_path.split("\\")[-1].split(".")[0] + '.json'
jsonFullPath = "\\".join([filePath, fileName])
print(f"filePath({filePath})")
print(f"fileName({fileName})")
print(f"jsonFullPath({jsonFullPath})")
with open(jsonFullPath, 'w+') as f:
    f.write(j)