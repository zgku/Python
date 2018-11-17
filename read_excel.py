# extract number of rows using Python
import xlrd

# Give the location of the file
loc = ("D:\study.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

#获取目标EXCEL文件sheet名
sheet_names = wb.sheet_names()
print(sheet_names)

# Extracting number of rows 提取行数
print(sheet.nrows)

# extract number of columns in Python 提取列数
print(sheet.ncols)


# extracting all columns name in Python 提取所有列的名称
for i in range(sheet.ncols):
    print(sheet.cell_value(0, i))


#extracting first column  提取第一列
sheet = wb.sheet_by_index(0)
for i in range(sheet.nrows):
    print(sheet.cell_value(i, 2))

#extract a particular row value 提取特定行值
sheet = wb.sheet_by_index(0)
print(sheet.row_values(1))




