xlrd 是 Python 中用于读取 Excel 文件（主要是 .xls 格式）的常用库。下面为你详细介绍 xlrd 读取 Excel 文件的多种方式。
安装 xlrd
如果你还没有安装 xlrd，可以使用以下命令进行安装：
bash
pip install xlrd
基本读取步骤
1. 打开 Excel 文件
python
import xlrd

# 打开 Excel 文件
```
workbook = xlrd.open_workbook('example.xls')
```
2. 获取工作表
可以通过索引或表名获取工作表。
python
# 通过索引获取第一个工作表
```
sheet_by_index = workbook.sheet_by_index(0)
```
# 通过表名获取指定工作表
```
sheet_by_name = workbook.sheet_by_name('Sheet1')
```
3. 获取工作表的基本信息
python
# 获取工作表的行数和列数
```
rows = sheet_by_index.nrows
columns = sheet_by_index.ncols
```

print(f'工作表的行数为: {rows}，列数为: {columns}')
读取单元格数据的方式
1. 读取单个单元格数据
python
# 读取指定单元格的数据
```
cell_value = sheet_by_index.cell_value(0, 0)  # 读取第 1 行第 1 列的数据
print(f'第 1 行第 1 列的数据为: {cell_value}')
```
2. 遍历所有单元格数据
python
# 遍历所有单元格数据
```
for row in range(rows):
    for col in range(columns):
        cell_value = sheet_by_index.cell_value(row, col)
        print(f'第 {row + 1} 行第 {col + 1} 列的数据为: {cell_value}')
```
3. 按行读取数据
python
# 按行读取数据
```
for row in range(rows):
    row_values = sheet_by_index.row_values(row)
    print(f'第 {row + 1} 行的数据为: {row_values}')
```
4. 按列读取数据
python
# 按列读取数据
```
for col in range(columns):
    col_values = sheet_by_index.col_values(col)
    print(f'第 {col + 1} 列的数据为: {col_values}')
```
处理不同类型的单元格数据
Excel 单元格的数据类型有多种，如字符串、数字、日期等，xlrd 可以通过 cell_type 方法获取单元格的数据类型。
```
python
import xlrd

workbook = xlrd.open_workbook('example.xls')
sheet = workbook.sheet_by_index(0)

# 数据类型常量
XL_CELL_EMPTY = 0
XL_CELL_TEXT = 1
XL_CELL_NUMBER = 2
XL_CELL_DATE = 3
XL_CELL_BOOLEAN = 4
XL_CELL_ERROR = 5

for row in range(sheet.nrows):
    for col in range(sheet.ncols):
        cell = sheet.cell(row, col)
        cell_type = cell.ctype
        cell_value = cell.value

        if cell_type == XL_CELL_TEXT:
            print(f'第 {row + 1} 行第 {col + 1} 列的数据类型为字符串，值为: {cell_value}')
        elif cell_type == XL_CELL_NUMBER:
            print(f'第 {row + 1} 行第 {col + 1} 列的数据类型为数字，值为: {cell_value}')
        elif cell_type == XL_CELL_DATE:
            # 处理日期类型的数据
            date_tuple = xlrd.xldate_as_tuple(cell_value, workbook.datemode)
            print(f'第 {row + 1} 行第 {col + 1} 列的数据类型为日期，值为: {date_tuple}')
        elif cell_type == XL_CELL_BOOLEAN:
            print(f'第 {row + 1} 行第 {col + 1} 列的数据类型为布尔值，值为: {bool(cell_value)}')
        elif cell_type == XL_CELL_ERROR:
            print(f'第 {row + 1} 行第 {col + 1} 列的数据类型为错误，值为: {cell_value}')
        else:
            print(f'第 {row + 1} 行第 {col + 1} 列的数据类型为空，值为: {cell_value}')

```
总结
xlrd 提供了多种方式来读取 Excel 文件中的数据，你可以根据具体需求选择合适的读取方式。同时，要注意处理不同类型的单元格数据，特别是日期类型的数据需要进行额外的转换。需要注意的是，从 2.0.0 版本开始，xlrd 不再支持读取 .xlsx 文件，仅支持 .xls 文件。