import xlrd
import xlwt
from xlutils.copy import copy

# 数据类型常量
XL_CELL_EMPTY = 0
XL_CELL_TEXT = 1
XL_CELL_NUMBER = 2
XL_CELL_DATE = 3
XL_CELL_BOOLEAN = 4
XL_CELL_ERROR = 5

def test_read():
    # 打开 Excel 文件
    workbook = xlrd.open_workbook(r'C:\Users\chengzhaofeng\Desktop\python\xlrd\test_xlrd.xls')

    # 或者通过表名获取工作表
    sheet = workbook.sheet_by_name('Sheet1')

    # 获取工作表的行数和列数
    rows = sheet.nrows
    columns = sheet.ncols

    print(f'工作表的行数为: {rows}，列数为: {columns}')

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

def test_write():
    # 创建一个新的 Excel 工作簿
    # 打开要修改的 Excel 文件
    filename    = r'C:\Users\chengzhaofeng\Desktop\python\xlrd\test_xlrd.xls';
    workbook = xlrd.open_workbook(filename, formatting_info=True)

    # 或者通过表名获取工作表
    sheet = workbook.sheet_by_name('Sheet1')

    # 使用 xlutils.copy 复制工作簿为可写的副本
    wb_copy = copy(workbook)

    # 或者通过表名获取工作表
    sheet_copy  = wb_copy.get_sheet(0)

    # 修改指定单元格的值
    sheet_copy.write(1, 1, '修改后的值')  # 将第 2 行第 2 列的单元格值修改为 '修改后的值'

    # 保存修改后的副本
    wb_copy.save(filename)

test_read();
test_write();
