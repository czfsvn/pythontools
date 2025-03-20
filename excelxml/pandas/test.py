import pandas as pd
import xlrd

# 读取 .xls 文件
file_path = r'C:\Users\chengzhaofeng\Desktop\python\pandas\pands_test.xls'
df = pd.read_excel(file_path, engine='xlrd')

# 假设要更新第 2 行（索引为 1）、名为 'Column1' 的列的值
row_index = 1
column_name = 'Column1'
new_value = 'Updated Value'

# 更新指定单元格的值
df.at[row_index, column_name] = new_value

# 将更新后的数据保存到新的 .xls 文件
output_file = 'updated_file.xls'
df.to_excel(output_file, index=False, engine='openpyxl'