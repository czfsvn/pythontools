import xlwings as xw
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
from pathlib import Path

# --------------------------
# 步骤 1：读取 Excel 数据
# --------------------------
def read_excel(file_path):
    """读取 Excel 数据并返回 DataFrame"""
    with xw.Book(file_path) as wb:
        sheet = wb.sheets[0]
        # 读取数据到 DataFrame
        df = sheet.range("A1").expand().options(pd.DataFrame, index=False).value
        print (df)
        return df

# --------------------------
# 步骤 2：数据处理（示例：筛选年龄 > 25 的数据）
# --------------------------
def process_data(df):
    """数据清洗与处理"""
    # 示例：筛选年龄大于 25 的行
    df_processed = df[df["Age"] > 25]
    # 添加新列（示例：计算出生年份）
    df_processed["BirthYear"] = 2023 - df_processed["Age"]
    return df_processed

# --------------------------
# 步骤 3：写入到新 Excel 文件
# --------------------------
def write_excel(output_path, df):
    """将 DataFrame 写入 Excel"""
    with xw.Book() as wb:
        sheet = wb.sheets[0]
        sheet.range("A1").value = df
        wb.save(output_path)
        print(f"数据已写入 Excel：{output_path}")

# --------------------------
# 步骤 4：生成 XML 文件
# --------------------------
def write_xml(output_path, df):
    """将 DataFrame 转换为 XML"""
    # 创建根节点
    root = ET.Element("Data")
    
    # 遍历 DataFrame 的每一行
    for _, row in df.iterrows():
        record = ET.SubElement(root, "Record")
        for field in df.columns:
            ET.SubElement(record, field).text = str(row[field])
    
    # 美化 XML 格式
    xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="  ")
    
    # 写入文件
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(xml_str)
    print(f"数据已写入 XML：{output_path}")

# --------------------------
# 步骤 3：写入到新 Excel 文件
# --------------------------
def append_excel(output_path, df):
    """将 DataFrame 写入 Excel"""
    with xw.Book(output_path) as wb:
        sheet = wb.sheets[0]
        # 查找最后一个非空行
        last_row = sheet.range("A1").end("down").row
        sheet.range(f"A{last_row + 1}").value = df
        wb.save()
        print("数据追加完成！")

def write_excel_ver2(output_path, df):
    """将 DataFrame 写入 Excel"""
    try:
        # 尝试打开现有的 Excel 文件
        app = xw.App(visible=False)
        wb = app.books.open(output_path)
        sheet = wb.sheets[0]
        last_row = sheet.range("A1").end("down").row
        print(f"成功打开文件: {output_path}")
    except FileNotFoundError:
        # 如果文件不存在，创建一个新的工作簿
        app = xw.App(visible=False)
        wb = app.books.add()
        sheet = wb.sheets[0]
        last_row = 0
        print(f"文件 {output_path} 不存在，已创建新的 Excel 文件。")
    except Exception as e:
        # 处理其他异常
        print(f"打开文件时出现其他错误: {e}")
        app.quit()
        raise

    if last_row == 0:
        sheet.range(f"A{last_row + 1}").value = df.columns.tolist()
        last_row = 1;
    
    # 查找最后一个非空行
    sheet.range(f"A{last_row + 1}").value = df.values
    wb.save(output_path)
    wb.close();
    app.quit();
    print("数据追加完成！")

# --------------------------
# 主程序
# --------------------------
if __name__ == "__main__":
    # 文件路径配置
    path=Path(__file__).parent;
    input_excel = path/"input.xlsx"
    output_excel = path/"output.xlsx"
    output_xml = path/"output.xml"
    
    # 读取数据
    df = read_excel(input_excel)
    print("原始数据：\n", df)
    
    # 处理数据
    df_processed = process_data(df)
    print("处理后的数据：\n", df_processed)
    
    # 写入 Excel
    #write_excel(output_excel, df_processed)

    for row in df_processed.values:
        print(row);
    
    print (df_processed.columns)
    print (df_processed.columns.tolist())

    
    write_excel_ver2(output_excel, df_processed)
    # 生成 XML
    #write_xml(output_xml, df_processed)
