import xml.etree.ElementTree as ET
from xml.dom import minidom
from pathlib import Path

from openpyxl import load_workbook

def safe_read_headers(file_path, sheet_name=0):
    try:
        wb = load_workbook(file_path)
        if isinstance(sheet_name, int):
            ws = wb.worksheets[sheet_name]
        else:
            ws = wb[sheet_name]
        
        # 检查是否为空表
        if ws.max_row == 0:
            return []
        
        # 返回表头
        return [cell.value for cell in ws[1]]
    except FileNotFoundError:
        print(f"错误：文件 {file_path} 不存在！")
        return []
    except KeyError:
        print(f"错误：工作表 {sheet_name} 不存在！")
        return []

def read_excel(path):
    wb = load_workbook(path);
    ws = wb.active;

    # 遍历所有行
    datas = [];
    for row in ws.iter_rows(min_row=2, values_only=True):
        print("row: ", row);
        datas.append(row);

    last_row_num = ws.max_row;
    last_col_num = ws.max_column;

    print(f"last_row_num: {last_row_num}");
    print(f"last_col_num: {last_col_num}");

    for row in range(1, last_row_num + 1):
        for col in range(1, last_col_num + 1):
            cell_value = ws.cell(row=row, column=col).value;
            #print("cell_value: ", cell_value);
    
    # 获取最后一行的数据ss
    last_row_data = []
    for cell in ws[last_row_num]:
        last_row_data.append(cell.value)

    print("最后一行的数据:", last_row_data)

    print("最后一行的数据:", ws[last_row_num - 1])

    # 追加excel文件
    #ws.append(last_row_data)
    #wb.save(path);

    #return last_row_data;
    return datas;


def write_excel(path, data):
    print("write_excel");

def save_pretty_xml(element, filename):
    # 转换为字符串并解析
    raw_str = ET.tostring(element, encoding="utf-8")
    dom = minidom.parseString(raw_str)
    # 生成格式化 XML
    pretty_str = dom.toprettyxml(indent="\t", encoding="utf-8")
    # 过滤空行
    filtered_str = "\n".join(
        [line for line in pretty_str.decode("utf-8").split("\n") if line.strip()]
    )
    # 保存文件
    with open(filename, "w", encoding="utf-8") as f:
        f.write(filtered_str)

def indent(elem, level=0):
    i = "\n" + level*"  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level+1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i

def write_xml(headers, data, output_path):
    print("headers:",  headers);

     # 创建 XML 根元素
    root = ET.Element('Config')

    for row in data:
        record = ET.SubElement(root, 'Data')
        for col_index, value in enumerate(row):
            # print("colname=%s, col_index=%u, value=%s " %(headers[col_index], col_index, value));    
            # field = ET.SubElement(record, headers[col_index])
            record.set(headers[col_index], str(value));

    # 第一种格式化方式
    # save_pretty_xml(root, output_path)

    # 第二种格式化方式
    indent(root)
    # 创建 XML 树
    tree = ET.ElementTree(root)
    # 写入 XML 文件
    tree.write(output_path, encoding='utf-8', xml_declaration=True)

# --------------------------
# 主程序
# --------------------------
if __name__ == "__main__":
    # 文件路径配置
    input_excel = Path(__file__).parent/'input.xlsx';
    output_xml = Path(__file__).parent/"output.xml"
    output_excel = Path(__file__).parent/"output.xlsx"
    
    headers = safe_read_headers(input_excel)
    print("\nheaders: ", headers);

    # 执行操作
    data = read_excel(input_excel)
    print("data: ", data);

    write_xml(headers, data, output_xml);
    # write_excel(data, output_excel)
    # append_excel(data, output_excel) 
    
    print("处理完成！XML 和 Excel 文件已生成。")