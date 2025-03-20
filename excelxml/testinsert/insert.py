import xml.etree.ElementTree as ET

# 读取 XML 文件
tree = ET.parse(r'C:\Users\chengzhaofeng\Desktop\python\testinsert\example.xml')
root = tree.getroot()

# 找到 id="2" 的节点
target_node = None
for item in root.findall('item'):
    if item.get('id') == '2':
        target_node = item
        break

if target_node is not None:
    # 创建新节点
    new_node = ET.Element('item', {'id': '4'})
    name_element = ET.SubElement(new_node, 'name')
    name_element.text = 'Item 4'

    # 在目标节点后插入新节点
    index = list(root).index(target_node)
    root.insert(index + 1, new_node)

    # 写回原来的文件
    tree.write(r'C:\Users\chengzhaofeng\Desktop\python\testinsert\example.xml', encoding='utf-8', xml_declaration=True)
    print("新节点已添加，文件已更新")
else:
    print("未找到目标节点")