

import pandas as pd
import xml.etree.ElementTree as ET
import json
from pathlib import Path

class ConfigGenerator:
    def __init__(self):
        self.base_dir = Path(__file__).parent
        self.load_defaults()
    
    def load_defaults(self):
        # 加载默认值配置
        with open(self.base_dir/'defaults/a_defaults.json') as f:
            self.a_defaults = json.load(f)
        
        with open(self.base_dir/'defaults/b_defaults.json') as f:
            self.b_defaults = json.load(f)

    def process_input(self):
        # 读取用户输入
        input_path = self.base_dir/'input/C.xlsx'
        self.c_df = pd.read_excel(input_path, engine='openpyxl')

    def merge_data(self, row):
        """合并用户数据与默认值"""
        merged_a = {**self.a_defaults, **row.to_dict()}
        merged_b = {**self.b_defaults, **row.to_dict()}
        return merged_a, merged_b

    def generate_excel(self, data):
        # 生成A.xlsx
        output_path = self.base_dir/'output/A.xlsx'
        df = pd.DataFrame(data)
        df.to_excel(output_path, index=False, engine='openpyxl')
        print(f"生成A表完成：{output_path}")

    def generate_xml(self, data):
        # 生成B.xml
        root = ET.Element("Config")
        for item in data:
            record = ET.SubElement(root, "Record")
            for k, v in item.items():
                field = ET.SubElement(record, k)
                field.text = str(v)
        
        output_path = self.base_dir/'output/B.xml'
        tree = ET.ElementTree(root)
        tree.write(output_path, encoding='utf-8', xml_declaration=True)
        print(f"生成B表完成：{output_path}")

    def run(self):
        self.process_input()
        
        a_data = []
        b_data = []
        for _, row in self.c_df.iterrows():
            merged_a, merged_b = self.merge_data(row)
            a_data.append(merged_a)
            b_data.append(merged_b)
        
        self.generate_excel(a_data)
        self.generate_xml(b_data)

if __name__ == "__main__":
    generator = ConfigGenerator()
    generator.run()