# -*- coding: GB2312 -*-
"""
Created on Wed Feb 20 18:05:35 2019

@author: Administrator
"""

import xml.etree.ElementTree as ET
from pathlib import Path
from openpyxl import load_workbook
import copy, os, sys;

class Config:
    def __init__(self):
        self.obj_excel = ""
        self.obj_sheet =  "";
        self.gift_excel =  "";
        self.gift_sheet =  "";
        self.gift_output_xml =  "";
        self.gift_template_xml = "";

class GiftXmlTempate:
    def __init__(self):
        self.root = None;

class NewGift:
    def __init__(self, id):
        self.id = id;
        self.field = [];   # gift的字段
        self.objs = []; 
        self.row = [];
        self.objs = [];
        self.row = [];
        self.xml = None
        self.templateID = 0;
        self.columns = []  # 存储需要修改的列名称
        self.newdata = [];  # 存储需要修改的列数据


def indent(elem, level=0):
    i = "\n" + level*"  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "	"
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level+1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i

class ConfigGenerator:
    def __init__(self, config):
        self.config = config; 
        self.gifts=[];
    
    def read_headers_openpyxl(self):
        try:
            wb = load_workbook(self.config.obj_excel)
            ws = wb[self.config.obj_sheet]
            
            # 检查是否为空表
            if ws.max_row == 0:
                return []
            
            # 返回表头
            return [cell.value for cell in ws[1]]
        except FileNotFoundError:
            print(f"错误：[read_headers] 文件 {self.config.obj_excel}/{self.config.obj_sheet} 不存在！")
            input("Press Enter to continue...")
            return []
        except KeyError:
            print(f"错误：[read_headers] 工作表 {self.config.obj_excel}/{self.config.obj_sheet} 不存在！")
            input("Press Enter to continue...")
            return []
        except Exception as e:
            print(f"错误：{e}, [read_headers] 工作表 {self.config.obj_excel}/{self.config.obj_sheet} 不可用");
            input("Press Enter to continue...")
            return []

    def read_lastrow_openpyxl(self):
        try:
            wb = load_workbook(self.config.obj_excel);
            ws = wb[self.config.obj_sheet];

            last_row_num = ws.max_row;
            last_col_num = ws.max_column;
        
            #print(f"last_row_num: {last_row_num}");
            #print(f"last_col_num: {last_col_num}");
    
            # 获取最后一行的数据
            last_row_data = []
            for cell in ws[last_row_num]:
                last_row_data.append(cell.value)

            #print("lastrow data: ", last_row_data);
            return last_row_data;
        except FileNotFoundError:
            print(f"错误：[read_lastrow] 文件 {self.config.obj_excel}/{self.config.obj_sheet} 不存在！")
            input("Press Enter to continue...")
            return []
        except KeyError:
            print(f"错误：[read_lastrow] 工作表 {self.config.obj_excel}/{self.config.obj_sheet} 不存在！")
            input("Press Enter to continue...")
            return []
        except Exception as e:
            print(f"错误：{e}, [read_lastrow] 工作表 {self.config.obj_excel}/{self.config.obj_sheet} 不可用")
            input("Press Enter to continue...")
            return None
        
    
    def read_objexcel_by_id(self, id):
        try:
            wb = load_workbook(self.config.obj_excel)
            ws = wb[self.config.obj_sheet]

            last_row_num = ws.max_row;
            last_col_num = ws.max_column;

            #print(f"last_row_num: {last_row_num}");
            #print(f"last_col_num: {last_col_num}");
        
            for row in ws.iter_rows(min_row=2, values_only=True):
                if int(row[0]) == id:
                    return row;
    
            return [];        
        except FileNotFoundError:
            print(f"错误：[read_objexcel_by_id] 文件 {self.config.obj_excel}/{self.config.obj_sheet} 不存在！")
            return []
        except KeyError:
            print(f"错误：[read_objexcel_by_id] 工作表 {self.config.obj_excel}/{self.config.obj_sheet} 不存在！")
            return []
        except Exception as e:
            print(f"错误：{e}, [read_objexcel_by_id] 工作表 {self.config.obj_excel}/{self.config.obj_sheet} 不可用！")
            input("Press Enter to continue...")
            return []
        
    def read_gift_template(self):
        try:        
            with open(self.config.gift_template_xml, "r", encoding="GB2312") as file:
                content = file.read()

            root = ET.fromstring(content, parser=ET.XMLParser(encoding="GB2312"))
            #tree = ET.parse(self.config.gift_template_xml)
            #root = tree.getroot()
            config = Config();
            for child in root:
                if child.tag == "bag":
                    self.bagattr = child.attrib;
                    for bagchild in child:
                        if bagchild.tag == "rule":
                            self.rulletype = bagchild.attrib["type"];
        except FileNotFoundError:
            print(f"错误：[read_gift_template] 文件 {self.config.gift_template_xml} 不存在！")
            input("Press Enter to continue...")
            return
        except KeyError:
            print(f"错误：[read_gift_template] 工作表 {self.config.gift_template_xml} 不存在！")
            input("Press Enter to continue...")
            return
        except Exception as e:
            print(f"错误：{e}, [read_gift_template] 工作表 {self.config.gift_template_xml} 不可用")
            input("Press Enter to continue...")
            return
        
    def getTemplateData(self, teamplateid, oldid):
        try:
            wb = load_workbook(self.config.obj_excel)
            ws = wb[self.config.obj_sheet]

            last_row_num = ws.max_row;
            last_col_num = ws.max_column;

            teamplatedata = [];
            olddata = [];
            max_id = 0;
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] == None:
                    continue;
                id = int(row[0]);
                if id > max_id:
                    max_id = id;
                if id == teamplateid:
                    teamplatedata = list(row);
                if id == oldid:
                    olddata = list(row);                    
    
            return (max_id, teamplatedata, olddata);
        except FileNotFoundError:
            print(f"错误：[getTemplateData] 文件 {self.config.obj_excel}/{self.config.obj_sheet} 不存在！")
            input("Press Enter to continue...")
            return []
        except KeyError:
            print(f"错误：[getTemplateData] 工作表 {self.config.obj_excel}/{self.config.obj_sheet} 不存在！")
            input("Press Enter to continue...")
            return []
        except Exception as e:
            print(f"错误：{e}, [getTemplateData] 工作表 {self.config.obj_excel}/{self.config.obj_sheet} 不可用")
            input("Press Enter to continue...")
            return []

    def generate_giftxml_old(self):
        print("generate xml");
        root = ET.Element('Config')
        for gift in self.gifts:
            gift_root = ET.SubElement(root, 'GiftBag')
            gift_root.set("giftid", str(gift.id));
            for row in gift.objs:
                record = ET.SubElement(gift_root, 'Data')
                for col_index, value in enumerate(row):
                    record.set(gift.field[col_index], str(value));
    
        indent(root)
        # 创建 XML 树
        tree = ET.ElementTree(root)
        # 写入 XML 文件
        tree.write(self.config.gift_output_xml, encoding='GB2312', xml_declaration=True)

    def generate_giftxml(self):
        #print("generate xml");
        root = ET.Element('Config')
        for gift in self.gifts:
            bagroot = ET.SubElement(root, 'bag')
            for key, value in self.bagattr.items():
                if key == "id":
                    bagroot.set("id", str(gift.id));
                else:
                    bagroot.set(key, str(value));

            rulenode = ET.SubElement(bagroot, 'rule')
            rulenode.set("type", str(self.rulletype));
            
            for row in gift.objs:
                record = ET.SubElement(bagroot, 'item')
                for col_index, value in enumerate(row):
                    if value == None:
                        continue;
                    record.set(gift.field[col_index], str(value));
            
        indent(root)
        # 创建 XML 树
        tree = ET.ElementTree(root)
        # 写入 XML 文件
        tree.write(self.config.gift_output_xml, encoding='GB2312', xml_declaration=True)


    
    
    def update_old(self, gift):
        try:
            wb = load_workbook(self.config.obj_excel)
            ws = wb[self.config.obj_sheet]

            updated = {};
            log_dict = {};
            for index, headername in enumerate(self.header):
                for col_index, value in enumerate(gift.columns):
                    if headername == value:
                        updated[index + 1] = gift.newdata[col_index];
                        log_dict[headername] = gift.newdata[col_index];

            needsave = False;
            for row in ws.iter_rows(min_row=2):
                findrow = False;
                savedcount = 0;
                for cell in row:
                    if findrow == False and cell.column != 1:
                        continue;
                    
                    if cell.value == gift.id:
                        #print(f"值: {cell.value}")
                        #print(f"坐标: {cell.coordinate}")  # 如 A2、B3
                        #print(f"行号: {cell.row}, 列号: {cell.column}")
                        #print("---")
                        findrow = True;

                    if findrow == True:
                        if cell.column in updated:
                            cell.value = updated[cell.column];
                            needsave = True;
                            savedcount = savedcount + 1;

                            
                    if savedcount == len(updated):
                        break;
                
                if savedcount == len(updated):
                    break;

            if needsave:
                wb.save(self.config.obj_excel)
                print("[OK]: update old data success, id: ", gift.id)
                print("[OK]: update old data success, updated: ", log_dict);        
    
        except FileNotFoundError:
            print(f"错误：[update_old] 文件 {self.config.gift_template_xml} 不存在！")
            input("Press Enter to continue...")
            return []
        except KeyError:
            print(f"错误：[update_old] 工作表 {self.config.gift_template_xml} 不存在！")
            input("Press Enter to continue...")
            return []
        except Exception as e:
            print(f"错误：{e}, [update_old] 工作表 {self.config.gift_template_xml} 不可用")
            input("Press Enter to continue...")
            return []
    
    def fill_gift_row_new(self, gift):
        for index, headername in enumerate(self.header):
            #print(f"索引 {index}: {headername}")
            for col_index, value in enumerate(gift.columns):
                if headername == value:
                    gift.row[index] = gift.newdata[col_index];
                    break;

    def fill_gift_row(self, gift):
        data = self.getTemplateData(gift.templateID, gift.id);
        if data == None:
            print("error: getTemplateData failed");
            return;

        max_id = data[0];
        teamplatedata = data[1];
        olddata = data[2];
    
        if gift.id == 0:    # 需要产出新道具
            if teamplatedata == None or len(teamplatedata) == 0:
                print("error: teamplatedata is None");
                return;
        
            gift.id = max_id + 1;
            gift.row = teamplatedata;
            gift.row[0] = gift.id;
            self.fill_gift_row_new(gift);
            self.write_objexcel(gift);

        elif gift.id > 0:
            if olddata != None and len(olddata) > 0: # 需要覆盖老道具信息, 且使用gift.id
                # todo: 覆盖旧数据
                self.update_old(gift);
                return;            
            else:   # 需要产出新道具信息， 且使用gift.id
                if teamplatedata == None or len(teamplatedata) == 0:
                    print("error: teamplatedata is None");
                    return;
                gift.row = teamplatedata;
                gift.row[0] = gift.id;
                self.fill_gift_row_new(gift);
                self.write_objexcel(gift);
    
    
    def write_objexcel(self, gift):
        try:
            if len(gift.row) == 0:
                return;

            wb = load_workbook(self.config.obj_excel)
            ws = wb[self.config.obj_sheet]
            
            #print("write excel, giftid: ", gift.id);
            #print("write excel, giftrow: ", gift.row);

            ws.append(gift.row)
            wb.save(self.config.obj_excel);
        except FileNotFoundError:
            print(f"错误：[write_objexcel] 文件 {self.config.obj_excel}/{self.config.obj_sheet}  不存在！")
            input("Press Enter to continue...")
            return []
        except KeyError:
            print(f"错误：[write_objexcel] 工作表 {self.config.obj_excel}/{self.config.obj_sheet} 不存在！")
            input("Press Enter to continue...")
            return []
        except Exception as e:
            print(f"错误：{e}, [write_objexcel] 工作表 {self.config.obj_excel}/{self.config.obj_sheet}  不可用！")
            input("Press Enter to continue...")
            return []
    
    def process_all_gifts(self):        
        for gift in self.gifts:
            self.fill_gift_row(gift);
        
        print("[%s][%s] [excel]更新完成" %(self.config.obj_excel, self.config.obj_sheet))
        
        self.generate_giftxml();
        print("[%s][xml]更新完成" %(self.config.gift_output_xml))
    
    def read_gifts(self):
        try: 
            wb = load_workbook(self.config.gift_excel)
            ws = wb[self.config.gift_sheet]

            last_row_num = ws.max_row;
            last_col_num = ws.max_column;
        
            #print(f"last_row_num: {last_row_num}");
            #print(f"last_col_num: {last_col_num}");

            # 遍历所有行
            gift = NewGift(0);
            
            for row in ws.iter_rows(min_row=1, values_only=True):
                if row[0] == None:
                    continue;
                if  row[0] == 'id':
                    if len(gift.field) == 0:
                        gift.field = row;
                elif row[0] == '礼包' and row[1] == '新道具':
                    #print("new gift: 需要创建新的道具");
                    self.emplace_gift(gift)
                    gift = NewGift(0);
                elif row[0] == '礼包' and row[1] != None and int(row[1] > 0):
                    #print("new gift: 需要覆盖老的道具");
                    self.emplace_gift(gift)
                    gift = NewGift(int(row[1]));
                    continue;
                elif row[0] == 'templateID':
                    gift.columns = row;
                else:
                    if len(gift.columns) != 0 and len(gift.newdata) == 0 and gift.templateID == 0:
                        gift.newdata = row;
                        gift.templateID = int(row[0]);
                    else:
                        gift.objs.append(row);
            
            self.emplace_gift(gift);            
        except FileNotFoundError:
            print(f"错误：[read_gifts] 文件 {self.config.gift_excel}/{self.config.gift_sheet} 不存在！")
            input("Press Enter to continue...")
            return []
        except KeyError:
            print(f"错误：[read_gifts] 工作表 {self.config.gift_excel}/{self.config.gift_sheet} 不存在！")
            input("Press Enter to continue...")
            return []
        except Exception as e:
            print(f"错误：{e}, [read_gifts] 工作表 {self.config.gift_excel}/{self.config.gift_sheet} 不可用")
            input("Press Enter to continue...")
            return []

    def run(self):
        #last_row_data = self.read_lastrow();
        #print("last_row_data: ",last_row_data);
        #self.header = self.read_headers();
        #print("headers: ",header);
        #self.read_gift_template();
        self.read_gifts();
        self.process_all_gifts();
        

def read_config(configpath):    
    try:
        with open(configpath, "r", encoding="GB2312") as file:
            content = file.read()

        root = ET.fromstring(content, parser=ET.XMLParser(encoding="GB2312"))
        config = Config();
        for child in root:
            #print(f"tagname={child.tag}, attrib={child.attrib}");
            if child.tag == "objitem":
                config.obj_excel = child.attrib["filepath"];
                config.obj_sheet = child.attrib["sheet"];
            if child.tag == "giftitem":
                config.gift_excel = child.attrib["filepath"]    
                config.gift_sheet = child.attrib["sheet"];
            elif child.tag == "giftxml":
                config.gift_output_xml = child.attrib["outpath"];
                config.gift_template_xml = child.attrib["templatexml"];

        return config;
    except FileNotFoundError:
        print(f"错误：[read_config] 文件 {configpath} 不存在！")
        input("Press Enter to continue...")
        return None
    except KeyError:
        print(f"错误：[read_config] 工作表 {configpath} 不存在！")
        input("Press Enter to continue...")
        return None
    except Exception as e:
        print(f"错误：{e}, [read_config] 工作表[read_config] {configpath} 不可用")
        input("Press Enter to continue...")
        return None


if __name__=="__main__":
    base_path = Path(__file__).parent/"config.xml"
    config = read_config(base_path);
    #config = read_config('./config.xml');
    
    generator = ConfigGenerator(config);
    generator.run()
    input("Press Enter to continue...")

