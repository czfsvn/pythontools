# -*- coding: GB2312 -*-

import xlrd
import xlwt
from pathlib import Path
import xml.etree.ElementTree as ET
import os, sys, xlutils;
from xlutils.copy import copy
import copy as cp
import re


from lxml import etree

import logging
import inspect
import traceback
import subprocess

# 数据类型常量
XL_CELL_EMPTY = 0
XL_CELL_TEXT = 1
XL_CELL_NUMBER = 2
XL_CELL_DATE = 3
XL_CELL_BOOLEAN = 4
XL_CELL_ERROR = 5

# 创建日志记录器
logger = logging.getLogger(__name__)

def custom_exception_handler(exc_type, exc_value, exc_traceback):
    # 获取错误发生的行号
    line_number = exc_traceback.tb_lineno
    # 获取错误发生的文件名
    file_name = inspect.getframeinfo(exc_traceback.tb_frame).filename
    logger.critical(f"在文件 {file_name} 的第 {line_number} 行发生了 {exc_type.__name__} 异常: {exc_value}")

# 设置自定义异常处理函数
sys.excepthook = custom_exception_handler

# 自定义解析器，保留注释
class CommentParser(ET.XMLParser):
    def __init__(self):
        super().__init__()
        self._parser.CommentHandler = self.handle_comment

    def handle_comment(self, data):
        self._target.start(ET.Comment, {})
        self._target.data(data)
        self._target.end(ET.Comment)

class PackToolConfig:
    def __init__(self):
        self.exec_path = "";
        self.input_path = "";
        self.output_path = "";
        self.arg = "";

class Config:
    def __init__(self):
        self.obj_excel = ""
        self.obj_sheet =  "";
        self.gift_excel =  "";
        self.gift_sheet =  "";
        self.gift_output_xml =  "";
        self.gift_template_xml = "";
        self.packtool = []; # 打包工具配置

class NewGift:
    def __init__(self, id):
        self.id = id;
        self.field = [];   # gift的字段
        self.objs = []; 
        self.row = [];
        self.row = [];
        self.xml = None
        self.templateID = 0;
        self.columns = []  # 存储需要修改的列名称
        self.newdata = [];  # 存储需要修改的列数据

def initlogger():
    logger.setLevel(logging.DEBUG)

    # 创建文件处理器
    file_handler = logging.FileHandler('./app.log')
    file_handler.setLevel(logging.DEBUG)

    # 创建控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # 定义日志格式
    formatter = logging.Formatter('%(asctime)s [%(levelname)s][%(filename)s:%(lineno)d] %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # 将处理器添加到日志记录器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)


def read_config(configpath):    
    try:
        with open(configpath, "r", encoding="GB2312") as file:
            content = file.read()

        root = ET.fromstring(content, parser=ET.XMLParser(encoding="GB2312"))
        config = Config();
        for child in root:
            #logging.debug(f"tagname={child.tag}, attrib={child.attrib}");
            if child.tag == "objitem":
                config.obj_excel = child.attrib["filepath"];
                config.obj_sheet = child.attrib["sheet"];
            if child.tag == "giftitem":
                config.gift_excel = child.attrib["filepath"]    
                config.gift_sheet = child.attrib["sheet"];
            elif child.tag == "giftxml":
                config.gift_output_xml = child.attrib["outpath"];
                config.gift_template_xml = child.attrib["templatexml"];
            elif child.tag == "packtools":
                packtool = PackToolConfig();
                packtool.exec_path = child.attrib["execpath"];
                packtool.input_path = child.attrib["inputpath"]
                packtool.output_path = child.attrib["outputpath"]
                packtool.arg = child.attrib["arg"]
                config.packtool.append(packtool);


        return config;
    except FileNotFoundError:
        logger.error(f"错误：[read_config] 文件 {configpath} 不存在,  {traceback.format_exc()}")
        #input("Press Enter to continue...")
        return None
    except KeyError:
        logger.error(f"错误：[read_config] 工作表 {configpath} 不存在,  {traceback.format_exc()}")
        #input("Press Enter to continue...")
        return None
    except Exception as e:
        logger.error(f"错误：{e}, [read_config] 工作表[read_config] {configpath} 不可用,  {traceback.format_exc()}")
        #input("Press Enter to continue...")
        return None
    
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
    
    def read_excel_headers(self):
        try:
            logger.debug(f"read excel headers begin {self.config.obj_excel}\\{self.config.obj_sheet}");
            # 打开 Excel 文件
            workbook = xlrd.open_workbook(self.config.obj_excel, formatting_info=True)

            # 或者通过表名获取工作表
            sheet = workbook.sheet_by_name(self.config.obj_sheet)

            # 获取工作表的行数和列数
            rows = sheet.nrows
            columns = sheet.ncols

            self.headers = {};
            for col in range(columns):
                header = sheet.cell_value(0, col)
                self.headers[col] = header;
                #print(f"header: {header}, col: {col}");

            logger.debug(f"headers: {self.headers}");
            logger.debug(f"read excel headers success");
        except FileNotFoundError:
            logger.error(f"错误：[read_excel_headers] 文件 {self.config.obj_excel}\\{self.config.obj_sheet}  不存在,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        except KeyError:
            logger.error(f"错误：[read_excel_headers] 工作表 {self.config.obj_excel}\\{self.config.obj_sheet} 不存在,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        except Exception as e:
            logger.error(f"错误：{e}, [read_excel_headers] 工作表 {self.config.obj_excel}\\{self.config.obj_sheet} 不可用,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        
    
    def emplace_gift(self, gift):
        if (len(gift.objs) == 0):
            return;

        self.gifts.append(gift);
    
    def getTemplateData(self, teamplateid, oldid):
        try:
            workbook = xlrd.open_workbook(self.config.obj_excel, formatting_info=True)
            sheet = workbook.sheet_by_name(self.config.obj_sheet)

            max_row_num = sheet.nrows;
            max_col_num = sheet.ncols;

            teamplatedata = [];
            olddata = [];
            max_id = 0;
            for row in range(1, max_row_num):
                row_values = sheet.row_values(row)
                if row_values[0] == None or not row_values[0]:
                    continue;
                
                if row_values[0].is_integer() == False:
                    continue;
                
                id = int(row_values[0]);
                if id > max_id:
                    max_id = id;
                if id == teamplateid:
                    teamplatedata = list(row_values);
                if id == oldid:
                    olddata = list(row_values);                    
    
            return (max_id, teamplatedata, olddata);
        except FileNotFoundError:
            logger.error(f"错误：[getTemplateData] 文件 {self.config.obj_excel}\\{self.config.obj_sheet} 不存在,  {traceback.format_exc()}")
            #("Press Enter to continue...")
            return []
        except KeyError:
            logger.error(f"错误：[getTemplateData] 工作表 {self.config.obj_excel}\\{self.config.obj_sheet} 不存在,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return []
        except Exception as e:
            logger.error(f"错误：{e}, [getTemplateData] 工作表 {self.config.obj_excel}\\{self.config.obj_sheet} 不可用,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return []
        
    def update_old(self, gift):
        try:
            # 打开现有的 Excel 文件
            old_workbook = xlrd.open_workbook(self.config.obj_excel, formatting_info=True)
            # 获取第一个工作表
            old_sheet = old_workbook.sheet_by_name(self.config.obj_sheet)
        except FileNotFoundError:
            logger.error(f"错误：[update_old] 文件 {self.config.obj_excel}\\{self.config.obj_sheet}  不存在,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        except KeyError:
            logger.error(f"错误：[update_old] 工作表 {self.config.obj_excel}\\{self.config.obj_sheet} 不存在,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        except Exception as e:
            logger.error(f"错误：{e}, [update_old] 工作表 {self.config.obj_excel}\\{self.config.obj_sheet}  不可用,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        
        try:
            # 或者通过表名获取工作表
            # 使用 xlutils.copy 复制工作簿为可写的副本
            wb_copy = copy(old_workbook)

            # 获取原工作簿中该工作表的索引
            sheet_index = old_workbook.sheet_names().index(self.config.obj_sheet)
            # 或者通过表名获取工作表
            sheet_copy  = wb_copy.get_sheet(sheet_index)

            needsave = False;
            for row in range(1, old_sheet.nrows):
                cell = old_sheet.cell(row, 0)
                if cell == None:
                    continue;
                
                cell_type = cell.ctype
                cell_value = cell.value
                if cell_type == XL_CELL_NUMBER:
                    cell_value = int(cell_value)
                    if cell_value != gift.id:
                        continue;

                for index, value in enumerate(gift.columns):
                    if value == None or not value:
                        continue;

                    col_index = self.getColIndexByColumnName(value);
                    if col_index == None or col_index < 0 or index >= len(gift.newdata):
                        #print(f"error: getColIndexByColumnName failed, col_index: {col_index}, value: {value}");
                        continue;

                    needsave = True;
                    sheet_copy.write(row, col_index, gift.newdata[index])
                    logger.debug(f"update_excel: {gift.id}, {value}, {gift.newdata[index]}");

            if needsave == True:
                wb_copy.save(self.config.obj_excel)
                logger.debug(f"修改[{self.config.obj_excel}\\{self.config.obj_sheet}]数据成功，新文件已保存！ giftid={gift.id}")


        except Exception as e:
            logger.error(f"错误：{e}, [update_old] 工作表 {self.config.obj_excel}\\{self.config.obj_sheet}  不可用,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        
    def write_objexcel(self, gift):
        try:
            # 打开现有的 Excel 文件
            old_workbook = xlrd.open_workbook(self.config.obj_excel)
            # 获取第一个工作表
            old_sheet = old_workbook.sheet_by_name(self.config.obj_sheet)
            # 获取现有数据的行数
            rows = old_sheet.nrows

            # 创建一个新的工作簿
            new_workbook = xlwt.Workbook()
            # 在新工作簿中添加一个工作表
            new_sheet = new_workbook.add_sheet(self.config.obj_sheet)
            for row in range(rows):
                for col in range(old_sheet.ncols):
                    new_sheet.write(row, col, old_sheet.cell_value(row, col))

            # 追加新数据
            for col_index, value in enumerate(gift.row):
                new_sheet.write(rows, col_index, value);
                
            logger.debug(f"write_excel: {gift.id}, {gift.row}");

            # 保存修改后的 Excel 文件
            # 保存新的工作簿到新文件路径
            new_workbook.save(self.config.obj_excel)
            logger.debug(f"新增[{self.config.obj_excel}\\{self.config.obj_sheet}]数据成功，新文件已保存！ giftid={gift.id}")
            
        except FileNotFoundError:
            logger.error(f"错误：[write_objexcel] 文件 {self.config.obj_excel}\\{self.config.obj_sheet}  不存在,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        except KeyError:
            logger.error(f"错误：[write_objexcel] 工作表 {self.config.obj_excel}\\{self.config.obj_sheet} 不存在,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        except Exception as e:
            logger.error(f"错误：{e}, [write_objexcel] 工作表 {self.config.obj_excel}\\{self.config.obj_sheet}  不可用,  {traceback.format_exc()}")
            #input("Press Enter to continue...")

        
            
    def getColIndexByColumnName(self, column_name):
        for col_index, header in self.headers.items():
            if header == column_name:
                return col_index;
        
    def fill_gift_row_new(self, gift):
        for index, value in enumerate(gift.columns):
            if value == None or not value:
                continue;

            col_index = self.getColIndexByColumnName(value);
            if col_index == None or col_index < 0 or col_index >= len(gift.row) or index >= len(gift.newdata):
                #print(f"error: getColIndexByColumnName failed, col_index: {col_index}, value: {value}");
                continue;
            
            gift.row[col_index] = gift.newdata[index];
    
    def process_excel(self, gift):
        #logger.debug(f"process_excel: {gift.id}");
        data = self.getTemplateData(gift.templateID, gift.id);
        if data == None:
            #logger.debug("error: getTemplateData failed");
            return;

        max_id = data[0];
        teamplatedata = data[1];
        olddata = data[2];
    
        if gift.id == 0:    # 需要产出新道具
            if teamplatedata == None or len(teamplatedata) == 0:
                #logger.debug("error: teamplatedata is None");
                return;
        
            gift.id = max_id + 1;
            gift.row = cp.copy(teamplatedata);
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
                    logger.debug("error: teamplatedata is None");
                    return;
                gift.row = cp.copy(teamplatedata);
                gift.row[0] = gift.id;
                self.fill_gift_row_new(gift);
                self.write_objexcel(gift);
    
    def read_gifts(self):
        try: 
            logger.debug(f"read gift template begin {self.config.gift_excel}\\{self.config.gift_sheet}");
            # 打开 Excel 文件
            workbook = xlrd.open_workbook(self.config.gift_excel)
            # 或者通过表名获取工作表
            sheet = workbook.sheet_by_name('Sheet1')

            # 获取工作表的行数和列数
            rows = sheet.nrows
            columns = sheet.ncols

            gift = NewGift(0);

            for row in range(rows):
                row_values = sheet.row_values(row)
                #print(f'第 {row + 1} 行的数据为: {row_values}')

                if row_values[0] == None or not row_values[0]:
                    continue;
                if  row_values[0] == 'id':
                    if len(gift.field) == 0:
                        gift.field = row_values;
                elif row_values[0] == '礼包' and row_values[1] == '新道具':
                    #print("new gift: 需要创建新的道具");
                    self.emplace_gift(gift)
                    gift = NewGift(0);
                elif row_values[0] == '礼包' and row_values[1] != None and int(row_values[1] > 0):
                    #print("new gift: 需要覆盖老的道具");
                    self.emplace_gift(gift)
                    gift = NewGift(int(row_values[1]));
                    continue;
                elif row_values[0] == 'templateID':
                    gift.columns = row_values;
                else:
                    if len(gift.columns) != 0 and len(gift.newdata) == 0 and gift.templateID == 0:
                        gift.newdata = row_values;
                        gift.templateID = int(row_values[0]);
                    else:
                        gift.objs.append(row_values);
        
            self.emplace_gift(gift);
            logger.debug(f"read gift template success, gift count: {len(self.gifts)}");
        except FileNotFoundError:
            logger.error(f"错误：[read_gifts] 文件 {self.config.gift_excel}\\{self.config.gift_sheet} 不存在,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return []
        except KeyError:
            logger.error(f"错误：[read_gifts] 工作表 {self.config.gift_excel}\\{self.config.gift_sheet} 不存在,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return []
        except Exception as e:
            logger.error(f"错误：{e}, [read_gifts] 工作表 {self.config.gift_excel}\\{self.config.gift_sheet} 不可用,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return []
        
    def read_gift_template(self):
        try:        
            logger.debug(f"read gift template begin {self.config.gift_template_xml}");

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
            
            logger.debug(f"read gift template success");
        except FileNotFoundError:
            logger.error(f"错误：[read_gift_template] 文件 {self.config.gift_template_xml} 不存在！,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return
        except KeyError:
            logger.error(f"错误：[read_gift_template] 工作表 {self.config.gift_template_xml} 不存在,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return
        except Exception as e:
            logger.error(f"错误：{e}, [read_gift_template] 工作表 {self.config.gift_template_xml} 不可用,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return
        
    def generate_giftxml(self):
        logger.debug(f"start generate xml {self.config.gift_output_xml}");
        try:
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
                        if value == None or value=='':
                            continue;
                        
                        key = gift.field[col_index];
                        if key == None or not key:
                            continue;
                        
                        if key == 'name':                            
                            pattern = re.compile(r'[\u00B7]')
                            new_str = pattern.sub('-', value)

                            record.set(key, pattern.sub('-', value));
                        else:
                            if value.is_integer():
                                value = int(value);
                            
                            record.set(key, str(value));
            indent(root)
            # 创建 XML 树
            tree = ET.ElementTree(root)
            # 写入 XML 文件
            tree.write(self.config.gift_output_xml, encoding='GB2312', xml_declaration=True)

            logger.debug(f"generate xml end {self.config.gift_output_xml}");
        except Exception as e:
            logger.error(f"错误：{e}, [generate_giftxml] 工作表 {self.config.gift_output_xml} 不可用:  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return
    
    def append_giftxml(self, id):
        try:
            with open(self.config.gift_output_xml, "r", encoding="GB2312") as file:
                content = file.read()

            root = ET.fromstring(content, parser=ET.XMLParser(encoding="GB2312"))

            target_node = None
            for child in root.findall('bag'):
                if child.tag != "bag":
                    continue;
                
                if int(child.attrib["id"]) == id:
                    target_node = child;
                    break;

            if target_node == None:
                logger.error(f"错误：[append_giftxml] 工作表 {self.config.gift_output_xml} target_node 节点:{id} 不存在,  {traceback.format_exc()}")
                return;

            index = list(root).index(target_node)

            for gift in self.gifts:
                bagroot = ET.Element('bag')
                root.insert(index + 1, bagroot)
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
                        if value == None or value=='':
                            continue;
                        
                        key = gift.field[col_index];
                        if key == None or not key:
                            continue;
                        
                        if key == 'name':                            
                            pattern = re.compile(r'[\u00B7]')
                            new_str = pattern.sub('-', value)

                            record.set(key, pattern.sub('-', value));
                        else:
                            if value.is_integer():
                                value = int(value);
                            
                            record.set(key, str(value));

            indent(root)
            # 创建 XML 树
            tree = ET.ElementTree(root)
            # 写入 XML 文件
            tree.write(self.config.gift_output_xml, encoding='GB2312', xml_declaration=True)

        except Exception as e:
            logger.error(f"错误：{e}, [generate_giftxml] 工作表 {self.config.gift_output_xml} 不可用:  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return
        
    def append_giftxml_atlast_et(self):
        try:
            with open(self.config.gift_output_xml, "r", encoding="GB2312") as file:
                content = file.read()

            # 第一种方式-无法处理注释
            # root = ET.fromstring(content, parser=ET.XMLParser(encoding="GB2312"))

            # 第二种方式， _parser 报错
            # 将 XML 内容转换为字节
            # xml_bytes = content.encode('GB2312')
            # 使用自定义解析器解析 XML
            # root = ET.fromstring(xml_bytes, parser=CommentParser())

            myparser = etree.XMLParser(remove_comments=False, encoding='GB2312')
            root = etree.fromstring(content.encode('GB2312'), parser=myparser)

            # 找到最后一个节点
            items = root.findall('bag')
            if len(items) > 0:
                last_node = items[-1]
            else:
                last_node = None

            for gift in self.gifts: 
                bagroot = ET.Element('bag')
                if last_node == None:
                    root.append(bagroot)
                else:
                    root.insert(list(root).index(last_node) + 1, bagroot)

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
                        if value == None or value=='':
                            continue;

                        key = gift.field[col_index];
                        if key == None or not key:
                            continue;

                        if key == 'name':
                            pattern = re.compile(r'[\u00B7]')
                            new_str = pattern.sub('-', value)

                            record.set(key, pattern.sub('-', value));
                        else:
                            if value.is_integer():
                                value = int(value);

                            record.set(key, str(value));

            indent(root)
            # 创建 XML 树
            tree = ET.ElementTree(root)
            # 写入 XML 文
            tree.write(self.config.gift_output_xml, encoding='GB2312', xml_declaration=True)
    
        except Exception as e:
            logger.error(f"错误：{e}, [generate_giftxml] 工作表 {self.config.gift_output_xml} 不可用:  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return
        
    def append_giftxml_atlast_lxml(self):
        try:
            myparser = etree.XMLParser(remove_comments=False, encoding='GB2312')
            tree = etree.parse(self.config.gift_output_xml, parser=myparser)

            root = tree.getroot()

            # 找到最后一个节点
            items = root.findall('bag')
            if len(items) > 0:
                last_node = items[-1]
            else:
                last_node = None

            for gift in self.gifts: 
                bagroot = etree.Element('bag')
                if last_node == None:
                    root.append(bagroot)
                else:
                    last_node.addnext(bagroot)

                for key, value in self.bagattr.items():
                    if key == "id":
                        bagroot.set("id", str(gift.id));
                    else:
                        bagroot.set(key, str(value));

                rulenode = etree.SubElement(bagroot, 'rule')
                rulenode.set("type", str(self.rulletype));

                for row in gift.objs:
                    record = etree.SubElement(bagroot, 'item')
                    for col_index, value in enumerate(row):
                        if value == None or value=='':
                            continue;

                        key = gift.field[col_index];
                        if key == None or not key:
                            continue;

                        if key == 'name':
                            pattern = re.compile(r'[\u00B7]')
                            new_str = pattern.sub('-', value)

                            record.set(key, pattern.sub('-', value));
                        else:
                            if value.is_integer():
                                value = int(value);

                            record.set(key, str(value));

            indent(root)
            # 写入 XML 文
            tree.write(self.config.gift_output_xml, encoding='GB2312', pretty_print=True, xml_declaration=True)
    
        except Exception as e:
            logger.error(f"错误：{e}, [generate_giftxml] 工作表 {self.config.gift_output_xml} 不可用:  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return
        
    def giftBagNodeTostr(self):
        tostr = "";
        for gift in self.gifts: 
            bagroot = ET.Element('bag')
            for key, value in self.bagattr.items():
                if key == "id":
                    bagroot.set("id", str(gift.id));
                else:
                    bagroot.set(key, str(value));

            rulenode = ET.SubElement(bagroot, 'rule')
            rulenode.set("type", str(self.rulletype));

            default_item = {"sex": "0", "needSpace": "1"};

            for row in gift.objs:
                newkeys = []
                record = ET.SubElement(bagroot, 'item')
                for col_index, value in enumerate(row):
                    if value == None or value=='':
                        continue;

                    key = gift.field[col_index];
                    if key == None or not key:
                        continue;

                    if key == 'name':
                        pattern = re.compile(r'[\u00B7]')
                        new_str = pattern.sub('-', value)

                        record.set(key, pattern.sub('-', value));
                        newkeys.append(key);
                    else:
                        if value.is_integer():
                            value = int(value);

                        record.set(key, str(value));
                        newkeys.append(key);

                for key in default_item.keys():
                    if key not in newkeys:
                        record.set(key, str(default_item[key]));

            indent(bagroot, 2)
            tostr += ET.tostring(bagroot,encoding='GB2312', xml_declaration=False).decode('GB2312')
        return tostr;
        
    def append_giftxml_atlast(self):
        try:
            targ_str = "</giftbagconfig>";
            with open(self.config.gift_output_xml, "r", encoding="GB2312") as file:
                lines  = file.readlines()
                linenum = 0
                for line in lines:
                    linenum += 1;
                    str= line.strip();
                    if str == targ_str:
                        #print("linenum={0}, line={1}".format(linenum, str));
                        giftstr = self.giftBagNodeTostr();
                        lines[linenum - 1] = "\t" + giftstr + "\n</giftbagconfig>";
                        break;

            with open(self.config.gift_output_xml, "w", encoding="GB2312", newline="\n") as file:
                file.writelines(lines)
                    
        except Exception as e:
            logger.error(f"错误：{e}, [generate_giftxml] 工作表 {self.config.gift_output_xml} 不可用:  {traceback.format_exc()}")
        
    def process_all_gifts(self):        
        for gift in self.gifts:
            self.process_excel(gift);
        
        logger.info("[%s][%s] [excel]更新完成" %(self.config.obj_excel, self.config.obj_sheet))

        self.append_giftxml_atlast();
        logger.info("[%s][xml]更新完成" %(self.config.gift_output_xml))

    def runPacktools(self):
        try:
            for packtool in self.config.packtool:
                logger.info(f"runPacktools: {packtool.exec_path} {packtool.input_path} {packtool.output_path} {packtool.arg}");
                result = subprocess.run(
                    [packtool.exec_path, packtool.input_path, packtool.output_path, packtool.arg],
                    check=False  # 若进程退出码非零，会抛出异常
                )
            logger.info("成完毕，别忘记提交IXML文件");
        except subprocess.CalledProcessError as e:
            logger.error(f"执行失败: {e}")
        except Exception as e:
            logger.error(f"错误：{e}, [runPacktools] 失败:  {traceback.format_exc()}")

    def run(self):
        try:
            self.read_excel_headers();

            self.read_gifts();
            self.read_gift_template();
            self.process_all_gifts();
            self.runPacktools();
        except Exception as e:
            logger.error(f"错误：{e}, [run] 失败: {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return None


if __name__=="__main__":
    initlogger();
    logger.info("start");
    #base_path = Path(__file__).parent/"config.xml"
    base_path = './config.xml';
    config = read_config(base_path);

    if len(sys.argv) == 2:
        # 处理拖放的多个文件
        file = sys.argv[1];
        config.gift_excel = file;
        logger.info(f"config.gift_excel: {config.gift_excel}");
        print(f"config.gift_excel: {config.gift_excel}");
        #input("Press Enter to continue...")
    
    generator = ConfigGenerator(config);
    generator.run()
    logger.info("end");
    input("Press Enter to continue...")
    

    