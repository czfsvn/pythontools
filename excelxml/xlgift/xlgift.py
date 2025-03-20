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

# �������ͳ���
XL_CELL_EMPTY = 0
XL_CELL_TEXT = 1
XL_CELL_NUMBER = 2
XL_CELL_DATE = 3
XL_CELL_BOOLEAN = 4
XL_CELL_ERROR = 5

# ������־��¼��
logger = logging.getLogger(__name__)

def custom_exception_handler(exc_type, exc_value, exc_traceback):
    # ��ȡ���������к�
    line_number = exc_traceback.tb_lineno
    # ��ȡ���������ļ���
    file_name = inspect.getframeinfo(exc_traceback.tb_frame).filename
    logger.critical(f"���ļ� {file_name} �ĵ� {line_number} �з����� {exc_type.__name__} �쳣: {exc_value}")

# �����Զ����쳣������
sys.excepthook = custom_exception_handler

# �Զ��������������ע��
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
        self.packtool = []; # �����������

class NewGift:
    def __init__(self, id):
        self.id = id;
        self.field = [];   # gift���ֶ�
        self.objs = []; 
        self.row = [];
        self.row = [];
        self.xml = None
        self.templateID = 0;
        self.columns = []  # �洢��Ҫ�޸ĵ�������
        self.newdata = [];  # �洢��Ҫ�޸ĵ�������

def initlogger():
    logger.setLevel(logging.DEBUG)

    # �����ļ�������
    file_handler = logging.FileHandler('./app.log')
    file_handler.setLevel(logging.DEBUG)

    # ��������̨������
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    # ������־��ʽ
    formatter = logging.Formatter('%(asctime)s [%(levelname)s][%(filename)s:%(lineno)d] %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # ����������ӵ���־��¼��
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
        logger.error(f"����[read_config] �ļ� {configpath} ������,  {traceback.format_exc()}")
        #input("Press Enter to continue...")
        return None
    except KeyError:
        logger.error(f"����[read_config] ������ {configpath} ������,  {traceback.format_exc()}")
        #input("Press Enter to continue...")
        return None
    except Exception as e:
        logger.error(f"����{e}, [read_config] ������[read_config] {configpath} ������,  {traceback.format_exc()}")
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
            # �� Excel �ļ�
            workbook = xlrd.open_workbook(self.config.obj_excel, formatting_info=True)

            # ����ͨ��������ȡ������
            sheet = workbook.sheet_by_name(self.config.obj_sheet)

            # ��ȡ�����������������
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
            logger.error(f"����[read_excel_headers] �ļ� {self.config.obj_excel}\\{self.config.obj_sheet}  ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        except KeyError:
            logger.error(f"����[read_excel_headers] ������ {self.config.obj_excel}\\{self.config.obj_sheet} ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        except Exception as e:
            logger.error(f"����{e}, [read_excel_headers] ������ {self.config.obj_excel}\\{self.config.obj_sheet} ������,  {traceback.format_exc()}")
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
            logger.error(f"����[getTemplateData] �ļ� {self.config.obj_excel}\\{self.config.obj_sheet} ������,  {traceback.format_exc()}")
            #("Press Enter to continue...")
            return []
        except KeyError:
            logger.error(f"����[getTemplateData] ������ {self.config.obj_excel}\\{self.config.obj_sheet} ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return []
        except Exception as e:
            logger.error(f"����{e}, [getTemplateData] ������ {self.config.obj_excel}\\{self.config.obj_sheet} ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return []
        
    def update_old(self, gift):
        try:
            # �����е� Excel �ļ�
            old_workbook = xlrd.open_workbook(self.config.obj_excel, formatting_info=True)
            # ��ȡ��һ��������
            old_sheet = old_workbook.sheet_by_name(self.config.obj_sheet)
        except FileNotFoundError:
            logger.error(f"����[update_old] �ļ� {self.config.obj_excel}\\{self.config.obj_sheet}  ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        except KeyError:
            logger.error(f"����[update_old] ������ {self.config.obj_excel}\\{self.config.obj_sheet} ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        except Exception as e:
            logger.error(f"����{e}, [update_old] ������ {self.config.obj_excel}\\{self.config.obj_sheet}  ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        
        try:
            # ����ͨ��������ȡ������
            # ʹ�� xlutils.copy ���ƹ�����Ϊ��д�ĸ���
            wb_copy = copy(old_workbook)

            # ��ȡԭ�������иù����������
            sheet_index = old_workbook.sheet_names().index(self.config.obj_sheet)
            # ����ͨ��������ȡ������
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
                logger.debug(f"�޸�[{self.config.obj_excel}\\{self.config.obj_sheet}]���ݳɹ������ļ��ѱ��棡 giftid={gift.id}")


        except Exception as e:
            logger.error(f"����{e}, [update_old] ������ {self.config.obj_excel}\\{self.config.obj_sheet}  ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        
    def write_objexcel(self, gift):
        try:
            # �����е� Excel �ļ�
            old_workbook = xlrd.open_workbook(self.config.obj_excel)
            # ��ȡ��һ��������
            old_sheet = old_workbook.sheet_by_name(self.config.obj_sheet)
            # ��ȡ�������ݵ�����
            rows = old_sheet.nrows

            # ����һ���µĹ�����
            new_workbook = xlwt.Workbook()
            # ���¹����������һ��������
            new_sheet = new_workbook.add_sheet(self.config.obj_sheet)
            for row in range(rows):
                for col in range(old_sheet.ncols):
                    new_sheet.write(row, col, old_sheet.cell_value(row, col))

            # ׷��������
            for col_index, value in enumerate(gift.row):
                new_sheet.write(rows, col_index, value);
                
            logger.debug(f"write_excel: {gift.id}, {gift.row}");

            # �����޸ĺ�� Excel �ļ�
            # �����µĹ����������ļ�·��
            new_workbook.save(self.config.obj_excel)
            logger.debug(f"����[{self.config.obj_excel}\\{self.config.obj_sheet}]���ݳɹ������ļ��ѱ��棡 giftid={gift.id}")
            
        except FileNotFoundError:
            logger.error(f"����[write_objexcel] �ļ� {self.config.obj_excel}\\{self.config.obj_sheet}  ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        except KeyError:
            logger.error(f"����[write_objexcel] ������ {self.config.obj_excel}\\{self.config.obj_sheet} ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
        except Exception as e:
            logger.error(f"����{e}, [write_objexcel] ������ {self.config.obj_excel}\\{self.config.obj_sheet}  ������,  {traceback.format_exc()}")
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
    
        if gift.id == 0:    # ��Ҫ�����µ���
            if teamplatedata == None or len(teamplatedata) == 0:
                #logger.debug("error: teamplatedata is None");
                return;
        
            gift.id = max_id + 1;
            gift.row = cp.copy(teamplatedata);
            gift.row[0] = gift.id;
            self.fill_gift_row_new(gift);
            self.write_objexcel(gift);

        elif gift.id > 0:
            if olddata != None and len(olddata) > 0: # ��Ҫ�����ϵ�����Ϣ, ��ʹ��gift.id
                # todo: ���Ǿ�����
                self.update_old(gift);
                return;            
            else:   # ��Ҫ�����µ�����Ϣ�� ��ʹ��gift.id
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
            # �� Excel �ļ�
            workbook = xlrd.open_workbook(self.config.gift_excel)
            # ����ͨ��������ȡ������
            sheet = workbook.sheet_by_name('Sheet1')

            # ��ȡ�����������������
            rows = sheet.nrows
            columns = sheet.ncols

            gift = NewGift(0);

            for row in range(rows):
                row_values = sheet.row_values(row)
                #print(f'�� {row + 1} �е�����Ϊ: {row_values}')

                if row_values[0] == None or not row_values[0]:
                    continue;
                if  row_values[0] == 'id':
                    if len(gift.field) == 0:
                        gift.field = row_values;
                elif row_values[0] == '���' and row_values[1] == '�µ���':
                    #print("new gift: ��Ҫ�����µĵ���");
                    self.emplace_gift(gift)
                    gift = NewGift(0);
                elif row_values[0] == '���' and row_values[1] != None and int(row_values[1] > 0):
                    #print("new gift: ��Ҫ�����ϵĵ���");
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
            logger.error(f"����[read_gifts] �ļ� {self.config.gift_excel}\\{self.config.gift_sheet} ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return []
        except KeyError:
            logger.error(f"����[read_gifts] ������ {self.config.gift_excel}\\{self.config.gift_sheet} ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return []
        except Exception as e:
            logger.error(f"����{e}, [read_gifts] ������ {self.config.gift_excel}\\{self.config.gift_sheet} ������,  {traceback.format_exc()}")
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
            logger.error(f"����[read_gift_template] �ļ� {self.config.gift_template_xml} �����ڣ�,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return
        except KeyError:
            logger.error(f"����[read_gift_template] ������ {self.config.gift_template_xml} ������,  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return
        except Exception as e:
            logger.error(f"����{e}, [read_gift_template] ������ {self.config.gift_template_xml} ������,  {traceback.format_exc()}")
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
            # ���� XML ��
            tree = ET.ElementTree(root)
            # д�� XML �ļ�
            tree.write(self.config.gift_output_xml, encoding='GB2312', xml_declaration=True)

            logger.debug(f"generate xml end {self.config.gift_output_xml}");
        except Exception as e:
            logger.error(f"����{e}, [generate_giftxml] ������ {self.config.gift_output_xml} ������:  {traceback.format_exc()}")
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
                logger.error(f"����[append_giftxml] ������ {self.config.gift_output_xml} target_node �ڵ�:{id} ������,  {traceback.format_exc()}")
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
            # ���� XML ��
            tree = ET.ElementTree(root)
            # д�� XML �ļ�
            tree.write(self.config.gift_output_xml, encoding='GB2312', xml_declaration=True)

        except Exception as e:
            logger.error(f"����{e}, [generate_giftxml] ������ {self.config.gift_output_xml} ������:  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return
        
    def append_giftxml_atlast_et(self):
        try:
            with open(self.config.gift_output_xml, "r", encoding="GB2312") as file:
                content = file.read()

            # ��һ�ַ�ʽ-�޷�����ע��
            # root = ET.fromstring(content, parser=ET.XMLParser(encoding="GB2312"))

            # �ڶ��ַ�ʽ�� _parser ����
            # �� XML ����ת��Ϊ�ֽ�
            # xml_bytes = content.encode('GB2312')
            # ʹ���Զ������������ XML
            # root = ET.fromstring(xml_bytes, parser=CommentParser())

            myparser = etree.XMLParser(remove_comments=False, encoding='GB2312')
            root = etree.fromstring(content.encode('GB2312'), parser=myparser)

            # �ҵ����һ���ڵ�
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
            # ���� XML ��
            tree = ET.ElementTree(root)
            # д�� XML ��
            tree.write(self.config.gift_output_xml, encoding='GB2312', xml_declaration=True)
    
        except Exception as e:
            logger.error(f"����{e}, [generate_giftxml] ������ {self.config.gift_output_xml} ������:  {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return
        
    def append_giftxml_atlast_lxml(self):
        try:
            myparser = etree.XMLParser(remove_comments=False, encoding='GB2312')
            tree = etree.parse(self.config.gift_output_xml, parser=myparser)

            root = tree.getroot()

            # �ҵ����һ���ڵ�
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
            # д�� XML ��
            tree.write(self.config.gift_output_xml, encoding='GB2312', pretty_print=True, xml_declaration=True)
    
        except Exception as e:
            logger.error(f"����{e}, [generate_giftxml] ������ {self.config.gift_output_xml} ������:  {traceback.format_exc()}")
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
            logger.error(f"����{e}, [generate_giftxml] ������ {self.config.gift_output_xml} ������:  {traceback.format_exc()}")
        
    def process_all_gifts(self):        
        for gift in self.gifts:
            self.process_excel(gift);
        
        logger.info("[%s][%s] [excel]�������" %(self.config.obj_excel, self.config.obj_sheet))

        self.append_giftxml_atlast();
        logger.info("[%s][xml]�������" %(self.config.gift_output_xml))

    def runPacktools(self):
        try:
            for packtool in self.config.packtool:
                logger.info(f"runPacktools: {packtool.exec_path} {packtool.input_path} {packtool.output_path} {packtool.arg}");
                result = subprocess.run(
                    [packtool.exec_path, packtool.input_path, packtool.output_path, packtool.arg],
                    check=False  # �������˳�����㣬���׳��쳣
                )
            logger.info("����ϣ��������ύIXML�ļ�");
        except subprocess.CalledProcessError as e:
            logger.error(f"ִ��ʧ��: {e}")
        except Exception as e:
            logger.error(f"����{e}, [runPacktools] ʧ��:  {traceback.format_exc()}")

    def run(self):
        try:
            self.read_excel_headers();

            self.read_gifts();
            self.read_gift_template();
            self.process_all_gifts();
            self.runPacktools();
        except Exception as e:
            logger.error(f"����{e}, [run] ʧ��: {traceback.format_exc()}")
            #input("Press Enter to continue...")
            return None


if __name__=="__main__":
    initlogger();
    logger.info("start");
    #base_path = Path(__file__).parent/"config.xml"
    base_path = './config.xml';
    config = read_config(base_path);

    if len(sys.argv) == 2:
        # �����ϷŵĶ���ļ�
        file = sys.argv[1];
        config.gift_excel = file;
        logger.info(f"config.gift_excel: {config.gift_excel}");
        print(f"config.gift_excel: {config.gift_excel}");
        #input("Press Enter to continue...")
    
    generator = ConfigGenerator(config);
    generator.run()
    logger.info("end");
    input("Press Enter to continue...")
    

    