from lxml import etree
import pandas as pd
from datetime import datetime

class ExcelToXMLConverter:
    @staticmethod
    def convert_excel_to_xml(excel_file):
        # 讀取 Excel 檔案
        df = pd.read_excel(excel_file)
        
        # 定義命名空間
        nsmap = {
            None: "http://pcstd.pcc.gov.tw/2003/eTender",
            "xsi": "http://www.w3.org/2001/XMLSchema-instance",
            "mml": "http://www.w3.org/1998/Math/MathML"
        }
        
        # 創建 XML 根元素
        root = etree.Element("ETenderSheet", nsmap=nsmap)
        root.set("{http://www.w3.org/2001/XMLSchema-instance}schemaLocation",
                 "http://pcstd.pcc.gov.tw/2003/eTender eTender.1.J.2004.D.xsd")
        root.set("createdDate", datetime.now().strftime("%Y-%m-%d"))
        root.set("applicationName", "Pcces")
        root.set("applicationVersion", "4.3.1000.220")
        root.set("procurementType", "constructionWork")
        root.set("documentType", "contract")
        
        # 添加 TenderInformation
        tender_info = etree.SubElement(root, "TenderInformation")
        
        # 添加基本資訊
        procuring_entity_zh = etree.SubElement(tender_info, "ProcuringEntity")
        procuring_entity_zh.set("language", "zh-TW")
        procuring_entity_zh.text = "工程主辦機關"
        
        procuring_entity_en = etree.SubElement(tender_info, "ProcuringEntity")
        procuring_entity_en.set("language", "en")
        procuring_entity_en.text = " " * 60
        
        contract_title_zh = etree.SubElement(tender_info, "ContractTitle")
        contract_title_zh.set("language", "zh-TW")
        contract_title_zh.text = "工程名稱"
        
        contract_title_en = etree.SubElement(tender_info, "ContractTitle")
        contract_title_en.set("language", "en")
        contract_title_en.text = " " * 60
        
        # 添加 DetailList
        detail_list = etree.SubElement(root, "DetailList")
        detail_list.set("canAddItem", "false")
        detail_list.set("quantityDisplayDecimal", "1")
        detail_list.set("unitPriceDisplayDecimal", "0")
        detail_list.set("itemAmountDisplayDecimal", "0")
        
        # 處理每一行資料
        current_main_item = None
        for _, row in df.iterrows():
            item_no = str(row['項次'])
            
            # 判斷是否為主項目
            if len(item_no) == 1 or item_no.isalpha():  # 如果項次為單一數字或字母
                current_main_item = ExcelToXMLConverter._create_pay_item(
                    detail_list if current_main_item is None else current_main_item,
                    row, "mainItem"
                )
            else:
                # 一般項目
                ExcelToXMLConverter._create_pay_item(current_main_item, row, row['項目種類'])
        
        return root
    
    @staticmethod
    def _create_pay_item(parent, row, item_kind):
        pay_item = etree.SubElement(parent, "PayItem")
        pay_item.set("itemKey", str(row.get('項次', '')))
        pay_item.set("itemNo", str(row.get('項次', '')))
        pay_item.set("refItemCode", str(row.get('項目代碼', '')).ljust(20))
        pay_item.set("itemKind", item_kind)
        
        # 說明
        desc_zh = etree.SubElement(pay_item, "Description")
        desc_zh.set("language", "zh-TW")
        desc_zh.text = str(row.get('說明', ''))
        
        desc_en = etree.SubElement(pay_item, "Description")
        desc_en.set("language", "en")
        desc_en.text = " " * 120
        
        # 單位
        unit_zh = etree.SubElement(pay_item, "Unit")
        unit_zh.set("language", "zh-TW")
        unit_zh.text = str(row.get('單位', '')).ljust(10)
        
        unit_en = etree.SubElement(pay_item, "Unit")
        unit_en.set("language", "en")
        unit_en.text = " " * 10
        
        # 數量
        quantity = etree.SubElement(pay_item, "Quantity")
        quantity.text = str(row.get('數量', 0))
        
        # 單價
        price = etree.SubElement(pay_item, "Price")
        price.set("fixed", "false")
        price.text = str(row.get('單價', 0))
        
        # 金額
        amount = etree.SubElement(pay_item, "Amount")
        amount.set("fixed", "false")
        amount.text = str(row.get('金額', 0))
        
        # 備註
        remark = etree.SubElement(pay_item, "Remark")
        remark.text = "[發包]"
        
        # 百分比
        percent = etree.SubElement(pay_item, "Percent")
        percent.text = "0"
        
        if item_kind == "analysis":
            # 分析項目特有欄位
            labour = etree.SubElement(pay_item, "LabourRatio")
            labour.text = "0"
            
            equipment = etree.SubElement(pay_item, "EquipmentRatio")
            equipment.text = "0"
            
            material = etree.SubElement(pay_item, "MaterialRatio")
            material.text = "0"
            
            misc = etree.SubElement(pay_item, "MiscellaneaRatio")
            misc.text = "0"
        
        return pay_item
    
    @staticmethod
    def save_xml(xml_tree, output_file):
        # 設定 XML 格式化參數
        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.ElementTree(xml_tree)
        
        # 寫入檔案
        tree.write(
            output_file,
            encoding="UTF-8",
            xml_declaration=True,
            pretty_print=True,
            standalone="no"
        )
