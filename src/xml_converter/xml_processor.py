from lxml import etree as ET

# XML namespace
NAMESPACES = {
    'ns': 'http://pcstd.pcc.gov.tw/2003/eTender'
}

class XMLProcessor:
    @staticmethod
    def get_text_from_node(node, xpath):
        """從節點取得文字內容"""
        found_node = node.find(xpath, NAMESPACES)
        return found_node.text.strip() if found_node is not None and found_node.text else ""

    @staticmethod
    def get_element_depth(element):
        """取得節點的深度"""
        depth = 0
        parent = element.getparent()
        while parent is not None:
            if parent.tag.endswith('}PayItem') or parent.tag.endswith('}WorkItem'):
                depth += 1
            parent = parent.getparent()
        return depth

    @staticmethod
    def get_separator(item_kind, description, item_no):
        """根據 itemKind 取得分隔符號"""
        if item_kind == "mainItem":
            return "M"
        elif item_kind == "subtotal":
            return "S"
        elif "比例項目名稱" in description:
            return "%"
        return ""

    @classmethod
    def process_xml_node(cls, node, node_string):
        """處理 XML 節點"""
        if node_string == "WorkItem":
            item_code = node.attrib.get("itemCode", "")
            item_no = node.attrib.get("itemNo", "")
        else:
            item_code = node.attrib.get("refItemCode", "")
            item_no = node.attrib.get("itemNo", "")
        
        item_kind = node.attrib.get("itemKind", "")
        description = cls.get_text_from_node(node, ".//ns:Description[@language='zh-TW']")
        unit = cls.get_text_from_node(node, ".//ns:Unit[@language='zh-TW']")
        quantity = cls.get_text_from_node(node, ".//ns:Quantity")
        price = cls.get_text_from_node(node, ".//ns:Price")
        amount = cls.get_text_from_node(node, ".//ns:Amount")
        depth = cls.get_element_depth(node)
        sep = cls.get_separator(item_kind, description, item_no)

        if not (item_no if 'item_no' in locals() else '').startswith("VAR"):
            return {
                '項目代碼': item_code,
                '項目種類': item_kind,
                '項次': item_no,
                '說明': description,
                '單位': unit,
                '數量': quantity,
                '單價': price,
                '金額': amount,
                '階層': depth,
                '分隔符號': sep
            }
        return None

    @classmethod
    def process_xml_file(cls, file_path, item_type="PayItem"):
        """處理 XML 檔案並返回所有項目的資料"""
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        xpath = ".//ns:PayItem" if item_type == "PayItem" else ".//ns:CostBreakdownList/ns:WorkItem"
        items = root.findall(xpath, NAMESPACES)
        
        data = []
        for item in items:
            item_data = cls.process_xml_node(item, item_type)
            if item_data:
                data.append(item_data)
        
        return data

    @classmethod
    def process_cost_breakdown_tree(cls, file_path):
        """處理 CostBreakdownList 並返回樹狀結構資料"""
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        cost_breakdown = root.find(".//ns:CostBreakdownList", NAMESPACES)
        if cost_breakdown is None:
            return []
        
        def process_work_item(node, parent_id=""):
            item_code = node.attrib.get("itemCode", "")
            ref_item_no = node.attrib.get("refItemNo", "")
            item_kind = node.attrib.get("itemKind", "")
            output_quantity = node.attrib.get("analysisOutputQuantity", "")
            
            description = cls.get_text_from_node(node, ".//ns:Description[@language='zh-TW']")
            unit = cls.get_text_from_node(node, ".//ns:Unit[@language='zh-TW']")
            quantity = cls.get_text_from_node(node, ".//ns:Quantity")
            price = cls.get_text_from_node(node, ".//ns:Price")
            amount = cls.get_text_from_node(node, ".//ns:Amount")
            
            current_id = f"{parent_id}/{item_code}" if parent_id else item_code
            
            result = {
                'id': current_id,
                'parent': parent_id,
                'item_code': item_code,
                'ref_item_no': ref_item_no,
                'item_kind': item_kind,
                'description': description,
                'unit': unit,
                'quantity': quantity,
                'price': price,
                'amount': amount,
                'output_quantity': output_quantity,
                'children': []
            }
            
            # 處理子項目
            for child in node.findall(".//ns:WorkItem", NAMESPACES):
                if child.getparent() == node:  # 只處理直接子項目
                    child_data = process_work_item(child, current_id)
                    result['children'].append(child_data)
            
            return result
        
        # 處理頂層 WorkItem
        result = []
        for work_item in cost_breakdown.findall("./ns:WorkItem", NAMESPACES):
            item_data = process_work_item(work_item)
            result.append(item_data)
        
        return result
