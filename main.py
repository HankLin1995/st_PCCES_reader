from lxml import etree as ET
import csv

# 設定工作表名稱 (模擬 Excel 的處理)
SHTNAME_PAYITEM = "PayItemSheet"
SHTNAME_WORKITEM = "WorkItemSheet"

# XML namespace
NAMESPACES = {
    'ns': 'http://pcstd.pcc.gov.tw/2003/eTender'
}

# CSV檔案的標題
HEADERS = ['項目代碼', '項目種類', '項次', '說明', '階層', '單位', '數量', '單價', '金額', '分隔符號']

# 假設 append_data 和 set_interior_color_index 是處理 Excel 的函數
def append_data(sheet_name, data_array):
    csv_filename = f"{sheet_name}.csv"
    with open(csv_filename, 'a', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(data_array)

def initialize_csv_files():
    # 建立並初始化PayItem的CSV
    with open(f"{SHTNAME_PAYITEM}.csv", 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)
    
    # 建立並初始化WorkItem的CSV
    with open(f"{SHTNAME_WORKITEM}.csv", 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(HEADERS)

def set_interior_color_index():
    # 模擬設置 Excel 樣式的功能
    print("Setting interior color index...")

# 讀取 PayItem 資料
def get_pay_items(xml_file_path):
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    # 修改XPath以取得所有PayItem，包括子層級
    pay_items = root.findall(".//ns:PayItem", NAMESPACES)
    for pay_item in pay_items:
        process_xml_node(pay_item, "PayItem")
        print("*********")
    
    set_interior_color_index()

# 讀取 WorkItem 資料
def get_work_items(xml_file_path):
    tree = ET.parse(xml_file_path)
    root = tree.getroot()

    work_items = root.findall(".//ns:CostBreakdownList/ns:WorkItem", NAMESPACES)
    for work_item in work_items:
        process_xml_node(work_item, "WorkItem")
    
    set_interior_color_index()

# 處理 XML 節點
def process_xml_node(node, node_string):
    if node_string == "WorkItem":
        sheet_name = SHTNAME_WORKITEM
        item_code = node.attrib.get("itemCode", "")
        item_no = node.attrib.get("itemNo", "")  
    else:
        sheet_name = SHTNAME_PAYITEM
        item_code = node.attrib.get("refItemCode", "")
        item_no = node.attrib.get("itemNo", "")
    
    item_kind = node.attrib.get("itemKind", "")
    
    description = get_text_from_node(node, ".//ns:Description[@language='zh-TW']")
    unit = get_text_from_node(node, ".//ns:Unit[@language='zh-TW']")
    quantity = get_text_from_node(node, ".//ns:Quantity")
    price = get_text_from_node(node, ".//ns:Price")
    amount = get_text_from_node(node, ".//ns:Amount")
    
    depth = get_element_depth(node)

    sep = get_separator(item_kind, description, item_no)

    if not (item_no if 'item_no' in locals() else '').startswith("VAR"):
        data_array = [item_code, item_kind, item_no if 'item_no' in locals() else '', 
                     description, depth, unit, quantity, price, amount, sep]
        print(data_array)
        append_data(sheet_name, data_array)

# 取得節點的深度
def get_element_depth(element):
    depth = 0
    parent = element.getparent()
    while parent is not None:
        if parent.tag.endswith('}PayItem') or parent.tag.endswith('}WorkItem'):
            depth += 1
        parent = parent.getparent()
    return depth

# 從節點取得文字內容
def get_text_from_node(node, xpath):
    found_node = node.find(xpath, NAMESPACES)
    return found_node.text.strip() if found_node is not None and found_node.text else ""

# 根據 itemKind 取得分隔符號
def get_separator(item_kind, description, item_no):
    if item_kind == "mainItem":
        return "M"
    elif item_kind == "subtotal":
        return "S"
    elif "比例項目名稱" in description:
        return "%"
    return ""

# 主函數示例
if __name__ == "__main__":
    initialize_csv_files()
    xml_file_path = "example.xml"
    print("Processing Pay Items...")
    get_pay_items(xml_file_path)
    print("\nProcessing Work Items...")
    get_work_items(xml_file_path)
    print("\nCSV files have been generated:")
    print(f"1. {SHTNAME_PAYITEM}.csv")
    print(f"2. {SHTNAME_WORKITEM}.csv")
