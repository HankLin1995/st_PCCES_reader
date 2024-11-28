import streamlit as st
from lxml import etree as ET
import pandas as pd
import io
import base64

# XML namespace
NAMESPACES = {
    'ns': 'http://pcstd.pcc.gov.tw/2003/eTender'
}

def get_text_from_node(node, xpath):
    found_node = node.find(xpath, NAMESPACES)
    return found_node.text.strip() if found_node is not None and found_node.text else ""

def get_element_depth(element):
    depth = 0
    parent = element.getparent()
    while parent is not None:
        if parent.tag.endswith('}PayItem') or parent.tag.endswith('}WorkItem'):
            depth += 1
        parent = parent.getparent()
    return depth

def get_separator(item_kind, description, item_no):
    if item_kind == "mainItem":
        return "M"
    elif item_kind == "subtotal":
        return "S"
    elif "比例項目名稱" in description:
        return "%"
    return ""

def process_xml_node(node, node_string):
    if node_string == "WorkItem":
        item_code = node.attrib.get("itemCode", "")
        item_no = node.attrib.get("itemNo", "")
    else:
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
        return {
            '項目代碼': item_code,
            '項目種類': item_kind,
            '項次': item_no,
            '說明': description,
            '階層': depth,
            '單位': unit,
            '數量': quantity,
            '單價': price,
            '金額': amount,
            '分隔符號': sep
        }
    return None

def process_xml_file(uploaded_file, item_type="PayItem"):
    tree = ET.parse(uploaded_file)
    root = tree.getroot()
    
    if item_type == "PayItem":
        items = root.findall(".//ns:PayItem", NAMESPACES)
    else:
        items = root.findall(".//ns:CostBreakdownList/ns:WorkItem", NAMESPACES)
    
    data = []
    for item in items:
        item_data = process_xml_node(item, item_type)
        if item_data:
            data.append(item_data)
    
    return pd.DataFrame(data)

def get_csv_download_link(df, filename):
    csv = df.to_csv(index=False, encoding='utf-8-sig')
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download {filename}</a>'
    return href

def main():
    st.title("XML 處理工具")
    st.write("上傳XML檔案並處理PayItem和WorkItem資料")

    uploaded_file = st.file_uploader("選擇XML檔案", type=['xml'])

    if uploaded_file is not None:
        # 創建兩個分頁
        tab1, tab2 = st.tabs(["PayItem", "WorkItem"])
        
        with tab1:
            st.header("PayItem 資料")
            pay_items_df = process_xml_file(uploaded_file, "PayItem")
            st.dataframe(pay_items_df)
            st.markdown(get_csv_download_link(pay_items_df, "PayItemSheet.csv"), unsafe_allow_html=True)
        
        with tab2:
            st.header("WorkItem 資料")
            work_items_df = process_xml_file(uploaded_file, "WorkItem")
            st.dataframe(work_items_df)
            st.markdown(get_csv_download_link(work_items_df, "WorkItemSheet.csv"), unsafe_allow_html=True)

if __name__ == "__main__":
    main()
