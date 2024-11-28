from xml_processor import XMLProcessor
from csv_handler import CSVHandler

def set_interior_color_index():
    # 模擬設置 Excel 樣式的功能
    print("Setting interior color index...")

def process_items(xml_file_path, item_type):
    """處理 XML 檔案中的項目"""
    data = XMLProcessor.process_xml_file(xml_file_path, item_type)
    sheet_name = CSVHandler.SHTNAME_WORKITEM if item_type == "WorkItem" else CSVHandler.SHTNAME_PAYITEM
    
    for item in data:
        # 轉換為列表格式
        data_array = [item[key] for key in CSVHandler.HEADERS]
        print(data_array)
        CSVHandler.append_data(sheet_name, data_array)
    
    set_interior_color_index()

def main():
    CSVHandler.initialize_csv_files()
    xml_file_path = "example.xml"
    
    print("Processing Pay Items...")
    process_items(xml_file_path, "PayItem")
    
    print("Processing Work Items...")
    process_items(xml_file_path, "WorkItem")

if __name__ == "__main__":
    main()
