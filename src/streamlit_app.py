import streamlit as st
import pandas as pd
import os
import sys

# 添加項目根目錄到 Python 路徑
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

from src.xml_converter.xml_processor import XMLProcessor
from src.xml_converter.excel_to_xml import ExcelToXMLConverter

def flatten_tree_data(tree_data):
    """將樹狀結構轉換為平面列表"""
    flattened_data = []
    
    def process_node(node, level=0):
        row = {
            '項目代碼': node['item_code'],
            '參考編號': node['ref_item_no'],
            '項目種類': node['item_kind'],
            '說明': node['description'],
            '單位': node['unit'],
            '數量': node['quantity'],
            '單價': node['price'],
            '金額': node['amount'],
            '分析產出數量': node['output_quantity'],
            '階層': level
        }
        flattened_data.append(row)
        
        for child in node['children']:
            process_node(child, level + 1)
    
    for node in tree_data:
        process_node(node)
    
    return flattened_data

def process_analysis_data(tree_data):
    """處理分析項目數據，將每個分析項目分開"""
    analysis_tables = {}
    
    def process_node(node, parent_analysis=None):
        if node['item_kind'] == 'analysis':
            # 創建新的分析項目表格
            table_key = f"{node['item_code']} - {node['description']}"
            analysis_tables[table_key] = {
                '主項': {
                    '項目代碼': node['item_code'],
                    '參考編號': node['ref_item_no'],
                    '項目種類': node['item_kind'],
                    '說明': node['description'],
                    '單位': node['unit'],
                    '數量': node['quantity'],
                    '單價': node['price'],
                    '金額': node['amount'],
                    '分析產出數量': node['output_quantity']
                },
                '細項': []
            }
            current_analysis = table_key
        else:
            # 如果有父分析項目，將此項目加入到對應的細項中
            if parent_analysis and parent_analysis in analysis_tables:
                analysis_tables[parent_analysis]['細項'].append({
                    '項目代碼': node['item_code'],
                    '參考編號': node['ref_item_no'],
                    '項目種類': node['item_kind'],
                    '說明': node['description'],
                    '單位': node['unit'],
                    '數量': node['quantity'],
                    '單價': node['price'],
                    '金額': node['amount'],
                    '分析產出數量': node['output_quantity']
                })
            current_analysis = parent_analysis
        
        # 處理子節點
        for child in node['children']:
            process_node(child, current_analysis)
    
    # 處理所有頂層節點
    for node in tree_data:
        process_node(node)
    
    return analysis_tables

def process_main_items(tree_data):
    """處理總表數據"""
    main_items = []
    
    def process_node(node, step=0):
        # 檢查是否為主項目（step=0 且 item_kind 不為空）
        if step == 0 and node['item_kind']:
            main_items.append({
                '項目代碼': node['item_code'],
                '參考編號': node['ref_item_no'],
                '項目種類': node['item_kind'],
                '說明': node['description'],
                '單位': node['unit'],
                '數量': node['quantity'],
                '單價': node['price'],
                '金額': node['amount']
            })
        
        # 處理子節點，step加1
        for child in node['children']:
            process_node(child, step + 1)
    
    # 處理所有頂層節點
    for node in tree_data:
        process_node(node)
    
    return main_items

def read_excel_file(file):
    """讀取各種格式的 Excel 檔案"""
    try:
        # 嘗試使用 openpyxl 引擎
        return pd.read_excel(file, engine='openpyxl')
    except Exception as e1:
        try:
            # 如果失敗，嘗試使用 xlrd 引擎
            return pd.read_excel(file, engine='xlrd')
        except Exception as e2:
            try:
                # 最後嘗試使用 odf 引擎 (for .ods files)
                return pd.read_excel(file, engine='odf')
            except Exception as e3:
                st.error(f"無法讀取檔案，請確認檔案格式是否正確。錯誤訊息：{str(e3)}")
                return None

def get_work_item_details(tree_data, item_code):
    """獲取工作項目的細項資料"""
    details = []
    
    def find_item(node):
        if node['item_code'] == item_code:
            # 處理子項目
            for child in node['children']:
                details.append({
                    '項目代碼': child['item_code'],
                    '參考編號': child['ref_item_no'],
                    '項目種類': child['item_kind'],
                    '說明': child['description'],
                    '單位': child['unit'],
                    '數量': child['quantity'],
                    '單價': child['price'],
                    '金額': child['amount'],
                    '分析產出數量': child['output_quantity']
                })
            return True
        
        # 遞迴搜尋子節點
        for child in node['children']:
            if find_item(child):
                return True
        return False
    
    # 搜尋所有頂層節點
    for node in tree_data:
        if find_item(node):
            break
    
    return details

def main():
    st.title("XML 處理工具")
    
    # 創建兩個分頁
    tab1, tab2 = st.tabs(["XML 處理", "Excel 轉換"])
    
    with tab1:
        st.header("XML 檔案處理")
        xml_file = st.file_uploader("選擇 XML 檔案", type=['xml'], key="xml_uploader")

        if xml_file is not None:
            # 保存上傳的文件
            input_path = os.path.join('data', 'input', 'uploaded.xml')
            with open(input_path, 'wb') as f:
                f.write(xml_file.getvalue())

            # 創建三個子分頁
            subtab1, subtab2, subtab3 = st.tabs(["總表", "詳細價目表", "單價分析"])
            
            # 讀取 XML 樹狀結構（只讀取一次）
            tree_data = XMLProcessor.process_cost_breakdown_tree(input_path)
            
            pay_items_data = XMLProcessor.process_xml_file(input_path, "PayItem")

            with subtab1:
                st.header("總表")
                # if tree_data:
                    # main_items = process_main_items(tree_data)
                if pay_items_data:

                    pay_items_df=pd.DataFrame(pay_items_data)

                    main_items_df = pay_items_df[(pay_items_df['階層'] <= 1) & (pay_items_df['項目種類'].isin(['mainItem','subtotal', 'formula']))]

                    # main_items_df = pd.DataFrame(pay_items_data)
                    st.dataframe(
                        main_items_df[["項次","說明", "單位", "數量", "單價", "金額"]],  # 只顯示這幾個欄位,
                        hide_index=True,  # 隱藏索引
                        use_container_width=True  # 使用容器寬度
                    )
                    # # 準備 CSV 下載
                    # csv_path = os.path.join('data', 'output', 'MainItemSheet.csv')
                    # main_items_df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                    
                    # with open(csv_path, 'rb') as f:
                    #     st.download_button(
                    #         label="下載 CSV",
                    #         data=f,
                    #         file_name="MainItemSheet.csv",
                    #         mime="text/csv"
                    #     )
                else:
                    st.info("總表沒有數據可顯示")
                # else:
                    # st.error("無法讀取 XML 檔案中的數據")
            
            with subtab2:
                st.header("詳細價目表")
                # pay_items_data = XMLProcessor.process_xml_file(input_path, "PayItem")
                if pay_items_data:
                    pay_items_df = pd.DataFrame(pay_items_data)
                    # pay_items_df=pay_items_df[pay_items_df['階層']<=1]
                    st.dataframe(
                        pay_items_df[["項次","說明", "單位", "數量", "單價", "金額","項目代碼"]],  # 只顯示這幾個欄位,
                        hide_index=True,  # 隱藏索引
                        use_container_width=True  # 使用容器寬度
                    )
                    # 準備 CSV 下載
                    csv_path = os.path.join('data', 'output', 'DetailedPriceSheet.csv')
                    pay_items_df.to_csv(csv_path, index=False, encoding='utf-8-sig')
                    
                    with open(csv_path, 'rb') as f:
                        st.download_button(
                            label="下載 CSV",
                            data=f,
                            file_name="DetailedPriceSheet.csv",
                            mime="text/csv"
                        )
            
            with subtab3:
                st.header("單價分析")
                work_items_data = XMLProcessor.process_xml_file(input_path, "WorkItem")
                if work_items_data:
                    # 顯示工作項目表格
                    for item in work_items_data:
                        st.write(f"項目代碼: {item['項目代碼']} - {item['說明']}")
                        # if st.button("查看細項", key=f"details_{item['項目代碼']}"):
                            # 獲取細項資料
                        details = get_work_item_details(tree_data, item['項目代碼'])
                        
                        # 使用對話框顯示細項
                        # with st.expander(f"細項資料 - {item['說明']}"):
                        if details:
                            st.dataframe(pd.DataFrame(details), use_container_width=True)
                        else:
                            st.info("此項目沒有細項資料")
                        st.markdown("---")
                    
                    # 準備 CSV 下載
                    csv_path = os.path.join('data', 'output', 'UnitPriceAnalysis.csv')
                    pd.DataFrame(work_items_data).to_csv(csv_path, index=False, encoding='utf-8-sig')
                    
                    with open(csv_path, 'rb') as f:
                        st.download_button(
                            label="下載 CSV",
                            data=f,
                            file_name="UnitPriceAnalysis.csv",
                            mime="text/csv"
                        )

    with tab2:
        st.warning("Excel 轉換功能開發中...")
        # st.header("Excel 轉換為 XML")
    #     excel_file = st.file_uploader(
    #         "選擇 Excel 檔案",
    #         type=['xlsx', 'xls', 'xlsm', 'xlsb', 'odf', 'ods', 'odt'],
    #         key="excel_uploader"
    #     )
        
    #     if excel_file is not None:
    #         # 顯示 Excel 內容預覽
    #         df = read_excel_file(excel_file)
    #         if df is not None:
    #             st.subheader("Excel 內容預覽")
    #             st.dataframe(df)
                
    #             # 轉換按鈕
    #             if st.button("轉換為 XML"):
    #                 try:
    #                     # 保存上傳的 Excel 檔案
    #                     input_excel = os.path.join('data', 'input', 'temp.xlsx')
    #                     with open(input_excel, "wb") as f:
    #                         f.write(excel_file.getvalue())
                        
    #                     # 轉換為 XML
    #                     xml_tree = ExcelToXMLConverter.convert_excel_to_xml(input_excel)
    #                     output_xml = os.path.join('data', 'output', 'output.xml')
    #                     ExcelToXMLConverter.save_xml(xml_tree, output_xml)
                        
    #                     # 提供下載連結
    #                     with open(output_xml, "rb") as f:
    #                         xml_bytes = f.read()
                        
    #                     st.download_button(
    #                         label="下載 XML 檔案",
    #                         data=xml_bytes,
    #                         file_name="converted.xml",
    #                         mime="application/xml"
    #                     )
                        
    #                     # 清理臨時檔案
    #                     os.remove(input_excel)
    #                     os.remove(output_xml)
                        
    #                     st.success("轉換成功！")
    #                 except Exception as e:
    #                     st.error(f"轉換失敗：{str(e)}")
                
    #             # 顯示 Excel 格式說明
    #             st.subheader("Excel 格式說明")
    #             st.markdown("""
    #             Excel 檔案應包含以下欄位：
    #             - 項次
    #             - 項目代碼
    #             - 項目種類（預設為 analysis）
    #             - 說明
    #             - 單位
    #             - 數量
    #             - 單價
    #             - 金額
    #             """)

if __name__ == "__main__":
    main()
