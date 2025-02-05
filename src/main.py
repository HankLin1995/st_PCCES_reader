from xml_converter.xml_processor import XMLProcessor
import pandas as pd
import streamlit as st
import os
import io

def process_items(xml_file_path, item_type):
    """處理 XML 檔案中的項目"""
    data = XMLProcessor.process_xml_file(xml_file_path, item_type)
    if not data:
        return None
    
    # 使用 pandas 處理數據
    df = pd.DataFrame(data)
    return df

def main():
    st.title("PCCES XML 檔案處理器")
    
    uploaded_file = st.file_uploader("請選擇 XML 檔案", type=['xml'])
    
    if uploaded_file is not None:
        # 儲存上傳的檔案
        with open("temp.xml", "wb") as f:
            f.write(uploaded_file.getvalue())
        
        # 處理 PayItems
        st.subheader("Pay Items")
        pay_items_df = process_items("temp.xml", "PayItem")
        if pay_items_df is not None:
            st.dataframe(pay_items_df)
            
            # 將 DataFrame 轉換為 Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                pay_items_df.to_excel(writer, sheet_name='PayItems', index=False)
            
            # 提供下載按鈕
            st.download_button(
                label="下載 PayItems Excel",
                data=output.getvalue(),
                file_name="PayItems.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # 處理 WorkItems
        st.subheader("Work Items")
        work_items_df = process_items("temp.xml", "WorkItem")
        if work_items_df is not None:
            st.dataframe(work_items_df)
            
            # 將 DataFrame 轉換為 Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                work_items_df.to_excel(writer, sheet_name='WorkItems', index=False)
            
            # 提供下載按鈕
            st.download_button(
                label="下載 WorkItems Excel",
                data=output.getvalue(),
                file_name="WorkItems.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        # 清理臨時檔案
        os.remove("temp.xml")

if __name__ == "__main__":
    main()
