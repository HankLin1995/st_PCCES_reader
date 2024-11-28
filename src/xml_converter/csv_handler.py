import csv
import pandas as pd
import base64

class CSVHandler:
    # CSV檔案的標題
    HEADERS = ['項目代碼', '項目種類', '項次', '說明', '階層', '單位', '數量', '單價', '金額', '分隔符號']
    
    # 工作表名稱
    SHTNAME_PAYITEM = "PayItemSheet"
    SHTNAME_WORKITEM = "WorkItemSheet"

    @classmethod
    def initialize_csv_files(cls):
        """初始化 CSV 檔案"""
        # 建立並初始化PayItem的CSV
        cls._create_csv_file(cls.SHTNAME_PAYITEM)
        # 建立並初始化WorkItem的CSV
        cls._create_csv_file(cls.SHTNAME_WORKITEM)

    @classmethod
    def _create_csv_file(cls, sheet_name):
        """建立並初始化 CSV 檔案"""
        with open(f"{sheet_name}.csv", 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(cls.HEADERS)

    @staticmethod
    def append_data(sheet_name, data_array):
        """添加資料到 CSV 檔案"""
        csv_filename = f"{sheet_name}.csv"
        with open(csv_filename, 'a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(data_array)

    @staticmethod
    def data_to_csv(data):
        """將資料轉換為 CSV 格式"""
        df = pd.DataFrame(data)
        return df.to_csv(index=False, encoding='utf-8-sig')

    @staticmethod
    def get_csv_download_link(df, filename):
        """生成 CSV 下載連結"""
        csv = df.to_csv(index=False, encoding='utf-8-sig')
        b64 = base64.b64encode(csv.encode()).decode()
        return f'<a href="data:file/csv;base64,{b64}" download="{filename}">Download {filename}</a>'
