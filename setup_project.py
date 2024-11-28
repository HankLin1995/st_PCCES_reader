import os
import shutil

def setup_project_structure():
    # 創建主要目錄
    directories = [
        'src',
        'src/xml_converter',
        'tests',
        'examples',
        'data',
        'data/input',
        'data/output',
        'docs'
    ]
    
    for dir_path in directories:
        os.makedirs(dir_path, exist_ok=True)
        print(f"Created directory: {dir_path}")
    
    # 移動源代碼文件到 src/xml_converter
    source_files = {
        'excel_to_xml.py': 'src/xml_converter/excel_to_xml.py',
        'xml_processor.py': 'src/xml_converter/xml_processor.py',
        'streamlit_app.py': 'src/streamlit_app.py',
        'csv_handler.py': 'src/xml_converter/csv_handler.py',
        'main.py': 'src/main.py'
    }
    
    for src, dst in source_files.items():
        if os.path.exists(src):
            shutil.move(src, dst)
            print(f"Moved {src} to {dst}")
    
    # 移動示例文件到 examples
    example_files = {
        'create_example_excel.py': 'examples/create_example_excel.py',
        'example_template.xlsx': 'examples/example_template.xlsx'
    }
    
    for src, dst in example_files.items():
        if os.path.exists(src):
            shutil.move(src, dst)
            print(f"Moved {src} to {dst}")
    
    # 移動數據文件到 data 目錄
    data_files = {
        'converted.xml': 'data/output/converted.xml',
        'PayItemSheet.csv': 'data/output/PayItemSheet.csv',
        'WorkItemSheet.csv': 'data/output/WorkItemSheet.csv'
    }
    
    for src, dst in data_files.items():
        if os.path.exists(src):
            shutil.move(src, dst)
            print(f"Moved {src} to {dst}")
    
    # 創建 __init__.py 文件
    init_files = [
        'src/xml_converter/__init__.py',
        'tests/__init__.py'
    ]
    
    for init_file in init_files:
        with open(init_file, 'w') as f:
            f.write('# Package initialization\n')
        print(f"Created {init_file}")
    
    # 更新 requirements.txt 的位置
    if os.path.exists('requirements.txt'):
        shutil.copy('requirements.txt', 'requirements.txt.bak')
        print("Backed up requirements.txt")

    # 創建 README.md
    readme_content = """# XML Construction Project Converter

A flexible tool for converting construction project Excel files to XML format.

## Project Structure
- `src/`: Source code
  - `xml_converter/`: Core conversion logic
  - `streamlit_app.py`: Web interface
  - `main.py`: Command line interface
- `tests/`: Test files
- `examples/`: Example scripts and templates
- `data/`: Data files
  - `input/`: Input files
  - `output/`: Generated files
- `docs/`: Documentation

## Setup
1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run the Streamlit app:
   ```bash
   streamlit run src/streamlit_app.py
   ```

## Usage
1. Prepare your Excel file with the following columns:
   - 項次 (Item Number)
   - 項目代碼 (Item Code)
   - 項目種類 (Item Kind)
   - 說明 (Description)
   - 單位 (Unit)
   - 數量 (Quantity)
   - 單價 (Unit Price)
   - 金額 (Total Amount)

2. Upload the Excel file through the web interface
3. Convert to XML and download the result

## Example
Run the example script to generate a sample Excel file:
```bash
python examples/create_example_excel.py
```

## Command Line Interface
You can also use the command line interface:
```bash
python src/main.py input.xlsx output.xml
```
"""
    
    with open('README.md', 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print("Created README.md")

    print("\nProject structure reorganized successfully!")
    print("\nYou can now run the application using:")
    print("streamlit run src/streamlit_app.py")
    print("\nOr use the command line interface:")
    print("python src/main.py input.xlsx output.xml")

if __name__ == '__main__':
    setup_project_structure()
