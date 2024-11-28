# XML Construction Project Converter

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
