
import pandas as pd
import os

excel_path = r'C:\Projects\NC\input\铣刀参数_追加钻头数据.xlsx'
print(f"Checking Excel file: {excel_path}")

if not os.path.exists(excel_path):
    print("File does not exist!")
else:
    try:
        xl = pd.ExcelFile(excel_path)
        print(f"Sheet names found: {xl.sheet_names}")
    except Exception as e:
        print(f"Error reading excel file: {e}")
