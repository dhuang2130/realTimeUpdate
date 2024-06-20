import pandas as pd
import xlrd
from openpyxl import Workbook
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

# Function to convert .xls file to .xlsx
def convert_xls_to_xlsx(xls_path, xlsx_path):
    try:
        book = xlrd.open_workbook(xls_path)
        sheet = book.sheet_by_index(0)

        wb = Workbook()
        ws = wb.active

        for row in range(sheet.nrows):
            for col in range(sheet.ncols):
                ws.cell(row=row+1, column=col+1).value = sheet.cell_value(row, col)

        wb.save(xlsx_path)
        logging.info(f"Converted {xls_path} to {xlsx_path} successfully.")
    except Exception as e:
        logging.error(f"Error converting {xls_path} to {xlsx_path}: {e}")

# Paths to the original and converted files
d2dhmanu_xls_path = './data/Products/D2DHManu_2024.xls'
d2dhmanu_xlsx_path = './data/Products/D2DHManu_2024.xlsx'

# Convert the file
convert_xls_to_xlsx(d2dhmanu_xls_path, d2dhmanu_xlsx_path)

# Try reading the converted file
try:
    d2dhmanu_df = pd.read_excel(d2dhmanu_xlsx_path, engine='openpyxl', header=None)
    logging.info(f"Loaded converted file {d2dhmanu_xlsx_path} successfully.")
    print(d2dhmanu_df.head())
except Exception as e:
    logging.error(f"Error reading the converted file {d2dhmanu_xlsx_path}: {e}")
