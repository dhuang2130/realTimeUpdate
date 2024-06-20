import os
import openpyxl
import pandas as pd
from datetime import datetime

# Define the cell locations for each required piece of data
cell_locations = {
    "Date of Manu": "K2",
    "SKU": "A1",
    "Product Name": "B1",
    "Lot#": "K3",
    "Shelf Life": "I3",
    "Expiry": "K5",
    "Kit Lot Size": "I5"
}

def extract_data_from_sheet(sheet, cell):
    return sheet[cell].value

directory_path = './data/Products'

xlsx_file_paths = [os.path.join(directory_path, f) for f in os.listdir(directory_path) if f.endswith('.xlsx')]

extracted_data = []
for file_path in xlsx_file_paths:
    wb = openpyxl.load_workbook(file_path, data_only=True)
    for sheet_name in wb.sheetnames:
        if sheet_name.lower() != 'cofa':
            sheet = wb[sheet_name]
            data_row = {key: extract_data_from_sheet(sheet, cell) for key, cell in cell_locations.items()}
            if isinstance(data_row["Date of Manu"], datetime):
                data_row["Date of Manu"] = data_row["Date of Manu"].date()
            if isinstance(data_row["Expiry"], datetime):
                data_row["Expiry"] = data_row["Expiry"].date()
            extracted_data.append(data_row)

df = pd.DataFrame(extracted_data)

output_file_path = os.path.join(directory_path, 'illustrations.xlsx')
with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
    df.to_excel(writer, index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    worksheet.column_dimensions['A'].width = 15  # SKU
    worksheet.column_dimensions['B'].width = 20  # Product Name
    worksheet.column_dimensions['C'].width = 20  # Lot#
    worksheet.column_dimensions['D'].width = 12  # Date of Manu
    worksheet.column_dimensions['E'].width = 12  # Shelf Life
    worksheet.column_dimensions['F'].width = 12  # Expiry
    worksheet.column_dimensions['G'].width = 15  # Kit Lot Size

    date_format = 'yyyy-mm-dd'
    for cell in worksheet['D']:  
        cell.number_format = date_format
    for cell in worksheet['F']:  
        cell.number_format = date_format

print(f"Data extraction and writing to {output_file_path} completed.")
