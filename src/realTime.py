import pandas as pd
from openpyxl import load_workbook
import re
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import os

# Define the file paths
tracking_file_path = './data/realtime Tracking.xlsx'
sales_record_file_path = './data/Z2024SalesRecord.xlsx'

def parse_purchase(purchase):
    if isinstance(purchase, float) and pd.isna(purchase):
        return []
    purchase_str = str(purchase)
    items = re.findall(r'(\d+)\s([A-Z0-9-]+)', purchase_str)
    return [(item[1], int(item[0])) for item in items]

def update_tracking_file():
    print("Reading tracking file...")
    # Read the tracking file to get product codes
    tracking_df = pd.read_excel(tracking_file_path, sheet_name='Sheet1')
    product_codes = tracking_df.iloc[:, 0].dropna().tolist()  # Assume the first column is the product code

    # Initialize a dictionary to store the product quantities
    product_quantities = {code: 0 for code in product_codes}

    print("Reading sales record file...")
    # Read the sales record file
    sales_record_df = pd.read_excel(sales_record_file_path)

    print("Summing quantities for each product...")
    # Parse and sum the quantities for each product
    for purchase in sales_record_df['Purchase']:
        items = parse_purchase(purchase)
        for item, quantity in items:
            if item in product_quantities:
                product_quantities[item] += quantity

    print("Updating tracking file with new quantities...")
    # Load the existing workbook
    workbook = load_workbook(tracking_file_path)
    sheet = workbook['Sheet1']

    # Update the tracking file with the new quantities in the existing 'Sales' column (D column)
    sales_column_index = 4  # Column D
    for idx, code in enumerate(product_codes, start=2):
        if code in product_quantities:
            sheet.cell(row=idx, column=sales_column_index).value = product_quantities[code]

    # Save the workbook
    workbook.save(tracking_file_path)
    print("The quantities have been updated in the existing 'Sales' column of the tracking file.")

class SalesRecordHandler(FileSystemEventHandler):
    def on_modified(self, event):
        print(f"Detected change in: {event.src_path}")
        if os.path.abspath(event.src_path) == os.path.abspath(sales_record_file_path):
            print("Updating tracking file...")
            update_tracking_file()

if __name__ == "__main__":
    print("Running initial update...")
    update_tracking_file()  # Initial run to update the file with existing data

    event_handler = SalesRecordHandler()
    observer = Observer()
    observer.schedule(event_handler, path=os.path.dirname(sales_record_file_path), recursive=False)
    observer.start()
    print("Started monitoring for changes...")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
    print("Stopped monitoring for changes.")
