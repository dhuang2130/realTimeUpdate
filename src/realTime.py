import pandas as pd
from openpyxl import load_workbook
from xlrd import open_workbook
from xlutils.copy import copy
import re
from tkinter import Tk, Label, Button, filedialog
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import os

# Define the GUI application
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Sales Record Updater")

        self.label1 = Label(root, text="Select the Sales Record File:")
        self.label1.pack(pady=10)

        self.select_sales_file_button = Button(root, text="Select Sales File", command=self.select_sales_file)
        self.select_sales_file_button.pack(pady=5)

        self.label2 = Label(root, text="Select the Tracking File:")
        self.label2.pack(pady=10)

        self.select_tracking_file_button = Button(root, text="Select Tracking File", command=self.select_tracking_file)
        self.select_tracking_file_button.pack(pady=5)

        self.run_button = Button(root, text="Run", command=self.run)
        self.run_button.pack(pady=20)

        self.sales_record_file_path = None
        self.tracking_file_path = None

    def select_sales_file(self):
        self.sales_record_file_path = filedialog.askopenfilename(title="Select Sales Record File", filetypes=[("Excel files", "*.xls *.xlsx")])
        self.label1.config(text=f"Sales Record File: {self.sales_record_file_path}")

    def select_tracking_file(self):
        self.tracking_file_path = filedialog.askopenfilename(title="Select Tracking File", filetypes=[("Excel files", "*.xls *.xlsx")])
        self.label2.config(text=f"Tracking File: {self.tracking_file_path}")

    def run(self):
        if self.sales_record_file_path and self.tracking_file_path:
            try:
                update_tracking_file(self.sales_record_file_path, self.tracking_file_path)
                print("Running initial update...")
                event_handler = SalesRecordHandler(self.sales_record_file_path, self.tracking_file_path)
                observer = Observer()
                observer.schedule(event_handler, path=os.path.dirname(self.sales_record_file_path), recursive=False)
                observer.start()
                print("Started monitoring for changes...")
                try:
                    while True:
                        time.sleep(1)
                except KeyboardInterrupt:
                    observer.stop()
                    observer.join()
                    print("Stopped monitoring for changes.")
            except PermissionError as e:
                print(f"Permission error: {e}")
        else:
            print("Please select both files.")

def parse_purchase(purchase, product_key):
    if isinstance(purchase, float) and pd.isna(purchase):
        return []
    purchase_str = str(purchase)
    # Remove unwanted characters
    purchase_str = purchase_str.replace('Ã—', '').replace(',', '')
    items = re.findall(r'(\d+)\s([A-Z0-9-]+|\w+.*)', purchase_str)
    parsed_items = []
    for item in items:
        code = item[1]
        if code in product_key:
            code = product_key[code]
        parsed_items.append((code, int(item[0])))
    return parsed_items

def update_tracking_file(sales_record_file_path, tracking_file_path):
    print("Reading tracking file...")
    # Read the tracking file to get product codes
    if tracking_file_path.endswith('.xls'):
        tracking_df = pd.read_excel(tracking_file_path, sheet_name='Sheet1', engine='xlrd')
    else:
        tracking_df = pd.read_excel(tracking_file_path, sheet_name='Sheet1')

    product_codes = tracking_df.iloc[:, 0].dropna().tolist()  # Assume the first column is the product code

    # Initialize a dictionary to store the product quantities
    product_quantities = {code: 0 for code in product_codes}

    print("Reading product key from stats tab...")
    # Read the product key from the stats tab
    if sales_record_file_path.endswith('.xls'):
        stats_df = pd.read_excel(sales_record_file_path, sheet_name='Stats', usecols=[1, 2], engine='xlrd')
    else:
        stats_df = pd.read_excel(sales_record_file_path, sheet_name='Stats', usecols=[1, 2])

    product_key = dict(zip(stats_df.iloc[:, 1], stats_df.iloc[:, 0]))

    print("Reading sales record file...")
    # Read the sales record file
    if sales_record_file_path.endswith('.xls'):
        sales_record_df = pd.read_excel(sales_record_file_path, sheet_name='Raw', engine='xlrd')
    else:
        sales_record_df = pd.read_excel(sales_record_file_path, sheet_name='Raw')

    print("Summing quantities for each product...")
    # Parse and sum the quantities for each product
    for purchase in sales_record_df['Purchase']:
        items = parse_purchase(purchase, product_key)
        for item, quantity in items:
            if item in product_quantities:
                product_quantities[item] += quantity

    print("Updating tracking file with new quantities...")
    if tracking_file_path.endswith('.xls'):
        rb = open_workbook(tracking_file_path, formatting_info=True)
        wb = copy(rb)
        ws = wb.get_sheet(0)
        for idx, code in enumerate(product_codes, start=1):
            if code in product_quantities:
                ws.write(idx, 3, product_quantities[code])
        wb.save(tracking_file_path)
    else:
        # Load the existing workbook
        workbook = load_workbook(tracking_file_path)
        sheet = workbook['Sheet1']

        # Update the tracking file with the new quantities in the existing 'Sales' column (D column)
        sales_column_index = 4  # Column D (1-indexed for openpyxl)
        for idx, code in enumerate(product_codes, start=2):
            if code in product_quantities:
                sheet.cell(row=idx, column=sales_column_index).value = product_quantities[code]

        # Save the workbook
        workbook.save(tracking_file_path)

    print("The quantities have been updated in the existing 'Sales' column of the tracking file.")

class SalesRecordHandler(FileSystemEventHandler):
    def __init__(self, sales_record_file_path, tracking_file_path):
        self.sales_record_file_path = sales_record_file_path
        self.tracking_file_path = tracking_file_path

    def on_modified(self, event):
        print(f"Detected change in: {event.src_path}")
        if os.path.abspath(event.src_path) == os.path.abspath(self.sales_record_file_path):
            print("Updating tracking file...")
            try:
                update_tracking_file(self.sales_record_file_path, self.tracking_file_path)
            except PermissionError as e:
                print(f"Permission error: {e}")

# Run the GUI application
if __name__ == "__main__":
    root = Tk()
    app = App(root)
    root.mainloop()
