import time
import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

illustrations_path = './data/Products/illustrations.xlsx'
realtime_tracking_path = './data/realtime Tracking.xlsx'

class UpdateHandler(FileSystemEventHandler):
    def on_modified(self, event):
        if event.src_path.endswith('illustrations.xlsx'):
            print(f'{event.src_path} has been modified')
            update_realtime_tracking()

def update_realtime_tracking():
    illustrations_df = pd.read_excel(illustrations_path)
    realtime_tracking_df = pd.read_excel(realtime_tracking_path)

    aggregated_illustrations_df = illustrations_df.groupby('SKU')['Kit Lot Size'].sum().reset_index()

    realtime_tracking_df.rename(columns={realtime_tracking_df.columns[0]: 'Product'}, inplace=True)

    merged_df = pd.merge(realtime_tracking_df, aggregated_illustrations_df, left_on='Product', right_on='SKU', how='left')

    realtime_tracking_df['Manufacture'] = merged_df['Kit Lot Size']

    realtime_tracking_df.to_excel(realtime_tracking_path, index=False)

    print("Updated Realtime Tracking DataFrame:")
    print(realtime_tracking_df.head())

if __name__ == "__main__":
    event_handler = UpdateHandler()
    observer = Observer()
    observer.schedule(event_handler, path='./data/Products', recursive=False)
    observer.start()

    print("Monitoring Manufacture started. Press Ctrl+C to stop.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
