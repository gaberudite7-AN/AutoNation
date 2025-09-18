import xlwings as xw
import pandas as pd
from datetime import datetime, timedelta
import shutil
import time
import os

class LowPVRReport:
    
    def __init__(self, source_path, destination_folder):
        self.source_path = source_path
        self.destination_folder = destination_folder
        self.app = None
        self.wb = None
        self.start_str, self.end_str = self._get_date_range()

    def _get_date_range(self):
        today = datetime.today()
        last_monday = today - timedelta(days=today.weekday() + 7)
        last_sunday = last_monday + timedelta(days=6)
        start_str = f"{last_monday.month}.{last_monday.day}.{last_monday.strftime('%y')}"
        end_str = f"{last_sunday.month}.{last_sunday.day}.{last_sunday.strftime('%y')}"
        return start_str, end_str

    def open_workbook(self):
        self.app = xw.App(visible=True)
        self.wb = self.app.books.open(self.source_path)

    def close_workbook(self):
        if self.wb:
            self.wb.save()
            self.wb.close()
        if self.app:
            self.app.quit()

    def run_macro(self, macro_name):
        macro = self.wb.macro(macro_name)
        macro()

    def update_report(self):
        self.open_workbook()
        self.run_macro("Refresh_Report")
        time.sleep(120)
        self.close_workbook()
        self._save_copy()

    def email_report(self):
        self.open_workbook()
        self.run_macro("Create_LowPVR_Email")
        self.close_workbook()

    def _save_copy(self):
        filename = f"LOW PVR REPORT NEW & USED {self.start_str} - {self.end_str}.xlsm"
        full_path = os.path.join(self.destination_folder, filename)
        shutil.copy(self.source_path, full_path)

# Run the report
if __name__ == '__main__':
    start_time = time.time()

    source = r'W:\Corporate\Inventory\Reporting\Low PVR Deals\LOW PVR REPORT NEW & USED.xlsm'
    destination = r'W:\Corporate\Inventory\Reporting\Low PVR Deals\Reports'

    report = LowPVRReport(source, destination)
    report.update_report()
    report.email_report()

    elapsed_time = time.time() - start_time