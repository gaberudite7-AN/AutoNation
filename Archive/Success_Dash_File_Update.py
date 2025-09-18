import os
import time
import shutil
from datetime import datetime
import openpyxl  # Use openpyxl for .xlsx files

# Configuration
WATCH_FOLDER = r'W:\Corporate\Management Reporting Shared\Data'
#DEST_FOLDER = r'W:\Applications\PowerBI\Pricing and Inventory\Targets and Pace'
DESTINATION_FOLDER = r'W:\Applications\PowerBI\Pricing and Inventory\Targets and Pace'
FILENAME = 'Success Menu Dashboard Data Template_Sales Inventory.xlsx'  # Change to your actual filename
CELL_TO_CHECK = 'N5'

def was_modified_today(filepath):
    mod_time = os.path.getmtime(filepath)
    file_date = datetime.fromtimestamp(mod_time).date()
    return file_date == datetime.today().date()

def was_not_modified_today(filepath):
    mod_time = os.path.getmtime(filepath)
    file_date = datetime.fromtimestamp(mod_time).date()
    return file_date != datetime.today().date()

def check_cell_value(filepath, cell):
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    value = ws[cell].value
    wb.close()
    return value

def update_success_file():
    print("Monitoring file...")

    while True:
        file_path = os.path.join(WATCH_FOLDER, FILENAME)
        destination_path = os.path.join(DESTINATION_FOLDER, FILENAME)

        # if the file exists, if the file was modified today, and if the destination file was not modified today
        if os.path.exists(file_path) and was_modified_today(file_path) and was_not_modified_today(destination_path):
            print("File was modified today and the file needs to be updated")
            try:
                value = check_cell_value(file_path, CELL_TO_CHECK)
                print(f"Checked {CELL_TO_CHECK} value: {value}")

                if value != 0:
                    dest_path = os.path.join(DESTINATION_FOLDER, FILENAME)
                    shutil.copy(file_path, dest_path)
                    print(f"File is not updated for today and N5 is greater than 0. Moved file to: {dest_path}")
            except Exception as e:
                print(f"Error processing file: {e}")
        else:
            print("File either has not been updated today or the file has already been updated")
            time.sleep(10)
    return
if __name__ == "__main__":
    update_success_file()