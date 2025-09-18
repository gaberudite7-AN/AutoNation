import os
import time
import shutil
from datetime import datetime
import openpyxl  # Use openpyxl for .xlsx files
import psutil
import glob

# Configuration
# Success section
WATCH_FOLDER_SUCCESS = r'W:\Corporate\Management Reporting Shared\Data'
#DEST_FOLDER = r'W:\Applications\PowerBI\Pricing and Inventory\Targets and Pace'
DESTINATION_FOLDER_SUCCESS = r'W:\Applications\PowerBI\Pricing and Inventory\Targets and Pace'
FILENAME_SUCCESS = 'Success Menu Dashboard Data Template_Sales Inventory.xlsx'  # Change to your actual FILENAME_SUCCESS
CELL_TO_CHECK_SUCCESS = 'N5'

# Daily Section
WATCH_FOLDER_DAILY = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker'
PROCESSED_FILE_DAILY = os.path.join(WATCH_FOLDER_DAILY, "Dynamic_Daily_Sales.xlsb")

# Doc section
WATCH_FOLDER_DOC = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Brand_President_Tracker\DOC_File'
PROCESSED_FILE_DOC = os.path.join(WATCH_FOLDER_DOC, "Dynamic_DOC_Forecast.xlsb")

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

def Update_Success_File():
    print("Monitoring file...")

    while True:
        file_path = os.path.join(WATCH_FOLDER_SUCCESS, FILENAME_SUCCESS)
        destination_path = os.path.join(DESTINATION_FOLDER_SUCCESS, FILENAME_SUCCESS)

        # if the file exists, if the file was modified today, and if the destination file was not modified today
        if os.path.exists(file_path) and was_modified_today(file_path) and was_not_modified_today(destination_path):
            print("File was modified today and the file needs to be updated")
            try:
                value = check_cell_value(file_path, CELL_TO_CHECK_SUCCESS)
                print(f"Checked {CELL_TO_CHECK_SUCCESS} value: {value}")

                if value != 0:
                    dest_path = os.path.join(DESTINATION_FOLDER_SUCCESS, FILENAME_SUCCESS)
                    shutil.copy(file_path, dest_path)
                    print(f"File is not updated for today and N5 is greater than 0. Moved file to: {dest_path}")
            except Exception as e:
                print(f"Error processing file: {e}")
        else:
            print("File either has not been updated today or the file has already been updated")
            time.sleep(10)
    return

# Set low priority
try:
    p = psutil.Process(os.getpid())
    p.nice(psutil.IDLE_PRIORITY_CLASS)
except Exception as e:
    print(f"Could not set low priority: {e}")

def get_latest_file_daily():
    files = glob.glob(os.path.join(WATCH_FOLDER_DAILY, "*.xlsb"))
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def Process_Daily_Sales_File(latest_file):
    print(f"Processing file: {latest_file}")
    daily_files = glob.glob(os.path.join(WATCH_FOLDER_DAILY, "*Daily*.xlsb"))

    if len(daily_files) == 1:
        print(f"Only one file found: {daily_files[0]}")
    else:
        shutil.copyfile(latest_file, PROCESSED_FILE_DAILY)
        os.remove(latest_file)
        print("File renamed and original deleted.")

def Run_Daily_Sales_Functions():
    print("Polling for new .xlsb files...")
    last_processed = None
    start_time= time.time()
    
    while time.time() - start_time < 30: # Run for 60 seconds
        latest_file = get_latest_file_daily()

        if latest_file and latest_file != last_processed:
            Process_Daily_Sales_File(latest_file)
            last_processed = latest_file

        time.sleep(10)
    print("Waited for 30 seconds. Script completed")

# Set low priority
try:
    p = psutil.Process(os.getpid())
    p.nice(psutil.IDLE_PRIORITY_CLASS)
except Exception as e:
    print(f"Could not set low priority: {e}")

def get_latest_file():
    files = glob.glob(os.path.join(WATCH_FOLDER_DOC, "*.xlsb"))
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def Process_Daily_Sales_File(latest_file):
    print(f"Processing file: {latest_file}")
    daily_files = glob.glob(os.path.join(WATCH_FOLDER_DAILY, "*Dynamic*.xlsb"))

    if len(daily_files) == 1:
        print(f"Only one file found: {daily_files[0]}")
    else:
        shutil.copyfile(latest_file, PROCESSED_FILE_DAILY)
        os.remove(latest_file)
        print("File renamed and original deleted.")

def Run_DOC_Functions():
    print("Polling for new .xlsb files...")
    last_processed = None
    start_time= time.time()
    
    while time.time() - start_time < 30: # Run for 30 seconds
        latest_file = get_latest_file()

        if latest_file and latest_file != last_processed:
            Process_Daily_Sales_File(latest_file)
            last_processed = latest_file

        time.sleep(10)
    print("Waited for 30 seconds. Script completed")

if __name__ == "__main__":

    #Run_DOC_Functions() --  dont need for now
    Update_Success_File()
    #Run_Daily_Sales_Functions()