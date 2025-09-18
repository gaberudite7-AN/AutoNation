import os
import time
import glob
import shutil
import psutil

WATCH_FOLDER = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker'
PROCESSED_FILE = os.path.join(WATCH_FOLDER, "Dynamic_Daily_Sales.xlsb")

# Set low priority
try:
    p = psutil.Process(os.getpid())
    p.nice(psutil.IDLE_PRIORITY_CLASS)
except Exception as e:
    print(f"Could not set low priority: {e}")

def get_latest_file():
    files = glob.glob(os.path.join(WATCH_FOLDER, "*.xlsb"))
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def Process_Daily_Sales_File(latest_file):
    print(f"Processing file: {latest_file}")
    daily_files = glob.glob(os.path.join(WATCH_FOLDER, "*Daily*.xlsb"))

    if len(daily_files) == 1:
        print(f"Only one file found: {daily_files[0]}")
    else:
        shutil.copyfile(latest_file, PROCESSED_FILE)
        os.remove(latest_file)
        print("File renamed and original deleted.")

def Run_Daily_Sales_Functions():
    print("Polling for new .xlsb files...")
    last_processed = None
    start_time= time.time()
    
    while time.time() - start_time < 30: # Run for 60 seconds
        latest_file = get_latest_file()

        if latest_file and latest_file != last_processed:
            Process_Daily_Sales_File(latest_file)
            last_processed = latest_file

        time.sleep(10)
    print("Waited for 30 seconds. Script completed")

if __name__ == "__main__":

    Run_Daily_Sales_Functions()
    