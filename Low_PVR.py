# Imports
import xlwings as xw
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import shutil
import time
import os


def Update_LowPVR_Report():

    # Open file and process macro/Sql
    LowPVR_File = r'W:\Corporate\Inventory\Reporting\Low PVR Deals\LOW PVR REPORT NEW & USED.xlsm'
    
    app = xw.App(visible=True) 
    LowPVR_wb = app.books.open(LowPVR_File)

    # Gather today and last week dates for file
    today = datetime.today()

    # Calculate last week's Monday
    last_monday = today - timedelta(days=today.weekday() + 7)
    # Calculate last week's Sunday
    last_sunday = last_monday + timedelta(days=6)

    # Format dates as M.D.YY
    start_str = f"{last_monday.month}.{last_monday.day}.{last_monday.strftime('%y')}"
    end_str = f"{last_sunday.month}.{last_sunday.day}.{last_sunday.strftime('%y')}"

    # Run refresh macro
    Run_Macro = LowPVR_wb.macro("Refresh_Report")
    Run_Macro()
    
    # Notify when complated
    time.sleep(120)

    # Save and close the excel document(s)    
    LowPVR_wb.save()
    if LowPVR_wb:
        LowPVR_wb.close()
    if app:
        app.quit()

    # Create the filename
    filename = f"LOW PVR REPORT NEW & USED {start_str} - {end_str}.xlsm"
    destination_folder = r'W:\Corporate\Inventory\Reporting\Low PVR Deals\Reports'
    full_filename = os.path.join(destination_folder, filename)

    # Send copy of file with todays date
    shutil.copy(LowPVR_File, full_filename)

    return

def Email_LowPVR_Report():

    # Open file and process macro/Sql
    LowPVR_File = r'W:\Corporate\Inventory\Reporting\Low PVR Deals\LOW PVR REPORT NEW & USED.xlsm'

    app = xw.App(visible=True) 
    LowPVR_wb = app.books.open(LowPVR_File)

    # Run refresh macro
    Run_Macro = LowPVR_wb.macro("Create_LowPVR_Email")
    Run_Macro()

    # Save and close the excel document(s)    
    if LowPVR_wb:
        LowPVR_wb.close()
    if app:
        app.quit()

    return

#run function
if __name__ == '__main__':

    # Start timer
    start_time = time.time()

    Update_LowPVR_Report()
    Email_LowPVR_Report()

    # End Timer
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Script completed in {elapsed_time:.2f} seconds")