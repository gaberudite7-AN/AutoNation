# Imports
import xlwings as xw
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import shutil
import time
import os


def Update_AstonMartin_Report():

    # Open file and process macro/Sql
    Aston_File = r'W:\Corporate\Inventory\GoldA\Ad Hoc Reports & Data\MY24 Aston Martin Sales.xlsm'
    
    app = xw.App(visible=True) 
    Aston_wb = app.books.open(Aston_File)

    # Run refresh macro
    Run_Macro = Aston_wb.macro("Refresh_Report")
    Run_Macro()
    time.sleep(1)

    # Apply some monitoring (print if theres a change in the sold column?)
    Run_Macro2 = Aston_wb.macro("Monitor_Sales")
    Run_Macro2
    Aston_wb.save()


    # Save and close the excel document(s)    
    if Aston_wb:
        Aston_wb.close()
    if app:
        app.quit()

    # Create the filename
    filename = f"MY24 Aston Martin Sales.xlsm"
    destination_folder = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Aston_Martin'
    full_filename = os.path.join(destination_folder, filename)

    # Send copy of file with todays date
    shutil.copy(Aston_File, full_filename)

    return

def Email_AstonMartin_Report():

    # Open file and process macro/Sql
    Aston_File = r'W:\Corporate\Inventory\GoldA\Ad Hoc Reports & Data\MY24 Aston Martin Sales.xlsm'

    app = xw.App(visible=True) 
    Aston_wb = app.books.open(Aston_File)

    # Run refresh macro
    Run_Macro = Aston_wb.macro("Create_AstonMartin_Email")
    Run_Macro()

    # Save and close the excel document(s)    
    if Aston_wb:
        Aston_wb.close()
    if app:
        app.quit()

    return

#run function
if __name__ == '__main__':

    # Start timer
    start_time = time.time()

    Update_AstonMartin_Report()
    Email_AstonMartin_Report()

    # End Timer
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Script completed in {elapsed_time:.2f} seconds")