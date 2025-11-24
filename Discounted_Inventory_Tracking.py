# Imports
import xlwings as xw
import pandas as pd
import os
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import shutil
import numpy as np
import gc
import pyodbc
import time
import warnings
import glob
warnings.filterwarnings("ignore", category=UserWarning, message=".*pandas only supports SQLAlchemy.*")

# Begin timer
start_time = time.time()

################################################################################################################################
'''CONNECT SQL QUERIES TO PANDAS'''
################################################################################################################################


def Discounted_Inventory_Tracking_Update():

    # Get current date
    today = datetime.today()

    # Begin timer
    start_time = time.time()

    # Current date
    today = datetime.today()

    # Equivalent to SQL: DECLARE @StartDate DATE = '2025-06-01' (We want to take the past 2 months);
    # Go to the last day of the previous month
    current_week = today - timedelta(days=1)
    last_week = today - timedelta(days=8)
    current_week_str = f"{current_week.month}{current_week.day}"

    current_week = current_week.strftime('%Y-%m-%d')
    last_week = last_week.strftime('%Y-%m-%d')

    print(f"Prior Week date is {last_week} and Current Week date is {current_week}")
    print(f"Current Week str is {current_week_str}")
    # Run SQL queries using SQL Alchemy and dump into Data tab
    
    Discounted_Inventory_Query = f"""
        select SnapshotDate,
        regionName,
        marketName,
        StoreName,
        Hyperion,
        StockType,
        Vin,
        Year,
        Make,
        Model,
        Trim,
        styleid,
        Stylename,
        DaysInInventory,
        PriceTier_93 as WebsitePrice,
        PriceTier_95 as EComPrice,
        MSRP,
        InvoicePrice,
        Balance,
        VinPriced,
        status,
        ExService,
        Loaner

        from [NDDUsers].[vInventory_Daily_Snapshot]

        where SnapshotDate in ('{last_week}', '{current_week}')
        and regionName <> 'AND Corporate Management'
        and marketName not in ('market 97','market 98')
        and StockType = 'new'
        and MSRP>0
        and Year>2021
        and ValidForAN=1
        and ValidForPricing=1

        group by SnapshotDate,
        regionName,
        marketName,
        StoreName,
        Hyperion,
        StockType,
        Vin,
        Year,
        Make,
        Model,
        Trim,
        styleid,
        Stylename,
        DaysInInventory,
        PriceTier_93,
        PriceTier_95,
        MSRP,
        InvoicePrice,
        Balance,
        VinPriced,
        status,
        ExService,
        Loaner
    """

    try:
        with pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=nddprddb01,48155;'
                'DATABASE=NDD_ADP_RAW;'
                'Trusted_Connection=yes;'
        ) as conn:
            print("Reading in SQL Queries...")
            Discounted_Inventory_df = pd.read_sql(Discounted_Inventory_Query, conn)
        end_time = time.time()
        elapsed_time = end_time - start_time
        # Print time it took to load query
        print(f"Script read in Queries and converted to dataframes in {elapsed_time:.2f} seconds")
    except Exception as e:
        print("❌ Connection failed:", e)
################################################################################################################################
    '''CREATE FORMULAS'''
################################################################################################################################

    # Read In Approved Over MSRP Key tab
    Discounted_Inventory_File = r"W:\Corporate\Inventory\BesadaG\Discounted_Inventory_Tracking\Discounted_Inventory_Tracking.xlsm"
    app = xw.App(visible=True) 
    Discounted_Inventory_wb = app.books.open(Discounted_Inventory_File)
    Approved_Over_MSRP_Tab = Discounted_Inventory_wb.sheets['Approved Over MSRP Key']
    
    # Read existing data into a DataFrame
    Approved_Over_MRSP_data_range = Approved_Over_MSRP_Tab.range('A1').expand('down').resize(None, 7)  
    
    # Convert the range to a DataFrame
    Approved_Over_MRSP_df = Approved_Over_MRSP_data_range.options(pd.DataFrame, header=1, index=False).value

    Discounted_Inventory_df['Website_Price_Adj'] = np.where(Discounted_Inventory_df['WebsitePrice'] == 0, "Null", Discounted_Inventory_df['WebsitePrice'])
    Discounted_Inventory_df['Ecom_Price_Adj'] = np.where(Discounted_Inventory_df['EComPrice'] == 0, "Null", Discounted_Inventory_df['EComPrice'])
    # Discounted formula
    Discounted_Inventory_df['Discounted?'] = np.where(
        (Discounted_Inventory_df['WebsitePrice'] < Discounted_Inventory_df['MSRP']) | 
        (Discounted_Inventory_df['EComPrice'] < Discounted_Inventory_df['MSRP']), "Yes", "No")
    # Handle Null values (were defaulting to yes)
    Discounted_Inventory_df['Discounted?'] = np.where(Discounted_Inventory_df['Ecom_Price_Adj'] == "Null", "No", Discounted_Inventory_df['Discounted?'])

    Discounted_Inventory_df['Website_Discount'] = np.where(
        Discounted_Inventory_df['Discounted?'] == "Yes",  
        Discounted_Inventory_df['MSRP'] - Discounted_Inventory_df['WebsitePrice'], 0)
    Discounted_Inventory_df['Ecom_Discount'] = np.where(
        Discounted_Inventory_df['Discounted?'] == "Yes",  
        Discounted_Inventory_df['MSRP'] - Discounted_Inventory_df['EComPrice'], 0)    

    # List of luxury brands
    luxury_brands = ["Audi", "BMW", "Cadillac", "Jaguar", "Land Rover", "Acura", 
                    "MINI", "Porsche", "Mercedes-Benz", "MERCEDES-BENZ TRUCKS", 
                    "Mercedes Light Truck"]

    # Apply logic for avg discount: ignore website discount for luxury brands
    Discounted_Inventory_df['avg discount'] = np.where(
        Discounted_Inventory_df['Make'].isin(luxury_brands),
        Discounted_Inventory_df['Ecom_Discount'],
        Discounted_Inventory_df[['Website_Discount', 'Ecom_Discount']].mean(axis=1)
    )
    
    # Optional: Replace NaNs with "No Discount" if needed
    # Discounted_Inventory_df['Avg_Discount'] = Discounted_Inventory_df['Avg_Discount'].fillna("No Discount")

    Discounted_Inventory_df['Discount_Opportunity'] = np.where(
    (Discounted_Inventory_df['Discounted?'] == "No") &
    (Discounted_Inventory_df['DaysInInventory'].astype(int) > 45),  
    "Opportunity", "N/A")

    # Convert Snapshotdateto datetime
    Discounted_Inventory_df['SnapshotDate'] = pd.to_datetime(Discounted_Inventory_df['SnapshotDate'])

    Discounted_Inventory_df['month'] = Discounted_Inventory_df['SnapshotDate'].dt.month
    Discounted_Inventory_df['day'] = Discounted_Inventory_df['SnapshotDate'].dt.day
    Discounted_Inventory_df['combined'] = (
        Discounted_Inventory_df['month'].astype(str) + 
        Discounted_Inventory_df['day'].astype(str)
    )

    # Compare current day to combined    
    Discounted_Inventory_df['Current/Previous'] = np.where(Discounted_Inventory_df['combined'] == current_week_str, "Current", "Previous")

    Discounted_Inventory_df['Over_45_Days'] = np.where(Discounted_Inventory_df['DaysInInventory']>45, "Yes", "No")

    # Final formula: Lookup
    Discounted_Inventory_df['Key'] = (
        Discounted_Inventory_df['Make'].astype(str).str.strip() +
        Discounted_Inventory_df['Model'].astype(str).str.strip() +
        Discounted_Inventory_df['Trim'].astype(str).str.strip() +
        Discounted_Inventory_df['styleid'].fillna(0).astype(int).astype(str).str.strip() +
        Discounted_Inventory_df['Stylename'].astype(str).str.strip()
    )

    Discounted_Inventory_df = Discounted_Inventory_df.merge(Approved_Over_MRSP_df[['Key', 'Approved Over MSRP']], on='Key', how='left')

    
    Discounted_Inventory_df['Approved Over MSRP'] = Discounted_Inventory_df['Approved Over MSRP'].fillna('Not Approved')


    # Move 'Key' column to the front
    cols = ['Key'] + [col for col in Discounted_Inventory_df.columns if col != 'Key']
    Discounted_Inventory_df = Discounted_Inventory_df[cols]

    #Discounted_Inventory_df.to_csv(r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Discounted_Inventory_Tracking\test.csv')

    # Dump dataset in to Data
    Discounted_Inventory_wb_Sheet = Discounted_Inventory_wb.sheets['Data']
    Discounted_Inventory_wb_Sheet.clear_contents()
    Discounted_Inventory_wb_Sheet.range('A1').options(index=False).value = Discounted_Inventory_df

    # Run Macro
    Run_Macro = Discounted_Inventory_wb.macro("ExecuteMacros")
    Run_Macro()
    Discounted_Inventory_wb.save()

    # wait 60 seconds for file to save before sending copies
    time.sleep(60)

    # Make copies of files in desired folders
    
    # Get today's date in MM.DD.YY format
    today = datetime.today().strftime('%m.%d.%y')

    # Create the filename
    filename1 = f"Discounted Inventory Tracking {today}.xlsm"
    destination_folder = r'W:\Corporate\Inventory\Reporting\Discounted Inventory Tracking\2025'
    full_filename1 = os.path.join(destination_folder, filename1)
    filename2 = f"Discounted Inventory Tracking {today} HC.xlsm"
    full_filename2 = os.path.join(destination_folder, filename2)
    original_file_path1 = r'W:\Corporate\Inventory\BesadaG\Discounted_Inventory_Tracking\Discounted_Inventory_Tracking.xlsm'
    original_file_path2 = r'W:\Corporate\Inventory\BesadaG\Discounted_Inventory_Tracking\Discounted_Inventory_Tracking_HC.xlsm'

    # Copy Values for Discounted % Count then delete data sheet for space then make a HC and a historical copy with todays date

    # Select the sheet you want to keep
    Discounted_Count_Sheet = Discounted_Inventory_wb.sheets['Discounted % Count']

    # Copy all used data and paste as values
    used_range = Discounted_Count_Sheet.used_range
    data = used_range.value
    used_range.value = data  # This pastes it back as values only

    # Delete the 'Data' sheet directly
    Discounted_Inventory_wb.sheets['Data'].delete()

    # Save the HC workbook with a new name without the data tab
    Discounted_Inventory_wb.save(original_file_path2)
    
    # wait 30 seconds for file to save before sending copies
    time.sleep(30)

    # Save and close the excel document(s)    
    if Discounted_Inventory_wb:
        Discounted_Inventory_wb.close()
    if app:
        app.quit()

    # Send copy of HC with todays date
    shutil.copy(original_file_path2, full_filename2)

    # Send copy of full file with todays date
    shutil.copy(original_file_path1, full_filename1)

    # End Timer
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Script completed in {elapsed_time:.2f} seconds")

    return

def Discounted_Inventory_Tracking_Email():

    Discounted_Inventory_File = r"W:\Corporate\Inventory\BesadaG\Discounted_Inventory_Tracking\Discounted_Inventory_Tracking.xlsm"
    app = xw.App(visible=True)
    Discounted_Inventory_wb = app.books.open(Discounted_Inventory_File)
    
    Run_Macro = Discounted_Inventory_wb.macro("Create_Discounted_Inventory_Tracking_Email")
    Run_Macro()
    
    # wait 10 seconds
    time.sleep(10)

    # Save and close the excel document(s)    
    if Discounted_Inventory_wb:
        Discounted_Inventory_wb.close()
    if app:
        app.quit()

    return

#run function
if __name__ == '__main__':
    
    Discounted_Inventory_Tracking_Update()
    Discounted_Inventory_Tracking_Email()