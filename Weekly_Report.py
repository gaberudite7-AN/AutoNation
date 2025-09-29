# Imports
import xlwings as xw
import pandas as pd
import os
from datetime import datetime, timedelta
import shutil
import numpy as np
import gc
import pyodbc
import time
import warnings
import glob
import psutil
warnings.filterwarnings("ignore", category=UserWarning, message=".*pandas only supports SQLAlchemy.*")
from dateutil.relativedelta import relativedelta

# Selenium packages
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import undetected_chromedriver as uc
import traceback


# Run with low priority ( will allow script to run in background and yield CPU to other apps)
try:
    p = psutil.Process(os.getpid())
    p.nice(psutil.IDLE_PRIORITY_CLASS) # Windows only
except Exception as e:
    print(f"Could not set low priority: {e}")


def Process_Daily_Sales_File():

    # Open up most recent Daily Sales file: 
    # Get all .xlsm files in the Arrivals folder
    Dynamic_Daily_Sales_Path = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker'
    Dynamic_Daily_Sales_Files = glob.glob(os.path.join(Dynamic_Daily_Sales_Path, "*.xlsb"))
    Dynamic_Daily_Sales_File = max(Dynamic_Daily_Sales_Files, key=os.path.getmtime)
    print(f"Latest .xlsb file: {Dynamic_Daily_Sales_File}")

    # Since Powerautomate cannot change the name of the file we will use python to rename it
    # However if it has already been renamed and the new file deleted...we need an if statement to process as is 
    # otherwise continue creating new file and removing old

    # Get all .xlsb files with "Daily" in the name
    daily_files = glob.glob(os.path.join(Dynamic_Daily_Sales_Path, "*Daily*.xlsb"))

    if len(daily_files) == 1:
        # Only one file — use it as Dynamic_Daily_Sales.xlsb
        Dynamic_Daily_Sales_File = daily_files[0]
        print(f"Only one file found: {Dynamic_Daily_Sales_File}")

    else:   
        # Rename file to Dynamic_Daily_Sales.xlsb    
        shutil.copyfile(Dynamic_Daily_Sales_File, r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker\Dynamic_Daily_Sales.xlsb')
                
        # Delete old file
        os.remove(Dynamic_Daily_Sales_File)
        print("Successfully removed old file and replaced current with new.")
        return

################################################################################################################################
'''CONNECT SQL QUERIES TO PANDAS'''
################################################################################################################################

def Download_PWB():
# Function to download the latest Referall Sharepoint to local

    # Setup Chrome options
    chrome_options = uc.ChromeOptions()
    # chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36")

    # Paths
    chrome_driver_path = r"C:\Development\Chrome_Driver\chromedriver-win64\chromedriver.exe"
    downloads_folder = r"C:\Users\BesadaG\Downloads"
    destination_folder = r"W:\Corporate\Inventory\Weekly Reporting Package\PWB_Data"
    PWB_Link = "https://pricingworkbench.autonation.com/"

    # Format today's date as YYYY-MM-DD
    today_str = datetime.today().strftime("%Y-%m-%d")
    filename = f"PW_{today_str}.xlsx"

    # Start browser
    try:
        driver = uc.Chrome(
            driver_executable_path=chrome_driver_path,
            options=chrome_options,
            use_subprocess=True
        )

        actions = ActionChains(driver)
        driver.set_page_load_timeout(20)
        driver.get(PWB_Link)

        # Define wait AFTER driver is initialized
        wait = WebDriverWait(driver, 20)
        time.sleep(5)

        Email = "besadag@autonation.com"

        time.sleep(1)
        
        # Step 1: Enter email
        email_input = wait.until(EC.presence_of_element_located((By.ID, "i0116")))
        email_input.send_keys(Email)
        time.sleep(2)

        # Click the Submit button
        submit_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        submit_button.click()
        time.sleep(3)

        # Click the "Continue" button after entering email or password
        continue_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        continue_button.click()
        time.sleep(15) # wait for site data to load

        # Click the "Yes" button on the "Stay signed in?" prompt
        # yes_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        # yes_button.click()
        # time.sleep(3)

        # Update the XPath to the new value
        download_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/app-root/div/div/main/app-home/div/app-chart/img'))
        )
        download_button.click()
        time.sleep(15)  # Wait for download to complete

        # Move latest file to destination folder with adjusted name
        source_file = os.path.join(downloads_folder, filename)
        destination_file = os.path.join(destination_folder, filename)

        shutil.move(source_file, destination_file)
        print(f"Successfully moved file to: {destination_file}")

    except Exception as e:
        error_log_path = os.path.join(destination_folder, "PWB_error_log.txt")
        with open(error_log_path, "w") as f:
            f.write(traceback.format_exc())
        print(f"An error occurred. Details written to {error_log_path}")

    finally:
        if 'driver' in locals():
            def safe_del(self):
                try:
                    self.quit()
                except Exception:
                    pass  # Silently ignore all errors
            uc.Chrome.__del__ = safe_del
        
        return

def Update_PWB_Data():
    
    PWB_Data_folder = r"W:\Corporate\Inventory\Weekly Reporting Package\PWB_Data"
    
    # Format today's date as YYYY-MM-DD
    today_str = datetime.today().strftime("%Y-%m-%d")
    filename = f"PW_{today_str}.xlsx"    
    
    # Latest Data file
    latest_data = os.path.join(PWB_Data_folder, filename)
    PWB_file = r"W:\Corporate\Inventory\Weekly Reporting Package\Pricing_Workbench_Analysis_Used.xlsx"

    # Open with xlwings
    app = xw.App(visible=True)
    PWB_wb = app.books.open(PWB_file)
    latest_data_wb = app.books.open(latest_data)

    # Go to date column of PWB file and remove earliest date
    '''Update PWB'''
    PWB_sht = PWB_wb.sheets['Pricing_WB']
    # Read existing data into a DataFrame
    data_range = PWB_sht.range('A1').expand()
    PWB_df = data_range.options(pd.DataFrame, header=1, index=False).value

    # Find the earliest date
    PWB_df['Date'] = pd.to_datetime(PWB_df['Date'])
    earliest_date = PWB_df['Date'].min()

    # Read latest file data into dataframe
    latest_data_sht = latest_data_wb.sheets['ag-grid']
    data_range = latest_data_sht.range('A2').expand('down').resize(None, 64)
    latest_PWB_df = data_range.options(pd.DataFrame, header=1, index=False).value

    # Ensure RN is a column, not an index
    PWB_df = PWB_df.reset_index()
    latest_PWB_df = latest_PWB_df.reset_index()

    # Filter out rows with the earliest date
    filtered_PWB_Data = PWB_df[PWB_df['Date'] > earliest_date]

    # Add Current Date to current Allocation Dataset
    today = datetime.today()
    latest_PWB_df['Date'] = pd.to_datetime(today)

    # Append new data
    combined_df = pd.concat([filtered_PWB_Data, latest_PWB_df], ignore_index=True)

    # Drop the 'index' column if it exists
    if 'index' in combined_df.columns:
        combined_df = combined_df.drop('index', axis=1)

    # Convert date to not include time values
    combined_df['Date'] = combined_df['Date'].dt.date

    # Replace PWB worksheet with updated data
    PWB_sht.clear_contents()
    PWB_sht.range('A1').options(index=False).value = combined_df

    # Save and close the excel document and any other open instances
    PWB_wb.save(PWB_file)
    if PWB_wb:
        PWB_wb.close()
    if app:
        app.quit()

    return

def Weekly_Data_Update():

    # Begin timer
    start_time = time.time()

    # Extract dates for dynamic modification of queries (if need prior month and 2 months back)
    # today = datetime.today()
    # Current_accounting_month = today.replace(day=1) - timedelta(days=1)
    # Current_accounting_month = Current_accounting_month.replace(day=1)
    # Current_accounting_month = f"{Current_accounting_month.month}/{Current_accounting_month.day}/{Current_accounting_month.year}"
    # print(f"Current accounting month is {Current_accounting_month}")


    # Prior_accounting_month = (today.replace(day=1) - relativedelta(months=2)).replace(day=1)
    # Prior_accounting_month = Prior_accounting_month.replace(day=1)
    # Prior_accounting_month = f"{Prior_accounting_month.month}/{Prior_accounting_month.day}/{Prior_accounting_month.year}"
    # print(f"Prior accounting month is {Prior_accounting_month}")

    # Current month and prior month 
    today = datetime.today()
    Current_accounting_month = today
    Current_accounting_month = Current_accounting_month.replace(day=1)
    Current_accounting_month = f"{Current_accounting_month.month}/{Current_accounting_month.day}/{Current_accounting_month.year}"
    print(f"Current accounting month is {Current_accounting_month}")


    Prior_accounting_month = (today.replace(day=1) - relativedelta(months=1)).replace(day=1)
    Prior_accounting_month = Prior_accounting_month.replace(day=1)
    Prior_accounting_month = f"{Prior_accounting_month.month}/{Prior_accounting_month.day}/{Prior_accounting_month.year}"
    print(f"Prior accounting month is {Prior_accounting_month}")



    Used_Inventory_query = f"""
With temptable1 as (Select ROW_NUMBER() Over (Partition by a.VIN order by SnapshotDate desc) as RN,
SnapshotDate, hyperion, CASE WHEN Original_InventorySourceCode IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN Original_InventorySourceCode IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN Original_InventorySourceCode IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN Original_InventorySourceCode IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN Original_InventorySourceCode IN ('WEBUYYOURCAR') THEN 'WBYC'
      WHEN InventorySourceCode IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN InventorySourceCode IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN InventorySourceCode IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN InventorySourceCode IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN InventorySourceCode IN ('WEBUYYOURCAR') THEN 'WBYC'
      ELSE 'Other' END AS SourceCode
, Inventorysourcename, VIN, CONCAT(vin,hyperion) as HyperionVINFlag, targetprice

From nddusers.vInventory_Daily_Snapshot a

Where department = '320'
and CAST(snapshotdate as date) between CAST(GETDATE() - 50 as date) and CAST(GETDATE() - 0 as date) and status <> 'G' 

) 

, temptable2 as (

Select VIN, SourceCode as Source

from temptable1 a

where rn = 1
)

, temptableFINAL as (
Select case when VIN is null then 1 else ROW_NUMBER() Over (Partition by VIN order by TypeNo) end as RN, *, 
case when InventorySource in ('Trade-In','WBYC','Lease Return','Service Loaner') then 'Total: Int. Sourced' else 'Total: Ext. Sourced' end as 'Internal/External',
case when Week = 'Full Month' then 0
when Week = 'Week 1' then 1
when Week = 'Week 2' then 2
when Week = 'Week 3' then 3
when Week = 'Week 4' then 4
when Week = 'Week 5' then 5
else '' end as Flag


from (

Select VIN,
StoreName,
InventorySource,
Status,
SumInventory,
PriceBucket as Invpricebucket,
accountingmonth,
'Week 1' as Week,      ---------------------------------------------CHANGE--------------------------------------------------------
'' as SPACE,
ValidPriceBucket, AgeBucket,
Sumbalance,
SumTargetPrice, TypeNo, Type

From(
select AccountingMonth,
RegionName,
MarketName,
Hyperion,
StoreName,
Year,
Make,
Model,
a.Vin,
0 as SumSold,
0 as SumCashPrice,
0 as SumFrontGross,
0 as SumInterCoGross,
0 as SumVehicleAge,
0 as SumVehicleAgeAN,
sum(inventorycount) as SumInventory, 
sum(balance) as SumBalance,
sum(pricetier_93) as SumWebsitePrice, 
sum(targetprice) as SumTargetPrice, 
sum(pricetier_93) - sum(balance) as Margin,
--sum(pricetier_93)/sum(targetprice) as PriceToTarget,
--sum(balance)/sum(targetprice) as CostToTarget,
'Inventory' as RecordSource,

case  when DaysInInventoryAN IS NULL then 'N/A'
  when DaysInInventoryAN < 16 then '01: 0-15'
  when DaysInInventoryAN < 31 then '02: 16-30'
  when DaysInInventoryAN < 46 then '03: 31-45'
  when DaysInInventoryAN < 61 then '04: 46-60'
  when DaysInInventoryAN < 91 then '05: 61-90'
  when DaysInInventoryAN < 121 then '06: 91-120'
else '07: Over 120' end as 'AgeBucket',

case when DaysInInventoryAN is null then 'N/A'
when DaysInInventoryAN <45 then 'Under 45 Days'
when DaysInInventoryAN >=45 then 'Over/At 45 Days'
else 'N/A' end as '45DayBucket',

case when Status = 's' then 'S Status'
else 'All Other Status' end as 'Status',

case when PriceTier_93 between 0.01 and 20000 then '1:0-20,000'
 when PriceTier_93 between 20000 and 40000 then '2:20,001 - 40,000'
when PriceTier_93 between 40000 and 1000000 then '3:Over 40,001'
 when Balance between 0.01 and 20000 then '1:0-20,000'
 when Balance between 20000 and 40000 then '2:20,001 - 40,000'
 when Balance between 40000 and 1000000 then '3:Over 40,001'
 else 'N/A' end as 'PriceBucket',

case when pricetier_93 is null then 'N/A'
when PriceTier_93 <1 then 'N/A'
when Balance is null then 'N/A'
when Balance <1 then 'N/A'
when TargetPrice is null then 'N/A'
when TargetPrice <1 then 'N/A'
else 'Valid' end as 'ValidPriceBucket',

CASE WHEN Original_InventorySourceCode IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN Original_InventorySourceCode IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN Original_InventorySourceCode IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN Original_InventorySourceCode IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN Original_InventorySourceCode IN ('WEBUYYOURCAR') THEN 'WBYC'
      WHEN SourceType IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN SourceType IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN SourceType IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN SourceType IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN SourceType IN ('WEBUYYOURCAR') THEN 'WBYC'
      WHEN b.Source IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN b.Source IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN b.Source IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN b.Source IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN b.Source IN ('WEBUYYOURCAR') THEN 'WBYC'
      ELSE 'Other' END AS 'InventorySource', 1 as TypeNo, 'MonthEnd' as Type

from NDD_ADP_RAW.NDDUsers.vInventoryMonthEnd a
left join temptable2 b on a.VIN = b.VIN

where Department = '320'
and AccountingMonth in ('{Current_accounting_month}')
and RegionName <> 'AND Corporate Management'
and MarketName not in ('Franchise Market 98','Auction')


group by AccountingMonth,
RegionName,
MarketName,
Hyperion,
StoreName,
Year,
Make,
Model,
a.Vin,
case  when DaysInInventoryAN IS NULL then 'N/A'
  when DaysInInventoryAN < 16 then '01: 0-15'
  when DaysInInventoryAN < 31 then '02: 16-30'
  when DaysInInventoryAN < 46 then '03: 31-45'
  when DaysInInventoryAN < 61 then '04: 46-60'
  when DaysInInventoryAN < 91 then '05: 61-90'
  when DaysInInventoryAN < 121 then '06: 91-120'
else '07: Over 120' end,

case when DaysInInventoryAN is null then 'N/A'
when DaysInInventoryAN <45 then 'Under 45 Days'
when DaysInInventoryAN >=45 then 'Over/At 45 Days'
else 'N/A' end,

case when Status = 's' then 'S Status'
else 'All Other Status' end,

case when PriceTier_93 between 0.01 and 20000 then '1:0-20,000'
 when PriceTier_93 between 20000 and 40000 then '2:20,001 - 40,000'
when PriceTier_93 between 40000 and 1000000 then '3:Over 40,001'
 when Balance between 0.01 and 20000 then '1:0-20,000'
 when Balance between 20000 and 40000 then '2:20,001 - 40,000'
 when Balance between 40000 and 1000000 then '3:Over 40,001'
 else 'N/A' end,

case when pricetier_93 is null then 'N/A'
when PriceTier_93 <1 then 'N/A'
when Balance is null then 'N/A'
when Balance <1 then 'N/A'
when TargetPrice is null then 'N/A'
when TargetPrice <1 then 'N/A'
else 'Valid' end,

CASE WHEN Original_InventorySourceCode IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN Original_InventorySourceCode IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN Original_InventorySourceCode IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN Original_InventorySourceCode IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN Original_InventorySourceCode IN ('WEBUYYOURCAR') THEN 'WBYC'
      WHEN SourceType IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN SourceType IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN SourceType IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN SourceType IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN SourceType IN ('WEBUYYOURCAR') THEN 'WBYC'
      WHEN b.Source IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN b.Source IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN b.Source IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN b.Source IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN b.Source IN ('WEBUYYOURCAR') THEN 'WBYC'
      ELSE 'Other' END


having sum(inventorycount) =1
) A

Where status = 'S Status'

union all

Select a.VIN,
a.StoreName,
CASE WHEN Original_InventorySourceCode IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN Original_InventorySourceCode IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN Original_InventorySourceCode IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN Original_InventorySourceCode IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN Original_InventorySourceCode IN ('WEBUYYOURCAR') THEN 'WBYC'
      WHEN InventorySourceCode IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN InventorySourceCode IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN InventorySourceCode IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN InventorySourceCode IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN InventorySourceCode IN ('WEBUYYOURCAR') THEN 'WBYC'
  WHEN b.Source IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN b.Source IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN b.Source IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN b.Source IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN b.Source IN ('WEBUYYOURCAR') THEN 'WBYC'
      ELSE 'Other' END AS InventorySource,

'S Status' as Status,
1 as SumInventory,
case when PriceTier_93 is null then 'N/A'
 when PriceTier_93 <20001 then '1:0-20,000'
 when PriceTier_93 <40001 then '2:20,001 - 40,000'
 else '3:Over 40,001' end as 'PriceBucket',
 '{Current_accounting_month}' as Accountingmonth,
 'Week 1' as Week,      ---------------------------------------------CHANGE--------------------------------------------------------
 '' as Space,
case when pricetier_93 is null then 'N/A'
when PriceTier_93 <1 then 'N/A'
when Balance is null then 'N/A'
when Balance <1 then 'N/A'
when TargetPrice is null then 'N/A'
when TargetPrice <1 then 'N/A'
else 'Valid' end as 'ValidPriceBucket',
case  when DaysInInventoryAN IS NULL then 'N/A'
  when DaysInInventoryAN < 16 then '01: 0-15'
  when DaysInInventoryAN < 31 then '02: 16-30'
  when DaysInInventoryAN < 46 then '03: 31-45'
  when DaysInInventoryAN < 61 then '04: 46-60'
  when DaysInInventoryAN < 91 then '05: 61-90'
  when DaysInInventoryAN < 121 then '06: 91-120'
else '07: Over 120' end as 'AgeBucket',
 Balance as SumBalance,
 TargetPrice as SumTargetPrice, 2 as TypeNo, 'Live' as Type

From nddusers.vInventory a
left join temptable2 b on a.VIN = b.VIN

Where stocktype = 'Used'
 and ValidForAN = 1

) A

)

Select * from temptableFINAL where RN = 1
Order by TypeNo



        """

    Used_sales_query = f"""
    Select Accountingmonth as Month,
    case when month(accountingmonth) < 4 then 'Q1'
    when month(accountingmonth) < 7 then 'Q2'
    when month(accountingmonth) < 10 then 'Q3'
    when month(accountingmonth) < 13 then 'Q4' else 'Unk' end as Quarter,
    VehicleMakeName as Make,
    MarketName as Market,
    VehicleModelName as Model,
    StoreHyperion as Hyperion,
    StoreName as Store,
    VIN,
    sum(vehiclesoldcount) as SoldCount,
    sum(WholesaleInterCompanyCount) as ICCount,
    sum(WholesaleAuctionCount) as WholeSaleCount,
    sum(FITotalCount) as CFSCount,
    sum(frontgross) as BaseGross,
    sum(wholesaleintercompanygross) as ICGross,
    sum(WholesaleAuctionGross) as WholeSaleGross,
    sum(FrontRevenue) as RetailAVGPrice,
    sum(WholesaleInterCompanyRevenue) as ICAVGPrice,
    sum(WholesaleAuctionRevenue) as WholeSaleAVGPrice,
    sum(figross) as CFSGross,
    sum(OVI) as OVIGross,
    case when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('Active Loaner','ActiveLoaner','RetiredLoaner','Manufacturer') then 'ExLoaner'
    when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('Auction','CAAD','Rental-Co') then 'Auction'
    when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('Dealer-Trade','Trade-in','Trade-in-used') then 'Trade'
    when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('Inter-Co') then 'Interco'
    when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('Leasebuyout','Leasestrpurch') then 'Lease Return'
    when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('WEBUYYOURCAR') then 'WBYC'
    when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('Consign','EMCC','NULL') then 'Other' else 'Other' end as Source,

    max(abs(VehicleAgeAN)) * sum(vehiclesoldcount) as AvgANDaysToSell,
    0 as AvgANDaysInInv,

    case when max(abs(VehicleAgeAN)) between 0 and 45 then '0-45'
    else '>45' end as AN_Age,

    case when max(abs(VehicleAge)) between 0 and 45 then '0-45'
    else '>45' end as StoreAge,

    Case when max(abs(CashPrice)) between 1 and 20000 then '$0-$20k'
    when max(abs(CashPrice)) between 20000 and 40000 then '$20-$40k'
    when max(abs(CashPrice)) between 40000 and 1000000 then '$40k+'
    when max(abs(Wholesaleintercompanyrevenue)) between 1 and 20000 then '$0-$20k'
    when max(abs(Wholesaleintercompanyrevenue)) between 20000 and 40000 then '$20-$40k'
    when max(abs(Wholesaleintercompanyrevenue)) between 40000 and 1000000 then '$40k+' else 'Adjustments' end as PriceBand,

    case when sum(vehiclesoldcount) = 0 then 0 when max(abs(CashPrice)) is null then 0 when max(abs(CashPrice)) = 0 then 0
    when sum(frontcos) is null then 0 when sum(frontcos) = 0 then 0
    when max(abs(TargetPrice)) is null then 0 when max(abs(TargetPrice)) = 0 then 0  when max(abs(CashPrice)) > (max(abs(TargetPrice)) + 10000) then 0
    when max(abs(TargetPrice)) < 1000 then 0 when max(abs(CashPrice)) < 1000 then 0 when (max(abs(CashPrice)) + 10000) < max(abs(TargetPrice)) then 0
    else max(abs(CashPrice))  end as CashPrice,

    case when sum(vehiclesoldcount) = 0 then 0 when max(abs(CashPrice)) is null then 0 when max(abs(CashPrice)) = 0 then 0
    when sum(frontcos) is null then 0 when sum(frontcos) = 0 then 0
    when max(abs(TargetPrice)) is null then 0 when max(abs(TargetPrice)) = 0 then 0  when max(abs(CashPrice)) > (max(abs(TargetPrice)) + 10000) then 0
    when max(abs(TargetPrice)) < 1000 then 0 when max(abs(CashPrice)) < 1000 then 0 when (max(abs(CashPrice)) + 10000) < max(abs(TargetPrice)) then 0
    else sum(frontcos) end as Cost,

    case when sum(vehiclesoldcount) = 0 then 0 when max(abs(CashPrice)) is null then 0 when max(abs(CashPrice)) = 0 then 0
    when sum(frontcos) is null then 0 when sum(frontcos) = 0 then 0
    when max(abs(TargetPrice)) is null then 0 when max(abs(TargetPrice)) = 0 then 0  when max(abs(CashPrice)) > (max(abs(TargetPrice)) + 10000) then 0
    when max(abs(TargetPrice)) < 1000 then 0 when max(abs(CashPrice)) < 1000 then 0 when (max(abs(CashPrice)) + 10000) < max(abs(TargetPrice)) then 0
    else max(abs(TargetPrice)) end as TargetPrice,

    case when sum(vehiclesoldcount) = 0 then 0 when max(abs(CashPrice)) is null then 0 when max(abs(CashPrice)) = 0 then 0
    when sum(frontcos) is null then 0 when sum(frontcos) = 0 then 0
    when max(abs(TargetPrice)) is null then 0 when max(abs(TargetPrice)) = 0 then 0  when max(abs(CashPrice)) > (max(abs(TargetPrice)) + 10000) then 0
    when max(abs(TargetPrice)) < 1000 then 0 when max(abs(CashPrice)) < 1000 then 0 when (max(abs(CashPrice)) + 10000) < max(abs(TargetPrice)) then 0
    else max(abs(vehiclesoldcount)) end as SoldCountAdj,

    0 as InvCount,
    0 as InvPrice,
    0 as InvBalance,
    0 as InvTargetPrice,
    0 as InventoryCountAdj

    From NDDUsers.vSalesDetail_Vehicle

    Where departmentname = 'Used'
    and accountingmonth = '{Current_accounting_month}'
    and RegionName IN ('Region 1', 'Region 2')
    and recordsource = 'Accounting'

    Group by Accountingmonth,
    VehicleMakeName,
    MarketName,
    VehicleModelName,
    StoreHyperion,
    StoreName,
    VIN,
    case when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('Active Loaner','ActiveLoaner','RetiredLoaner','Manufacturer') then 'ExLoaner'
    when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('Auction','CAAD','Rental-Co') then 'Auction'
    when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('Dealer-Trade','Trade-in','Trade-in-used') then 'Trade'
    when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('Inter-Co') then 'Interco'
    when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('Leasebuyout','Leasestrpurch') then 'Lease Return'
    when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('WEBUYYOURCAR') then 'WBYC'
    when Case when left(inventorysourcename,9) = 'From 2090' then 'CAAD' else SourceType end in ('Consign','EMCC','NULL') then 'Other' else 'Other' end

    Having sum(vehiclesoldcount) <> 0 or sum(frontgross) <> 0 or sum(wholesaleintercompanygross) <> 0 or sum(OVI) <> 0 or sum(figross) <> 0
    or sum(WholesaleAuctionCount) <> 0 or sum(WholesaleInterCompanyCount) <> 0 or sum(WholesaleAuctionGross) <> 0
"""

    Used_Pricing_query = f"""
with temptable1 as (
select cast(snapshotdate as date) as SnapshotDate, 
MarketName, 
Storename, 
Status, 
VIN,
Year,
Make,
Model,
Trim,
Mileage,
CASE WHEN Original_InventorySourceCode IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN Original_InventorySourceCode IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN Original_InventorySourceCode IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN Original_InventorySourceCode IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN Original_InventorySourceCode IN ('WEBUYYOURCAR') THEN 'WBYC'
      WHEN InventorySourceCode IN ('Rental-Co','Auction') THEN 'AUCTION/RENTAL' 
      WHEN InventorySourceCode IN ('Leasestrpurch','Leasebuyout') THEN 'LEASE RETURN'
      WHEN InventorySourceCode IN ('TRADE-IN','TRADE-IN-USED') THEN 'TRADE-IN'
      WHEN InventorySourceCode IN ('Activeloaner','Retiredloaner') THEN 'SERVICE LOANER'
      WHEN InventorySourceCode IN ('WEBUYYOURCAR') THEN 'WBYC'
      ELSE 'Other' END AS 'OriginalSource',
DaysInInventoryAN as AN_Age,
DaysInInventory as StoreDays,
InternetPrice as Price, 
Balance as Cost,
isnull(targetprice,0) as Target, 
CarGurus_good_price as CG_GoodPrice,
 
 
case when InternetPrice >= CarGurus_poor_price + 0.01 then 'CG_Overpriced'
when InternetPrice between CarGurus_fair_price + 0.01 and CarGurus_poor_price then 'CG_High'
when InternetPrice between CarGurus_good_price + 0.01 and CarGurus_fair_price then 'CG_Fair'
when InternetPrice between CarGurus_great_price + 0.01  and CarGurus_good_price  then 'CG_Good'
when InternetPrice <= CarGurus_great_price then 'CG_Great' else 'CG_Unk' end as CG_Ranking,
 
ROW_NUMBER() OVER (Partition by VIN, StoreName order by snapshotdate asc) as VINCount, 
ROW_NUMBER() OVER (Partition by VIN, StoreName, InternetPrice order by snapshotdate asc) as DaysSinceLPChange, 
ROW_NUMBER() OVER (Partition by VIN, StoreName, InternetPrice order by snapshotdate desc) as PreviousRecord, 
case when ROW_NUMBER() OVER (Partition by VIN, StoreName, InternetPrice order by snapshotdate asc) = 1 then 1 else 0 end as PriceChangeFlag
 
 
from nddusers.vinventory_daily_snapshot
 
where CAST(snapshotdate as date) >= CAST(GETDATE() - 65 as date)
and status = 'S' and department = '320' and internetprice is not null
--and vin = '15GGD181251075024'
 
 
)
 
, temptable2 as (
Select *, (SELECT MAX(PriceChangeFlag) FROM temptable1 b WHERE b.VIN = a.VIN) AS MaxPriceChangeFlag, ROW_NUMBER() OVER (Partition by VIN, StoreName order by snapshotdate asc) as NewVINCount
from temptable1 a 
Where ((DaysSinceLPChange = 1 and VINCount <> 1) or PreviousRecord = 1) and (SELECT MAX(PriceChangeFlag) FROM temptable1 b WHERE b.VIN = a.VIN) = 1
 
)
 
, temptable3 as (
 
Select *, (SELECT MAX(NewVINCount) FROM temptable2 b WHERE b.VIN = a.VIN) MaxNewVINCount
from temptable2 a
 
)
 
Select SnapshotDate, MarketName, StoreName, Status, VIN, Year, Make, Model, Trim, Mileage, OriginalSource, AN_Age, StoreDays, Price, Cost,
 
case when Target is not null and Target > 0 and CG_GoodPrice is not null and CG_GoodPrice > 0 then Price else 0 end as PriceAdj,
case when Target is not null and Target > 0 and CG_GoodPrice is not null and CG_GoodPrice > 0 then Cost else 0 end as CostAdj,
case when Target is not null and Target > 0 and CG_GoodPrice is not null and CG_GoodPrice > 0 then Target else 0 end as Target,
case when Target is not null and Target > 0 and CG_GoodPrice is not null and CG_GoodPrice > 0 then CG_GoodPrice else 0 end as CG_GoodPrice,
 
CG_Ranking,
DaysSinceLPChange, PriceChangeFlag as PriceChangeCount, 1 as Count
from temptable3
Where MaxNewVINCount > 1 and CAST(snapshotdate as date) >= (GETDATE() - 14) --- change cutoff days to 14 and 7
Order by VIN, SnapshotDate
"""
    
    New_Inventory_query = f"""
-- Step 1: Master Brand List
WITH BrandList AS (
    SELECT BrandName FROM (VALUES
        ('CDJR'), ('Ford'), ('GM'),
        ('Acura'), ('Honda'), ('Hyundai'), ('INFINITI'), ('Mazda'),
        ('Nissan'), ('Subaru'), ('Toyota'), ('Volkswagen'), ('Volvo'),
        ('Aston Martin'), ('Audi'), ('BMW'), ('Bentley'), ('JLR'),
        ('Lexus'), ('Mercedes-Benz'), ('MINI'), ('Porsche'),
        ('Others')
    ) AS T(BrandName)
),

-- Step 2: Inventory Source (Brand, Category, INV)
MappedData AS (
    SELECT 
        AccountingMonth,
        CASE
            WHEN Make IN ('CHRYSLER', 'DODGE', 'JEEP', 'RAM') THEN 'CDJR'
            WHEN Make IN ('FORD', 'LINCOLN') THEN 'Ford'
            WHEN Make IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC') THEN 'GM'
            WHEN Make IN ('HYUNDAI', 'GENESIS') THEN 'Hyundai'
            WHEN Make IN ('JAGUAR', 'LAND ROVER') THEN 'JLR'
            WHEN Make IS NULL THEN 'Others'
            WHEN Make IN (
                'Acura','Honda','INFINITI','Mazda','Nissan','Subaru',
                'Toyota','Volkswagen','Volvo',
                'Aston Martin','Audi','BMW','Bentley','Lexus',
                'Mercedes-Benz','MINI','Porsche'
            ) THEN Make
            ELSE 'Others'
        END AS BrandGroup,

        CASE
            WHEN Make IN ('CHRYSLER', 'DODGE', 'JEEP', 'RAM', 'FORD', 'LINCOLN', 'BUICK', 'CADILLAC', 'CHEVROLET', 'GMC') THEN 'Domestic'
            WHEN Make IN ('HYUNDAI', 'GENESIS', 'ACURA', 'HONDA', 'INFINITI', 'MAZDA', 'NISSAN', 'SUBARU', 'TOYOTA', 'VOLKSWAGEN', 'VOLVO') THEN 'Import'
            WHEN Make IN ('JAGUAR', 'LAND ROVER', 'ASTON MARTIN', 'AUDI', 'BMW', 'BENTLEY', 'LEXUS', 'MERCEDES-BENZ', 'MINI', 'PORSCHE') THEN 'Premium Luxury'
            WHEN Make IS NULL THEN 'Others'
            ELSE 'Others'
        END AS Category,

        InventoryCount
    FROM NDDUsers.vInventoryMonthEnd
    WHERE 
        Department = '300' 
        AND AccountingMonth = '{Current_accounting_month}'
        AND Status <> 'G'
),

BrandSums AS (
    SELECT
        AccountingMonth,
        BrandGroup AS BrandName,
        Category,
        SUM(InventoryCount) AS INV
    FROM MappedData
    GROUP BY AccountingMonth, BrandGroup, Category
),

-- Step 3: Sales Source (Brand, Category, SoldCount)
SalesData AS (
    SELECT
        Accountingmonth,
        CASE
            WHEN VehicleMakeName IN ('CHRYSLER', 'DODGE', 'JEEP', 'RAM') THEN 'CDJR'
            WHEN VehicleMakeName IN ('FORD', 'LINCOLN') THEN 'Ford'
            WHEN VehicleMakeName IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC') THEN 'GM'
            WHEN VehicleMakeName IN ('HYUNDAI', 'GENESIS') THEN 'Hyundai'
            WHEN VehicleMakeName IN ('JAGUAR', 'LAND ROVER') THEN 'JLR'
            WHEN VehicleMakeName IN ('ACURA') THEN 'Acura'
            WHEN VehicleMakeName IN ('HONDA') THEN 'Honda'
            WHEN VehicleMakeName IN ('INFINITI') THEN 'INFINITI'
            WHEN VehicleMakeName IN ('MAZDA') THEN 'Mazda'
            WHEN VehicleMakeName IN ('NISSAN') THEN 'Nissan'
            WHEN VehicleMakeName IN ('SUBARU') THEN 'Subaru'
            WHEN VehicleMakeName IN ('TOYOTA') THEN 'Toyota'
            WHEN VehicleMakeName IN ('VOLKSWAGEN') THEN 'Volkswagen'
            WHEN VehicleMakeName IN ('VOLVO') THEN 'Volvo'
            WHEN VehicleMakeName IN ('ASTON MARTIN') THEN 'Aston Martin'
            WHEN VehicleMakeName IN ('AUDI') THEN 'Audi'
            WHEN VehicleMakeName IN ('BMW') THEN 'BMW'
            WHEN VehicleMakeName IN ('BENTLEY') THEN 'Bentley'
            WHEN VehicleMakeName IN ('LEXUS') THEN 'Lexus'
            WHEN VehicleMakeName IN ('MERCEDES-BENZ') THEN 'Mercedes-Benz'
            WHEN VehicleMakeName IN ('MINI') THEN 'MINI'
            WHEN VehicleMakeName IN ('PORSCHE') THEN 'Porsche'
            WHEN VehicleMakeName IS NULL THEN 'Others'
            ELSE 'Others'
        END AS BrandName,

        CASE
            WHEN VehicleMakeName IN ('CHRYSLER', 'DODGE', 'JEEP', 'RAM', 'FORD', 'LINCOLN', 'BUICK', 'CADILLAC', 'CHEVROLET', 'GMC') THEN 'Domestic'
            WHEN VehicleMakeName IN ('HYUNDAI', 'GENESIS', 'ACURA', 'HONDA', 'INFINITI', 'MAZDA', 'NISSAN', 'SUBARU', 'TOYOTA', 'VOLKSWAGEN', 'VOLVO') THEN 'Import'
            WHEN VehicleMakeName IN ('JAGUAR', 'LAND ROVER', 'ASTON MARTIN', 'AUDI', 'BMW', 'BENTLEY', 'LEXUS', 'MERCEDES-BENZ', 'MINI', 'PORSCHE') THEN 'Premium Luxury'
            WHEN VehicleMakeName IS NULL THEN 'Others'
            ELSE 'Others'
        END AS Category,

        SUM(vehiclesoldcount) AS SoldCount
    FROM NDDUsers.vSalesDetail_Vehicle
    WHERE 
        departmentname = 'New'
        AND accountingmonth = '{Current_accounting_month}'
    GROUP BY 
        Accountingmonth,
        VehicleMakeName
),

SalesSums AS (
    SELECT
        AccountingMonth,
        BrandName,
        Category,
        SUM(SoldCount) AS SoldCount
    FROM SalesData
    GROUP BY AccountingMonth, BrandName, Category
)

-- Step 4: Final join → BrandList + Dates x BrandSums + SalesSums
SELECT
    B.BrandName as Brand,
    COALESCE(SalesSums.Category, BrandSums.Category, 'Others') AS Category,
    MappedDate.AccountingMonth as Date,
    COALESCE(BrandSums.INV, 0) AS INV,
    COALESCE(SalesSums.SoldCount, 0) AS Sales
FROM 
    BrandList B
CROSS JOIN 
    (SELECT DISTINCT AccountingMonth FROM MappedData) MappedDate
LEFT JOIN 
    BrandSums ON B.BrandName = BrandSums.BrandName AND BrandSums.AccountingMonth = MappedDate.AccountingMonth
LEFT JOIN 
    SalesSums ON B.BrandName = SalesSums.BrandName AND SalesSums.AccountingMonth = MappedDate.AccountingMonth
ORDER BY 
    MappedDate.AccountingMonth,
    CASE WHEN B.BrandName = 'Others' THEN 999 ELSE 1 END,
    B.BrandName;
"""

    New_Sales_query = f"""
WITH Sales_By_Period AS (
    SELECT
        p.EndDate as Date,
        CASE
            WHEN VehicleMakeName IN ('CHRYSLER','DODGE','JEEP','RAM') THEN 'CDJR'
            WHEN VehicleMakeName IN ('FORD','LINCOLN') THEN 'Ford'
            WHEN VehicleMakeName IN ('BUICK','CADILLAC','CHEVROLET','GMC') THEN 'GM'
            WHEN VehicleMakeName IN ('HYUNDAI','GENESIS') THEN 'Hyundai'
            WHEN VehicleMakeName IN ('JAGUAR','LAND ROVER') THEN 'JLR'
            WHEN VehicleMakeName IN ('ACURA') THEN 'Acura'
            WHEN VehicleMakeName IN ('HONDA') THEN 'Honda'
            WHEN VehicleMakeName IN ('INFINITI') THEN 'INFINITI'
            WHEN VehicleMakeName IN ('MAZDA') THEN 'Mazda'
            WHEN VehicleMakeName IN ('NISSAN') THEN 'Nissan'
            WHEN VehicleMakeName IN ('SUBARU') THEN 'Subaru'
            WHEN VehicleMakeName IN ('TOYOTA') THEN 'Toyota'
            WHEN VehicleMakeName IN ('VOLKSWAGEN') THEN 'Volkswagen'
            WHEN VehicleMakeName IN ('VOLVO') THEN 'Volvo'
            WHEN VehicleMakeName IN ('ASTON MARTIN') THEN 'Aston Martin'
            WHEN VehicleMakeName IN ('AUDI') THEN 'Audi'
            WHEN VehicleMakeName IN ('BMW') THEN 'BMW'
            WHEN VehicleMakeName IN ('BENTLEY') THEN 'Bentley'
            WHEN VehicleMakeName IN ('LEXUS') THEN 'Lexus'
            WHEN VehicleMakeName IN ('MERCEDES-BENZ') THEN 'Mercedes-Benz'
            WHEN VehicleMakeName IN ('MINI') THEN 'MINI'
            WHEN VehicleMakeName IN ('PORSCHE') THEN 'Porsche'
            ELSE 'Others'
        END AS Brand,
 
        CASE
            WHEN VehicleMakeName IN ('CHRYSLER','DODGE','JEEP','RAM','FORD','LINCOLN','BUICK','CADILLAC','CHEVROLET','GMC') THEN 'Domestic'
            WHEN VehicleMakeName IN ('HYUNDAI','GENESIS','ACURA','HONDA','INFINITI','MAZDA','NISSAN','SUBARU','TOYOTA','VOLKSWAGEN','VOLVO') THEN 'Import'
            WHEN VehicleMakeName IN ('JAGUAR','LAND ROVER','ASTON MARTIN','AUDI','BMW','BENTLEY','LEXUS','MERCEDES-BENZ','MINI','PORSCHE') THEN 'Premium Luxury'
            ELSE 'Others'
        END AS Category,
 
        SUM(VehicleSoldCount) AS SoldCount,
        SUM(CASE WHEN RecordSource = 'Accounting' THEN BaseRetail ELSE 0 END) AS BaseGross,
        SUM(CASE WHEN RecordSource = 'Accounting' THEN Incentives ELSE 0 END) AS IncentiveGross,
        SUM(CASE WHEN RecordSource = 'Accounting' THEN figross ELSE 0 END) AS CFSGross,
        SUM(CASE WHEN RecordSource = 'Accounting' THEN OVI ELSE 0 END) AS OVIGross,
        SUM(CASE WHEN RecordSource = 'Accounting' THEN VehicleSoldCount ELSE 0 END) AS Accounting_Sold_count
 
    FROM (
        SELECT 'PM' AS PeriodType,
                DATEFROMPARTS(YEAR(DATEADD(MONTH, -1, GETDATE())), MONTH(DATEADD(MONTH, -1, GETDATE())), 1) AS StartDate,
                DATEADD(MONTH, -1, DATEADD(DAY, -1, GETDATE())) AS EndDate
 
        UNION ALL
 
       SELECT
            'IN-BETWEEN-1' AS PeriodType,
            DATEFROMPARTS(YEAR(DATEADD(MONTH, -1, GETDATE())), MONTH(DATEADD(MONTH, -1, GETDATE())), 1) AS StartDate,
            DATEADD(MONTH, -1, DATEADD(DAY, +7, GETDATE())) AS EndDate
 
        UNION ALL
 
        SELECT 'CM',
               DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1),
               DATEADD(DAY, -1, GETDATE())
 
        UNION ALL
 
        SELECT 'IN-BETWEEN-2',
               DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1),
               DATEADD(DAY, -7, GETDATE())
    ) p
 
    JOIN NDDUsers.vSalesDetail_Vehicle v
        ON v.ContractDate BETWEEN p.StartDate AND p.EndDate
    WHERE v.DepartmentName = 'New'
    GROUP BY p.PeriodType, p.StartDate, p.EndDate,
             CASE
                 WHEN VehicleMakeName IN ('CHRYSLER','DODGE','JEEP','RAM') THEN 'CDJR'
                 WHEN VehicleMakeName IN ('FORD','LINCOLN') THEN 'Ford'
                 WHEN VehicleMakeName IN ('BUICK','CADILLAC','CHEVROLET','GMC') THEN 'GM'
                 WHEN VehicleMakeName IN ('HYUNDAI','GENESIS') THEN 'Hyundai'
                 WHEN VehicleMakeName IN ('JAGUAR','LAND ROVER') THEN 'JLR'
                 WHEN VehicleMakeName IN ('ACURA') THEN 'Acura'
                 WHEN VehicleMakeName IN ('HONDA') THEN 'Honda'
                 WHEN VehicleMakeName IN ('INFINITI') THEN 'INFINITI'
                 WHEN VehicleMakeName IN ('MAZDA') THEN 'Mazda'
                 WHEN VehicleMakeName IN ('NISSAN') THEN 'Nissan'
                 WHEN VehicleMakeName IN ('SUBARU') THEN 'Subaru'
                 WHEN VehicleMakeName IN ('TOYOTA') THEN 'Toyota'
                 WHEN VehicleMakeName IN ('VOLKSWAGEN') THEN 'Volkswagen'
                 WHEN VehicleMakeName IN ('VOLVO') THEN 'Volvo'
                 WHEN VehicleMakeName IN ('ASTON MARTIN') THEN 'Aston Martin'
                 WHEN VehicleMakeName IN ('AUDI') THEN 'Audi'
                 WHEN VehicleMakeName IN ('BMW') THEN 'BMW'
                 WHEN VehicleMakeName IN ('BENTLEY') THEN 'Bentley'
                 WHEN VehicleMakeName IN ('LEXUS') THEN 'Lexus'
                 WHEN VehicleMakeName IN ('MERCEDES-BENZ') THEN 'Mercedes-Benz'
                 WHEN VehicleMakeName IN ('MINI') THEN 'MINI'
                 WHEN VehicleMakeName IN ('PORSCHE') THEN 'Porsche'
                 ELSE 'Others'
             END,
             CASE
                 WHEN VehicleMakeName IN ('CHRYSLER','DODGE','JEEP','RAM','FORD','LINCOLN','BUICK','CADILLAC','CHEVROLET','GMC') THEN 'Domestic'
                 WHEN VehicleMakeName IN ('HYUNDAI','GENESIS','ACURA','HONDA','INFINITI','MAZDA','NISSAN','SUBARU','TOYOTA','VOLKSWAGEN','VOLVO') THEN 'Import'
                 WHEN VehicleMakeName IN ('JAGUAR','LAND ROVER','ASTON MARTIN','AUDI','BMW','BENTLEY','LEXUS','MERCEDES-BENZ','MINI','PORSCHE') THEN 'Premium Luxury'
                 ELSE 'Others'
             END
)
 
SELECT *
FROM Sales_By_Period
WHERE SoldCount + BaseGross + IncentiveGross + CFSGross + OVIGross <> 0
ORDER BY Date, Brand

"""

    New_Pricing_query = f"""
SELECT 
    Accountingmonth,
    VIN,
    InventoryCount,
    StoreName,
    Make,
    DaysInInventoryAN,
    MSRP,
    PriceTier_93,
    PriceTier_95,
    PriceTier_99,

    -- Is_Discounted
    CASE 
        WHEN (
             (PriceTier_93 IS NOT NULL AND PriceTier_93 < MSRP)
             OR
             (PriceTier_95 IS NOT NULL AND PriceTier_95 < MSRP)
        ) THEN 1 ELSE 0
    END AS Is_Discounted,

    -- AN Discount: MSRP - Tier 93
    CASE 
        WHEN PriceTier_93 IS NOT NULL
        THEN MSRP - PriceTier_93
        ELSE NULL
    END AS AN_Discount,

    -- E-Com Discount: MSRP - Tier 95
    CASE 
        WHEN PriceTier_95 IS NOT NULL
        THEN MSRP - PriceTier_95
        ELSE NULL
    END AS Ecom_Discount,

    -- Incentives: Tier99 - Tier95 (negative)
    CASE 
        WHEN PriceTier_95 IS NOT NULL OR PriceTier_99 IS NOT NULL
        THEN PriceTier_95 - PriceTier_99
        ELSE NULL
    END AS Incentives,

    -- Total Discount: E-Com Discount + Incentives
    CASE 
        WHEN PriceTier_95 IS NOT NULL OR PriceTier_99 IS NOT NULL
        THEN (MSRP - PriceTier_95) + (PriceTier_95 - PriceTier_99)
        ELSE NULL
    END AS Total_Discount,

    -- Is_Aged: 1 if DaysInInventoryAN > 90, else 0
    CASE 
        WHEN DaysInInventoryAN > 90 THEN 1 ELSE 0
    END AS Is_Aged,

    -- Aged_InventoryCount: InventoryCount if Is_Aged = 1, else 0
    CASE 
        WHEN DaysInInventoryAN > 90 THEN InventoryCount ELSE 0
    END AS Aged_InventoryCount,

    -- Aged_And_Discounted: InventoryCount if Is_Aged = 1 AND Is_Discounted = 1, else 0
    CASE 
        WHEN DaysInInventoryAN > 90 AND (
            (PriceTier_93 IS NOT NULL AND PriceTier_93 < MSRP)
            OR
            (PriceTier_95 IS NOT NULL AND PriceTier_95 < MSRP)
        ) THEN InventoryCount ELSE 0
    END AS Aged_And_Discounted,

    -- Aged_Not_Discounted: InventoryCount if Is_Aged = 1 AND Is_Discounted = 0, else 0
    CASE 
        WHEN DaysInInventoryAN > 90 AND NOT (
            (PriceTier_93 IS NOT NULL AND PriceTier_93 < MSRP)
            OR
            (PriceTier_95 IS NOT NULL AND PriceTier_95 < MSRP)
        ) THEN InventoryCount ELSE 0
    END AS Aged_Not_Discounted,
    CASE
        WHEN Make IN ('CHRYSLER', 'DODGE', 'JEEP', 'RAM') THEN 'CDJR'
        WHEN Make IN ('FORD', 'LINCOLN') THEN 'Ford'
        WHEN Make IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC') THEN 'GM'
        WHEN Make IN ('HYUNDAI', 'GENESIS') THEN 'Hyundai'
        WHEN Make IN ('JAGUAR', 'LAND ROVER') THEN 'JLR'
        WHEN Make IN ('ACURA') THEN 'Acura'
        WHEN Make IN ('HONDA') THEN 'Honda'
        WHEN Make IN ('INFINITI') THEN 'INFINITI'
        WHEN Make IN ('MAZDA') THEN 'Mazda'
        WHEN Make IN ('NISSAN') THEN 'Nissan'
        WHEN Make IN ('SUBARU') THEN 'Subaru'
        WHEN Make IN ('TOYOTA') THEN 'Toyota'
        WHEN Make IN ('VOLKSWAGEN') THEN 'Volkswagen'
        WHEN Make IN ('VOLVO') THEN 'Volvo'
        WHEN Make IN ('ASTON MARTIN') THEN 'Aston Martin'
        WHEN Make IN ('AUDI') THEN 'Audi'
        WHEN Make IN ('BMW') THEN 'BMW'
        WHEN Make IN ('BENTLEY') THEN 'Bentley'
        WHEN Make IN ('LEXUS') THEN 'Lexus'
        WHEN Make IN ('MERCEDES-BENZ') THEN 'Mercedes-Benz'
        WHEN Make IN ('MINI') THEN 'MINI'
        WHEN Make IN ('PORSCHE') THEN 'Porsche'
        WHEN Make IS NULL THEN 'Others'
        ELSE 'Others'
    END AS Brand,
    CASE
        WHEN Make IN ('CHRYSLER', 'DODGE', 'JEEP', 'RAM', 'FORD', 'LINCOLN', 'BUICK', 'CADILLAC', 'CHEVROLET', 'GMC') THEN 'Domestic'
        WHEN Make IN ('HYUNDAI', 'GENESIS', 'ACURA', 'HONDA', 'INFINITI', 'MAZDA', 'NISSAN', 'SUBARU', 'TOYOTA', 'VOLKSWAGEN', 'VOLVO') THEN 'Import'
        WHEN Make IN ('JAGUAR', 'LAND ROVER', 'ASTON MARTIN', 'AUDI', 'BMW', 'BENTLEY', 'LEXUS', 'MERCEDES-BENZ', 'MINI', 'PORSCHE') THEN 'Premium Luxury'
        WHEN Make IS NULL THEN 'Others'
        ELSE 'Others'
    END AS Category

FROM NDDUsers.vInventoryMonthEnd
WHERE AccountingMonth = '{Current_accounting_month}'
      AND regionName <> 'AND Corporate Management'
      AND Department = '300'
      AND marketName NOT IN ('market 97', 'market 98')
      AND MSRP > 0
      AND ValidForAN = 1
      AND ValidForPricing = 1
      AND InventoryCount <> 0
"""

    New_Allocation_query = f"""
    SELECT --ParentBrand
    CASE
        WHEN Make IN ('CHRYSLER', 'DODGE','JEEP','RAM')
        THEN 'CDJR'
        WHEN Make in ('FORD', 'LINCOLN')
        THEN 'FORD'
        WHEN Make IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC')
        THEN 'GM'
        WHEN Make in ('HYUNDAI', 'GENESIS')
        THEN 'HYUNDAI'
        WHEN Make in ('JAGUAR', 'LAND ROVER')
        THEN 'JLR'
        Else Make
        END AS Brand,
    CASE
        WHEN Make IN ('CHRYSLER','DODGE','JEEP','RAM','FORD','LINCOLN','BUICK','CADILLAC','CHEVROLET','GMC')
        THEN 'Domestic'
        WHEN Make IN ('ACURA','HONDA','GENESIS','HYUNDAI','INFINITI','MAZDA', 'NISSAN','SUBARU','TOYOTA','VOLKSWAGEN', 'VOLVO')
        THEN 'Import'
        WHEN Make IN ('AUDI','BMW','JAGUAR','LAND ROVER','LEXUS','MERCEDES-BENZ','MINI')
        THEN 'Luxury'
        END AS Segment,
    SUM(EarnedM1) AS EarnedM1,
    SUM(CommitM1) AS CommitM1

    FROM(

    SELECT --PIVOTQUERY
    Make,
    SUM(CASE WHEN AccountingMonth = 'Earned_M1' THEN QTY ELSE 0 END) AS EarnedM1,
    SUM(CASE WHEN AccountingMonth = 'Commit_M1' THEN QTY ELSE 0 END) AS CommitM1

    FROM(

    SELECT 
        --DC.[Year], 
        --DC.[Month], 
        --D.[DealerCD] AS Hyperion,
        B.[BrandCD] AS Make,
        --BM.[BrandModelCD] AS Model, 
        --DC.[Mth],
            CASE
                WHEN DC.[ColumnCD] IN ('TCCM1', 'TCC2_M1','TCC3_M1')
                THEN 'Commit_M1'
                WHEN DC.[ColumnCD] IN ('EAM1', 'EA2_M1', 'EA3_M1')
                THEN 'Earned_M1'
                ELSE 'ERROR'
                END AS AccountingMonth,
            SUM(DC.[EnteredValue]) AS QTY

    FROM
    (
    SELECT
    DC.[Year], DC.[Month], DC.[DealerVehicleID], DC.[Mth], DC.[ColumnCD], DC.[EnteredValue], DC.[StatusID]
    FROM
    [SSPRv3].[dbo].[DealerCommits] AS DC
    WHERE
    DC.[Year] in (  Year(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)) )  
    AND
    --DC.[Month] in ( Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)) )
    DC.[Month] = MONTH(DATEADD(MONTH, 0, GETDATE()))

    UNION

    SELECT
    DF.[Year], DF.[Month], DF.[DealerVehicleID], DF.[Mth], DF.[ColumnCD], DF.[EnteredValue], DF.[StatusID]
    FROM
    [SSPRv3].[dbo].[DealerForecasts] AS DF
    WHERE
    DF.[Year] in (  Year(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)) ) 
    AND
    --DF.[Month] in ( Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)) )
    DF.[Month] = MONTH(DATEADD(MONTH, 0, GETDATE()))
    ) AS DC

    INNER JOIN
        [SSPRv3].[dbo].[DealerVehicles] AS DV ON DC.[DealerVehicleID] = DV.[DealerVehicleID]
    INNER JOIN
        [SSPRv3].[dbo].[Brands] AS B ON DV.[BrandID] = B.[BrandID]
    INNER JOIN
        [SSPRv3].[dbo].[BrandModels] AS BM ON DV.[BrandModelID] = BM.[BrandModelID] AND B.[BrandID] = BM.[BrandID]
    INNER JOIN
        [SSPRv3].[dbo].[Dealers] AS D ON DV.[DealerCD] = D.[DealerCD]
    INNER JOIN
        [SSPRv3].[dbo].[DealerVehicleCores] AS DVC ON DC.[DealerVehicleID] = DVC.[DealerVehicleID] AND DV.[DealerVehicleID] = DVC.[DealerVehicleID] AND DC.[Year] = DVC.[Year] AND DC.[Month] = DVC.[Month]
    INNER JOIN
        [SSPRv3].[dbo].[BrandModelDisplaySorts] AS BMDS ON BM.[BrandModelID] = BMDS.[BrandModelID]
    WHERE 1=1 
        --DV.[DealerCD] = @DealerCD
    AND
        DC.[Year] in (  Year(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)) ) 
    AND
        --DC.[Month] in ( Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)) )
        DC.[Month] = MONTH(DATEADD(MONTH, 0, GETDATE()))


    --AND BM.[BrandModelCD] = @BrandModelCD
    and DC.[ColumnCD] in ('EAM1' , 'TCCM1', 'EA2_M1' , 'TCC2_M1' , 'EA3_M1' , 'TCC3_M1')

    AND DC.[EnteredValue] <> '0'

    GROUP BY
    --D.[DealerCD],
        B.[BrandCD],
        --BM.[BrandModelCD] AS Model, 
        --DC.[Mth],
            CASE
                WHEN DC.[ColumnCD] IN ('TCCM1', 'TCC2_M1','TCC3_M1')
                THEN 'Commit_M1'
                WHEN DC.[ColumnCD] IN ('EAM1', 'EA2_M1', 'EA3_M1')
                THEN 'Earned_M1'
                ELSE 'ERROR'
                END

    )AS PIVOTQUERY

    GROUP BY
    Make

    )AS ParentBrand

    GROUP BY
    CASE
        WHEN Make IN ('CHRYSLER', 'DODGE','JEEP','RAM')
        THEN 'CDJR'
        WHEN Make in ('FORD', 'LINCOLN')
        THEN 'FORD'
        WHEN Make IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC')
        THEN 'GM'
        WHEN Make in ('HYUNDAI', 'GENESIS')
        THEN 'HYUNDAI'
        WHEN Make in ('JAGUAR', 'LAND ROVER')
        THEN 'JLR'
        Else Make
        END,
    CASE
        WHEN Make IN ('CHRYSLER','DODGE','JEEP','RAM','FORD','LINCOLN','BUICK','CADILLAC','CHEVROLET','GMC')
        THEN 'Domestic'
        WHEN Make IN ('ACURA','HONDA','GENESIS','HYUNDAI','INFINITI','MAZDA', 'NISSAN','SUBARU','TOYOTA','VOLKSWAGEN', 'VOLVO')
        THEN 'Import'
        WHEN Make IN ('AUDI','BMW','JAGUAR','LAND ROVER','LEXUS','MERCEDES-BENZ','MINI')
        THEN 'Luxury'
        END"""

    # right click on connection and go to properties to find the server name then select the database its looking at
    try:
        with pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=nddprddb01,48155;'
                'DATABASE=NDD_ADP_RAW;'
                'Trusted_Connection=yes;'
        ) as conn:
            Used_Inventory_df = pd.read_sql(Used_Inventory_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in Used_Inventory_df in {elapsed_time:.2f} seconds")
            Used_Sales_df = pd.read_sql(Used_sales_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in Used_Sales_df in {elapsed_time:.2f} seconds")
            # Used_Pricing_df = pd.read_sql(Used_Pricing_query, conn)           
            # end_time = time.time()
            # elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in Used_Pricing_df in {elapsed_time:.2f} seconds")
            New_Inventory_df = pd.read_sql(New_Inventory_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in New_Inventory_df in {elapsed_time:.2f} seconds")
            New_Sales_df = pd.read_sql(New_Sales_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in New_Sales_df in {elapsed_time:.2f} seconds")
            New_Pricing_df = pd.read_sql(New_Pricing_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in New_Pricing_df in {elapsed_time:.2f} seconds")

    except Exception as e:
        print("❌ Connection failed:", e)


    # right click on connection and go to properties to find the server name then select the database its looking at
    try:
        with pyodbc.connect(
                r'DRIVER={ODBC Driver 17 for SQL Server};'
                r'SERVER=BAPRDDB01\BAPRD,49174;'
                r'DATABASE=SSPRv3;'
                r'Trusted_Connection=yes;'
        ) as conn:
            New_Allocation_df = pd.read_sql(New_Allocation_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in New_Allocation_df in {elapsed_time:.2f} seconds")

    except Exception as e:
        print("❌ Connection failed:", e)

################################################################################################################################
    '''DROP IN SQL QUERIES INTO EXCEL FILE AND RUN MACRO FOR UPDATE'''
################################################################################################################################

    # Formulas for Used_Inventory
    # 1. Franchise formula
    Used_Inventory_df['Franchise_Type'] = np.where(
        Used_Inventory_df['StoreName'].str.startswith('AutoNation USA'),
        'Non-Franchise',
        'Franchise'
    )
    # 2. UsedBuckets Formula (filter age bucket and create >90 age bucket as well as process N/As to 0:15 Agebucket)
    # First if the age buckets that are greater than 90 days, create <90 days otherwise use the agebucket
    Used_Inventory_df['UsedBuckets_Formula'] = np.where(
        Used_Inventory_df['AgeBucket'].str.startswith('06:') | 
        Used_Inventory_df['AgeBucket'].str.startswith('07:'),
        "06: >90", Used_Inventory_df['AgeBucket']
    )
    # Next if the age bucket is categorized as N/A, move to the 0:15 age bucket, otherwise maintain the prior bucket formula
    Used_Inventory_df['UsedBuckets_Formula'] = np.where(
        Used_Inventory_df['AgeBucket'].str.startswith('N/A'),
        "01: 0-15", Used_Inventory_df['UsedBuckets_Formula']
    )

    # 3. Standardize the Sources
    Used_Inventory_df['InventorySource'] = np.where(
        Used_Inventory_df['InventorySource'].isin(['LEASE RETURN', 'Other', 'SERVICE LOANER']),
        "Others", Used_Inventory_df['InventorySource']
    )


    Used_Inventory_df['InventorySource'] = np.where(
        Used_Inventory_df['InventorySource'].isin(['TRADE-IN']),
        "Trade", Used_Inventory_df['InventorySource']
    )

    Used_Inventory_df['InventorySource'] = np.where(
        Used_Inventory_df['InventorySource'] == 'AUCTION/RENTAL',
        "Auction", Used_Inventory_df['InventorySource']
    )

    # Get today's date
    today = datetime.today()
    safe_date_str = today.strftime(r"%Y-%m-%d")  # '2025-05-27'

    # Add current date to files
    Used_Inventory_df['As_of'] = today
    Used_Sales_df['As_of'] = today

    # Open file and process macro/Sql
    Weekly_File = r'W:\Corporate\Inventory\Weekly Reporting Package\PowerBI\New_and_Used_Data.xlsm'
    Weekly_File_Historical = fr'W:\Corporate\Inventory\Weekly Reporting Package\PowerBI\Historical\New_and_Used_Data {safe_date_str}.xlsm'

    # Save historical copy of data
    shutil.copy(Weekly_File, Weekly_File_Historical)
    print(f"Saved historical file as {Weekly_File_Historical}")

    # Update Excel File
    app = xw.App(visible=True)
    app.display_alerts = True # Optional: suppress Excel alerts
    app.screen_updating = True # Optional: improve performance 
    Weekly_File_wb = app.books.open(Weekly_File)


    '''Update New Allocation'''
    New_Allocation_Tab = Weekly_File_wb.sheets['New_Allocation']
    # Read existing data into a DataFrame
    data_range = New_Allocation_Tab.range('A1').expand('down').resize(None, 5)
    existing_data = data_range.options(pd.DataFrame, header=1).value

    # Find the earliest date
    existing_data['Date'] = pd.to_datetime(existing_data['Date'])
    earliest_date = existing_data['Date'].min()

    # Ensure RN is a column, not an index
    existing_data = existing_data.reset_index()
    New_Allocation_df = New_Allocation_df.reset_index()

    # Filter out rows with the earliest date
    filtered_existing = existing_data[existing_data['Date'] > earliest_date]

    # Add Current Date to current Allocation Dataset
    New_Allocation_df['Date'] = pd.to_datetime(today)

    # Append new data
    combined_data = pd.concat([filtered_existing, New_Allocation_df], ignore_index=True) 

    # Replace accountingmonth date with todays date
    combined_data['Date'] = pd.to_datetime(combined_data['Date'])
    combined_data['Date'] = np.where(
        combined_data['Date'] == Current_accounting_month, 
        pd.to_datetime(today), 
        combined_data['Date'])
    
    # Drop extra index
    combined_data = combined_data.drop('index', axis=1)
    # combined_data = combined_data.drop('index', axis=1)
    combined_data['Date'] = pd.to_datetime(combined_data['Date'])
    combined_data['Date'] = combined_data['Date'].dt.strftime('%m/%d/%Y')

    # Clear and write updated data
    New_Allocation_Tab.clear_contents()
    New_Allocation_Tab.range('A1').options(index=False).value = combined_data
    
    '''Update New Sales'''
    # Clear and write updated data
    New_Sales_tab = Weekly_File_wb.sheets['New_Sales']
    New_Sales_tab.clear_contents()
    New_Sales_tab.range('A1').options(index=False).value = New_Sales_df


    '''Update New Inventory'''
    New_Inventory_tab = Weekly_File_wb.sheets['New_Inventory']

    # Read existing data into a DataFrame
    data_range = New_Inventory_tab.range('A1').expand('down').resize(None, 5)
    existing_data = data_range.options(pd.DataFrame, header=1).value

    # Find the earliest date
    existing_data['Date'] = pd.to_datetime(existing_data['Date'])
    earliest_date = existing_data['Date'].min()

    # Ensure RN is a column, not an index
    existing_data = existing_data.reset_index()
    New_Inventory_df = New_Inventory_df.reset_index()

    # Filter out rows with the earliest date
    filtered_existing = existing_data[existing_data['Date'] > earliest_date]

    # Append new data
    New_Inventory_df = New_Inventory_df.rename(columns={'Accountingmonth': 'Date'})
    New_Inventory_df['Date'] = pd.to_datetime(New_Inventory_df['Date'])
    combined_data = pd.concat([filtered_existing, New_Inventory_df], ignore_index=True)  

    # Replace accountingmonth date with todays date
    combined_data['Date'] = pd.to_datetime(combined_data['Date'])
    combined_data['Date'] = np.where(
        combined_data['Date'] == Current_accounting_month, 
        pd.to_datetime(today), 
        combined_data['Date'])
    
    # Drop extra index
    combined_data = combined_data.drop('index', axis=1)
    combined_data['Date'] = pd.to_datetime(combined_data['Date'])
    combined_data['Date'] = combined_data['Date'].dt.strftime('%m/%d/%Y')

    # Clear and write updated data
    New_Inventory_tab.clear_contents()
    New_Inventory_tab.range('A1').options(index=False).value = combined_data


    '''Update New Pricing'''
    New_Pricing_tab = Weekly_File_wb.sheets['New_Pricing']

    # Read existing data into a DataFrame
    data_range = New_Pricing_tab.range('A1').expand('down').resize(None, 21)
    existing_data = data_range.options(pd.DataFrame, header=1, index=False).value

    # Find the earliest date
    existing_data['Accountingmonth'] = pd.to_datetime(existing_data['Accountingmonth'])
    earliest_date = existing_data['Accountingmonth'].min()

    # Ensure RN is a column, not an index
    existing_data = existing_data.reset_index()
    New_Pricing_df = New_Pricing_df.reset_index()

    # Filter out rows with the earliest date
    filtered_existing = existing_data[existing_data['Accountingmonth'] > earliest_date]
    
    # Append new data
    combined_data = pd.concat([filtered_existing, New_Pricing_df], ignore_index=True)  
    # combined_data = combined_data.drop('index', axis=1)

    # Replace accountingmonth date with todays date
    combined_data['Accountingmonth'] = pd.to_datetime(combined_data['Accountingmonth'])
    combined_data['Accountingmonth'] = np.where(
        combined_data['Accountingmonth'] == Current_accounting_month, 
        pd.to_datetime(today), 
        combined_data['Accountingmonth'])
    combined_data = combined_data.drop('index', axis=1)
    combined_data['Accountingmonth'] = pd.to_datetime(combined_data['Accountingmonth'])
    combined_data['Accountingmonth'] = combined_data['Accountingmonth'].dt.strftime('%m/%d/%Y')

    # Clear and write updated data
    New_Pricing_tab.clear_contents()
    New_Pricing_tab.range('A1').options(index=False).value = combined_data

    '''Update Used Inventory'''
    Used_Inventory_tab = Weekly_File_wb.sheets['Used_Inventory_Data']

    # Read existing data into a DataFrame
    data_range = Used_Inventory_tab.range('A1').expand('down').resize(None, 21)
    existing_data = data_range.options(pd.DataFrame, header=1).value

    # Find the earliest date
    existing_data['As_of'] = pd.to_datetime(existing_data['As_of'])
    earliest_date = existing_data['As_of'].min()

    # Ensure RN is a column, not an index
    existing_data = existing_data.reset_index()
    Used_Inventory_df = Used_Inventory_df.reset_index()

    # Filter out rows with the earliest date
    filtered_existing = existing_data[existing_data['As_of'] > earliest_date]

    # Append new data
    combined_data = pd.concat([filtered_existing, Used_Inventory_df], ignore_index=True)  
    combined_data = combined_data.drop('index', axis=1)


    # Clear and write updated data
    Used_Inventory_tab.clear_contents()
    Used_Inventory_tab.range('A1').options(index=False).value = combined_data

    '''Update Used Sales'''
    # Read existing data into a DataFrame
    Used_Sales_tab = Weekly_File_wb.sheets['Used_Sales_Data']
    data_range = Used_Sales_tab.range('A1').expand('down').resize(None, 36)
    existing_data = data_range.options(pd.DataFrame, header=1).value

    # Find the earliest date
    existing_data['As_of'] = pd.to_datetime(existing_data['As_of'])
    earliest_date = existing_data['As_of'].min()

    # Count unique dates
    # unique_dates = existing_data['As_of'].nunique()

    # If more than 8 unique dates (two months), remove rows with the earliest date
    # if unique_dates > 8:
    #     earliest_date = existing_data['As_of'].min()
    #     existing_data = existing_data[existing_data['As_of'] > earliest_date]

    # Filter out earliest Date
    earliest_date = existing_data['As_of'].min()
    existing_data = existing_data[existing_data['As_of'] > earliest_date]

    # Ensure RN is a column, not an index
    existing_data = existing_data.reset_index()
    Used_Sales_df = Used_Sales_df.reset_index(drop=True)

    # Append new data
    combined_data = pd.concat([existing_data, Used_Sales_df], ignore_index=True)
    
    # Sort by oldest date
    combined_data = combined_data.sort_values('As_of')

    # Clear and write updated data
    Used_Sales_tab.clear_contents()
    Used_Sales_tab.range('A1').options(index=False).value = combined_data

    Dynamic_Daily_Sales_File = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker\Dynamic_Daily_Sales.xlsb'    
    Daily_sales_wb = app.books.open(Dynamic_Daily_Sales_File,update_links=False)    
    Run_Macro = Weekly_File_wb.macro("Update_Targets")
    Run_Macro()    
    
    # Save and close the excel document
    Weekly_File_wb.save(Weekly_File)
    if Weekly_File_wb:
        Weekly_File_wb.close()
    if app:
        app.quit()

    # End Timer
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Script completed in {elapsed_time:.2f} seconds")

    return

#run function
if __name__ == '__main__':

    Process_Daily_Sales_File()    
    #Download_PWB()
    #Update_PWB_Data()
    Weekly_Data_Update()