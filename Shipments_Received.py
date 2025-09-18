# Imports
import xlwings as xw
import pandas as pd
import os
from datetime import datetime, timedelta
import shutil
import numpy as np
import gc
import glob
import pyodbc
import time
from dateutil.relativedelta import relativedelta
import warnings
warnings.filterwarnings("ignore", category=UserWarning, message=".*pandas only supports SQLAlchemy.*")

#main code
def Update():

    # Begin timer
    start_time = time.time()

    # Get current date
    today = datetime.today()

    # Check if today's day is less than 7
    # if today.day < 7:
    #     # Go back one month
    #     first_of_month = today.replace(day=1)
    #     last_month = first_of_month - timedelta(days=1)
    #     target_date = last_month
    # else: # otherwise process as normal
    #     target_date = today

    target_date = today

    # Format as M/D/YYYY (no leading zeros)
    current_day = f"{target_date.month}/{target_date.day}/{target_date.year}"
    print(current_day)

    # Get 12 months prior date
    twelve_months_date = today - relativedelta(months=12)

    # Format dates manually without leading zeros
    twelve_months_date = f"{twelve_months_date.month}/{twelve_months_date.day}/{twelve_months_date.year}"  # e.g., '5/23/2025'
    print(twelve_months_date)

    Inv_Query = f"""
    select AccountingMonth,
    Make,
    Model,
    sum (InventoryCount) as SumofInventoryCount, 
    status, 
    EntryDate, 
    StoreName


    from NDDUsers.vInventoryMonthEnd

    where 1=1
    and Department = '300'
    and balance > 0 
    and accountingmonth between '{twelve_months_date}' and '{current_day}'
    and InventoryCount > 0


    group by AccountingMonth,
    Make,
    Model,
    status, 
    EntryDate, 
    StoreName
    """
    Sales_Query = f"""
    SELECT Accountingmonth, 
    VehicleMakeName,
    Sum(VehicleSoldCount) AS 'SoldMTD',
    ExService,
    RecordSource, 
    StoreName

    FROM NDD_ADP_RAW.NDDUsers.vSalesDetail_Vehicle vSalesDetail_Vehicle

    WHERE 1=1
    and AccountingMonth between '{twelve_months_date}' and '{current_day}'
    AND DepartmentName='new'
    and VehicleSoldCount <>0

    Group by Accountingmonth, 
    ExService,
    VehicleMakeName,
    RecordSource, 
    StoreName


    """
    Email_Query1 = """
    SELECT ExService,Sum(VehicleSoldCount) AS 'soldmmtd'
    FROM NDD_ADP_RAW.NDDUsers.vSalesDetail_Vehicle vSalesDetail_Vehicle
    WHERE AccountingMonth=DATEADD(MONTH,DATEDIFF(MONTH,0,GETDATE()),0) AND DepartmentName='new'
    Group by ExService
    """
    Email_Query2 = """
    SELECT vInventoryMonthEnd.Status, Sum(vInventoryMonthEnd.InventoryCount) AS 'sum_inv_count', vInventoryMonthEnd.AccountingMonth
    FROM NDD_ADP_RAW.NDDUsers.vInventoryMonthEnd vInventoryMonthEnd
    WHERE (vInventoryMonthEnd.AccountingMonth=DATEADD(MONTH,DATEDIFF(MONTH,0,GETDATE()),0)) 
    AND (vInventoryMonthEnd.Balance<>$0) 
    AND (vInventoryMonthEnd.Department='300')
                                            
    GROUP BY vInventoryMonthEnd.Status, vInventoryMonthEnd.AccountingMonth
    ORDER BY vInventoryMonthEnd.Status
    """
    Pipeline_Query = """
    SELECT vOnOrderInfo_Dly_SnapShot.make, Count(vOnOrderInfo_Dly_SnapShot.order_number) , vOnOrderInfo_Dly_SnapShot.an_status, UPDATE_DATE, hyperion_id
    FROM NDD_ADP_RAW.NDDUsers.vOnOrderInfo_Dly_SnapShot vOnOrderInfo_Dly_SnapShot
    WHERE (vOnOrderInfo_Dly_SnapShot.update_date=dateadd(day,-1, cast(getdate() as date))) AND (vOnOrderInfo_Dly_SnapShot.an_status<>'ignore')
    GROUP BY vOnOrderInfo_Dly_SnapShot.make, vOnOrderInfo_Dly_SnapShot.an_status, UPDATE_DATE, hyperion_id
    """

    try:
        with pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=nddprddb01,48155;'
                'DATABASE=NDD_ADP_RAW;'
                'Trusted_Connection=yes;'
        ) as conn:
            Inv_df = pd.read_sql(Inv_Query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            print(f"Script read in Inv_df in {elapsed_time:.2f} seconds")
            end_time = time.time()
            elapsed_time = end_time - start_time
            Sales_df = pd.read_sql(Sales_Query, conn)
            print(f"Script read in Sales_df in {elapsed_time:.2f} seconds")
            Email1_df = pd.read_sql(Email_Query1, conn)
            Email2_df = pd.read_sql(Email_Query2, conn)
            Pipeline_df = pd.read_sql(Pipeline_Query, conn)
            print(Pipeline_df)
            end_time = time.time()
            elapsed_time = end_time - start_time
            print(f"Script read in email1, email2, and pipeline_df in {elapsed_time:.2f} seconds")           
    except Exception as e:
        print("❌ Connection failed:", e)


    SSPR_Query = """
SELECT
Year,
Month,
BrandCD,
SUM(CASE WHEN Type = 'TCCM' THEN EnteredValue ELSE 0 END) AS Commit_,
SUM(CASE WHEN Type = 'NAM' THEN EnteredValue ELSE 0 END) AS NetADD


FROM(


--2023
SELECT                  
       DC.[Year],
       DC.[Month],
       B.[BrandCD],                              
       CASE
	   WHEN DC.[ColumnCD] IN ('TCCM1', 'TCC2_M1', 'TCC3_M1')   
	   THEN 'TCCM'
	   WHEN DC.[ColumnCD] IN ('NAM1', 'NAM2', 'NAM3', 'NAM4', 'NAM5', 'NAM6')
	   THEN 'NAM'
	   ELSE 'ERROR'
	   END AS Type,
       DC.[EnteredValue],
       D.DealerID,
       D.DealerName
FROM                   
       [SSPRv3].[dbo].[DealerCommits] AS DC                          
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

WHERE 
DC.[Year] >= (YEAR(GETDATE()) - 1)                                                        
AND DC.[ColumnCD] in ('TCCM1' ,'TCC2_M1', 'TCC3_M1' , 'NAM1' , 'NAM2' , 'NAM3' , 'NAM4' , 'NAM5' , 'NAM6')                                                
                     
                           



) AS SQ

GROUP BY 

Year,
Month,
BrandCD
    """

    try:
        with pyodbc.connect(
                r'DRIVER={ODBC Driver 17 for SQL Server};'
                r'SERVER=BAPRDDB01\BAPRD,49174;'
                r'DATABASE=SSPRv3;'
                r'Trusted_Connection=yes;'
        ) as conn:
            SSPR_df = pd.read_sql(SSPR_Query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            print(f"Script read in SSPR_df in {elapsed_time:.2f} seconds")
            start_time = time.time()

    except Exception as e:
        print("❌ Connection failed:", e)


    # Excel files for processing in VBA
    Arrivals_file = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Shipments_Received\Arrivals_Comp.xlsm'
    New_EV_file = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Shipments_Received\New_EV_Inventory_Sales_Trend.xlsm'
    Project_Purple_Rain_file = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Shipments_Received\Project_Purple_Rain.xlsx'
    Dynamic_Daily_Sales = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker\Dynamic_Daily_Sales.xlsb'

    # Copy files, process based on current year/month
    todays_date = datetime.today()
    Current_Year = todays_date.year

    # Format the Current Month to e.g 05 May
    month_formatted = todays_date.strftime("%m %B")
    month_formatted_EV = todays_date.strftime("%m")
        
    Current_EV_Path = fr'W:\Corporate\Inventory\Reporting\EV Inventory Trend\NEW\{month_formatted_EV}' 
    Current_Arrivals_Comp_Path = fr'W:\Corporate\Inventory\Reporting\Arrivals\{Current_Year}\{month_formatted}'

    # Get all .xlsm files in the Arrivals folder
    Arrivals_files = glob.glob(os.path.join(Current_Arrivals_Comp_Path, "*.xlsm"))
    latest_arrivals_file = max(Arrivals_files, key=os.path.getmtime)
    print(f"Latest .xlsm file: {latest_arrivals_file}")

    # Get all real .xlsm files, exclude temp (~$) files
    EV_files = [
        f for f in glob.glob(os.path.join(Current_EV_Path, "*.xlsm"))
        if os.path.isfile(f) and not os.path.basename(f).startswith("~$")
    ]

    if EV_files:
        latest_EV_file = max(EV_files, key=os.path.getmtime)
        print(f"Latest .xlsm file for EV: {latest_EV_file}")

    # Send updated files to location to process
    shutil.copyfile(r'W:\Corporate\Inventory\Reporting\Arrivals\Project Purple Rain.xlsx', Project_Purple_Rain_file)
    shutil.copyfile(latest_EV_file, New_EV_file)

    # Run Macro
    app = xw.App(visible=True) 
    Arrivals_wb = app.books.open(Arrivals_file)
    wb2 = app.books.open(New_EV_file)
    wb3 = app.books.open(Project_Purple_Rain_file)
    wb4 = app.books.open(Dynamic_Daily_Sales)

    # Dump data into tab section to process'
    Sales_tab = Arrivals_wb.sheets['Sales']
    Sales_tab.range('A1:F50000').clear_contents()
    Sales_tab.range('A1').options(index=False).value = Sales_df

    Inv_tab = Arrivals_wb.sheets['Inv']
    Inv_tab.range('A1:G500000').clear_contents()
    Inv_tab.range('A1').options(index=False).value = Inv_df 

    Purchases_tab = Arrivals_wb.sheets['Purchases']
    Purchases_tab.range('B1:F500000').clear_contents()
    Purchases_tab.range('B1').options(index=False).value = SSPR_df

    Pipeline_tab = Arrivals_wb.sheets['Pipeline']
    Pipeline_tab.range('A1:G500000').clear_contents()    
    Pipeline_tab.range('A1').options(index=False).value = Pipeline_df

    Email_tab = Arrivals_wb.sheets['Email']
    Email_tab.range('E27:G29').clear_contents()
    Email_tab.range('E27').options(index=False).value = Email1_df
    Email_tab.range('I27:K36').clear_contents()
    Email_tab.range('I27').options(index=False).value = Email2_df
       


    Run_Macro = Arrivals_wb.macro("Refresh")
    Run_Macro()
    Arrivals_wb.save()

    # Save and close the excel document    
    if Arrivals_wb:
        Arrivals_wb.close()
    if app:
        app.quit()

    # Make a copy to Arrivals folder directory with today's date
    shutil.copyfile(Arrivals_file, fr'W:\Corporate\Inventory\Reporting\Arrivals\Arrivals Comp {todays_date}.xlsm')
    
    # Make a pdf copy as well (or do it in VBA?)


    return

#run function
if __name__ == '__main__':
    
    Update()