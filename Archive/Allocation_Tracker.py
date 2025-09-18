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
warnings.filterwarnings("ignore", category=UserWarning, message=".*pandas only supports SQLAlchemy.*")

# Begin timer
start_time = time.time()

################################################################################################################################
'''CONNECT SQL QUERIES TO PANDAS'''
################################################################################################################################

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

        return


def Allocation_Tracker_Update():

    # Run SQL queries using SQL Alchemy and dump into Data tab
    NDD_query = """
    SELECT					
        CONVERT(VARCHAR, AccountingMonth, 101) AS AccountingMonth,								
        CASE
            WHEN Make IN ('CHRYSLER', 'DODGE','JEEP','RAM') THEN 'CDJR'
            WHEN Make IN ('FORD', 'LINCOLN') THEN 'FORD'
            WHEN Make IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC') THEN 'GM'
            WHEN Make IN ('HYUNDAI', 'GENESIS') THEN 'HYUNDAI'
            WHEN Make IN ('JAGUAR', 'LAND ROVER') THEN 'JLR'
            ELSE Make
        END AS Brand,
        SUM(CASE WHEN Type = 'Inv' THEN QTY ELSE 0 END) AS Inv,					
        SUM(CASE WHEN Type = 'Sales' THEN QTY ELSE 0 END) AS Sales					
    FROM (
        SELECT					
            AccountingMonth,								
            Make,
            'Inv' AS Type,					
            SUM(InventoryCount) AS QTY					
        FROM NDDUsers.vInventoryMonthEnd					
        WHERE AccountingMonth = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 1, 0)			
            AND Department = '300'		
        GROUP BY AccountingMonth, Make
        HAVING SUM(InventoryCount) <> 0					
        UNION ALL					
        SELECT					
            AccountingMonth,							
            VehicleMakeName AS Make,
            'Sales' AS Type,					
            SUM(VehicleSoldCount) AS QTY					
        FROM NDDUsers.vSalesDetail_Vehicle					
        WHERE AccountingMonth = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 1, 0)			
            AND DepartmentName = 'NEW'					
            AND RecordSource = 'Accounting'
        GROUP BY AccountingMonth, VehicleMakeName
        HAVING SUM(VehicleSoldCount) <> 0					
    ) AS Subquery					
    WHERE Make IS NOT NULL
    GROUP BY AccountingMonth,
        CASE
            WHEN Make IN ('CHRYSLER', 'DODGE','JEEP','RAM') THEN 'CDJR'
            WHEN Make IN ('FORD', 'LINCOLN') THEN 'FORD'
            WHEN Make IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC') THEN 'GM'
            WHEN Make IN ('HYUNDAI', 'GENESIS') THEN 'HYUNDAI'
            WHEN Make IN ('JAGUAR', 'LAND ROVER') THEN 'JLR'
            ELSE Make
        END
    """

    try:
        with pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=nddprddb01,48155;'
                'DATABASE=NDD_ADP_RAW;'
                'Trusted_Connection=yes;'
        ) as conn:
            NDD_df = pd.read_sql(NDD_query, conn)

    except Exception as e:
        print("❌ Connection failed:", e)


    SSPR_query = """
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
    SUM(CommitM1) AS CommitM1,
    SUM(TurnedDown) AS TurnedDown,
    (CAST(SUM(CommitM1) AS INT) * 100 / NULLIF(CAST(SUM(EarnedM1) AS INT), 0)) AS [Approval%]

    FROM(

    SELECT --PIVOTQUERY
    Make,
    SUM(CASE WHEN AccountingMonth = 'Earned_M1' THEN QTY ELSE 0 END) AS EarnedM1,
    SUM(CASE WHEN AccountingMonth = 'Commit_M1' THEN QTY ELSE 0 END) AS CommitM1,
    SUM(CASE WHEN AccountingMonth = 'Earned_M1' THEN QTY ELSE 0 END) - SUM(CASE WHEN AccountingMonth = 'Commit_M1' THEN QTY ELSE 0 END) AS TurnedDown

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
    DC.[Month] in ( Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)) )

    UNION

    SELECT
    DF.[Year], DF.[Month], DF.[DealerVehicleID], DF.[Mth], DF.[ColumnCD], DF.[EnteredValue], DF.[StatusID]
    FROM
    [SSPRv3].[dbo].[DealerForecasts] AS DF
    WHERE
    DF.[Year] in (  Year(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)) ) 
    AND
    DF.[Month] in ( Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)) )
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
        DC.[Month] in ( Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)) )

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
        END
        """
    SSPR_History_query = """
    SELECT --ParentBrand
    
    Year,
    
    Month,
    
    --Hyperion,
    
    CASE
    
        WHEN Make IN ('CHRYSLER','DODGE','JEEP','RAM', 'FIAT', 'FORD','LINCOLN','BUICK','CADILLAC','CHEVROLET','GMC')
    
        THEN 'Domestic'
    
        WHEN Make IN ('ACURA','HONDA','GENESIS','HYUNDAI','INFINITI', 'MAZDA', 'NISSAN','SUBARU','TOYOTA','VOLKSWAGEN', 'VOLVO')
    
        THEN 'Import'
    
        WHEN Make IN ('Audi', 'BMW','JAGUAR','LAND ROVER','LEXUS','MERCEDES-BENZ','MINI')
    
        THEN 'Luxury'
    
        END AS Segment,
    
    CASE
    
        WHEN Make IN ('CHRYSLER', 'DODGE','JEEP','RAM', 'FIAT')
    
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
    
        END AS ParentBrand,
    
    Make,
    
    --Model,
    
    SUM(EarnedM1) AS EarnedM1,
    
    SUM(CommitM1) AS CommitM1


    FROM(

    SELECT --PIVOTQUERY
    
    Year,
    
    Month,
    
    Hyperion,
    
    Make,
    
    Model,
    
    SUM(CASE WHEN AccountingMonth = 'Earned_M1' THEN QTY ELSE 0 END) AS EarnedM1,
    
    SUM(CASE WHEN AccountingMonth = 'Commit_M1' THEN QTY ELSE 0 END) AS CommitM1,
    
    SUM(CASE WHEN AccountingMonth = 'Earned_M1' THEN QTY ELSE 0 END) - SUM(CASE WHEN AccountingMonth = 'Commit_M1' THEN QTY ELSE 0 END) AS TurnedDown,
    
    SUM(CASE WHEN AccountingMonth = 'Fcast_M1' THEN QTY ELSE 0 END) AS Fcast_M1,
    
    SUM(CASE WHEN AccountingMonth = 'Fcast_M2' THEN QTY ELSE 0 END) AS Fcast_M2,
    
    SUM(CASE WHEN AccountingMonth = 'Fcast_M3' THEN QTY ELSE 0 END) AS Fcast_M3,
    
    SUM(CASE WHEN AccountingMonth = 'Net_Add' THEN QTY ELSE 0 END) AS NetAdd


    FROM(

    SELECT
    
        DC.[Year] AS Year,
    
        DC.[Month]AS Month,
    
        D.[DealerCD] AS Hyperion,
    
        B.[BrandCD] AS Make,
    
        BM.[BrandModelCD] AS Model,
    
        --DC.[Mth],
    
            CASE
    
    WHEN DC.[ColumnCD] IN ('TCCM1', 'TCC2_M1','TCC3_M1')
    
                THEN 'Commit_M1'
    
                WHEN DC.[ColumnCD] IN ('EAM1', 'EA2_M1', 'EA3_M1')
    
                THEN 'Earned_M1'
    
                WHEN ColumnCD IN ('NAM1','NAM2','NAM3','NAM4','NAM5','NAM6')

                THEN 'Net_Add'
    
                WHEN DC.[ColumnCD] = 'SARFM1'
    
                THEN 'Fcast_M1'
    
                WHEN DC.[ColumnCD] = 'SARFM2'
    
                THEN 'Fcast_M2'
    
                WHEN DC.[ColumnCD] = 'SARFM3'
    
                THEN 'Fcast_M3'
    
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
    
    DC.[Year] >= YEAR(GETDATE()) - 2
    
    --AND
    
    --DC.[Month] in ( Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-5, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-4, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-3, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-2, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0))  )

    UNION

    SELECT
    
    DF.[Year], DF.[Month], DF.[DealerVehicleID], DF.[Mth], DF.[ColumnCD], DF.[EnteredValue], DF.[StatusID]
    
    FROM
    
    [SSPRv3].[dbo].[DealerForecasts] AS DF
    
    WHERE
    
    DF.[Year] >= YEAR(GETDATE()) - 2
    
    --AND
    
    --DF.[Month] in ( Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-5, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-4, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-3, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-2, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0))  )
    
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
    
    DC.[Year] >= YEAR(GETDATE()) - 2
    
    --AND
    
    --DC.[Month] in ( Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-5, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-4, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-3, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-2, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0))  )

    --AND BM.[BrandModelCD] = @BrandModelCD
    
    and DC.[ColumnCD] in ('EAM1' , 'TCCM1', 'EA2_M1' , 'TCC2_M1' , 'EA3_M1' , 'TCC3_M1', 'SARFM1', 'SARFM2', 'SARFM3', 'NAM1','NAM2','NAM3','NAM4','NAM5','NAM6')

    AND DC.[EnteredValue] <> '0'

    GROUP BY
    
        DC.[Year],
    
        DC.[Month],
    
    D.[DealerCD],
    
        B.[BrandCD],
    
        BM.[BrandModelCD],
    
        --DC.[Mth],
    
            CASE
    
                WHEN DC.[ColumnCD] IN ('TCCM1', 'TCC2_M1','TCC3_M1')
    
                THEN 'Commit_M1'
    
                WHEN DC.[ColumnCD] IN ('EAM1', 'EA2_M1', 'EA3_M1')
    
                THEN 'Earned_M1'
    
                WHEN ColumnCD IN ('NAM1','NAM2','NAM3','NAM4','NAM5','NAM6')

                THEN 'Net_Add'
    
                WHEN DC.[ColumnCD] = 'SARFM1'
    
                THEN 'Fcast_M1'
    
                WHEN DC.[ColumnCD] = 'SARFM2'
    
                THEN 'Fcast_M2'
    
                WHEN DC.[ColumnCD] = 'SARFM3'
    
                THEN 'Fcast_M3'
    
                ELSE 'ERROR'
    
                END

    )AS PIVOTQUERY

    GROUP BY
    
    Year,
    
    Month,
    
    Make,
    
    Model,
    
    Hyperion

    )AS ParentBrand

    GROUP BY
    
    Year,
    
    Month,
    
    Make,
    
    --Model,
    
    --Hyperion,
    
    CASE
    
        WHEN Make IN ('CHRYSLER', 'DODGE','JEEP','RAM', 'FIAT')
    
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
    
        WHEN Make IN ('CHRYSLER','DODGE','JEEP','RAM', 'FIAT', 'FORD','LINCOLN','BUICK','CADILLAC','CHEVROLET','GMC')
    
        THEN 'Domestic'
    
        WHEN Make IN ('ACURA','HONDA', 'GENESIS','HYUNDAI','INFINITI', 'MAZDA', 'NISSAN','SUBARU','TOYOTA','VOLKSWAGEN', 'VOLVO')
    
        THEN 'Import'
    
        WHEN Make IN ('Audi', 'BMW','JAGUAR','LAND ROVER','LEXUS','MERCEDES-BENZ','MINI')
    
        THEN 'Luxury'
    
        END
"""    

    # right click on connection and go to properties to find the server name then select the database its looking at
    try:
        with pyodbc.connect(
                r'DRIVER={ODBC Driver 17 for SQL Server};'
                r'SERVER=BAPRDDB01\BAPRD,49174;'
                r'DATABASE=SSPRv3;'
                r'Trusted_Connection=yes;'
        ) as conn:
            SSPR_df = pd.read_sql(SSPR_query, conn)
            SSPR_history_df = pd.read_sql(SSPR_History_query, conn)

    except Exception as e:
        print("❌ Connection failed:", e)

################################################################################################################################
    '''DROP IN SQL QUERIES INTO EXCEL FILE AND RUN MACRO FOR UPDATE'''
################################################################################################################################

    # Open file and process macro/Sql
    Allocation_Tracker_File = r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker\Allocation_Tracker.xlsm"

    Process_Daily_Sales_File()

    # Use Dynamic_Daily_Sales.xlsb
    Dynamic_Daily_Sales_File = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker\Dynamic_Daily_Sales.xlsb'


    app = xw.App(visible=True) 
    Allocation_wb = app.books.open(Allocation_Tracker_File)
    # We need Daily sales open so that Excel VBA can process it
    Daily_sales_wb = app.books.open(Dynamic_Daily_Sales_File)

    # Dump data into tab section to process'
    SSPR_tab = Allocation_wb.sheets['Store SSPR DATA']
    SSPR_tab.range('N2:S100').clear_contents()
    SSPR_tab.range('N2').options(index=False).value = SSPR_df

    NDD_tab = Allocation_wb.sheets['NDD DATA']
    NDD_tab.clear_contents()
    NDD_tab.range('A1').options(index=False).value = NDD_df


    # APPEND/REPLACE CURRENT MONTH QUERY DATA: LOGIC WILL SEARCH FOR CURRENT MONTH IN THE MTH COLUMN AND REPLACE OTHERWISE APPEND

    # Get current date
    # today = datetime.today()
    # if today.day >= 7:
    #     # Get prior month/year as a string (e.g., "202301" for January 2023)
    #     formatted_date = today.strftime("%Y%m")
    #     print(formatted_date)
    #     # Use the 1st of the current month
    #     result_date = today.replace(day=1)
    #     result_date = f"{result_date.month}/{result_date.day}/{result_date.year}"
    # else:
    #     # Get prior month/year as a string (e.g., "202305" for June 2023)
    #     formatted_date = today.replace(day=1)
    #     formatted_date = formatted_date - timedelta(days=1)
    #     formatted_date = formatted_date.strftime("%Y%m")
    #     print(formatted_date)
    #     # Use the 1st of the previous month
    #     first_of_current = today.replace(day=1)
    #     result_date = first_of_current - timedelta(days=1)
    #     result_date = result_date.replace(day=1)
    #     result_date = f"{result_date.month}/{result_date.day}/{result_date.year}"

    # print(result_date)

    # adjust logic: we want to pull prior month at all times!

    today = datetime.today()
    formatted_date = today.replace(day=1)
    formatted_date = formatted_date - timedelta(days=1)
    formatted_date = formatted_date.strftime("%Y%m")
    print(formatted_date)
    # Use the 1st of the previous month
    first_of_current = today.replace(day=1)
    result_date = first_of_current - timedelta(days=1)
    result_date = result_date.replace(day=1)
    result_date = f"{result_date.month}/{result_date.day}/{result_date.year}"

    print(result_date)
 

    SSPR_Historic_Data_tab = Allocation_wb.sheets['SSPR Historic Data']

    # Read existing data (assuming headers in row 1, data from B2)
    data_range = SSPR_Historic_Data_tab.range('A1').expand('down').resize(None, 9)
    existing_df = data_range.options(pd.DataFrame, header=1).value

    # Add headers if needed (assuming DataFrame has same column structure)
    existing_df.reset_index(inplace=True)

    # Filter out rows for the current month in Year/Month column
    filtered_df = existing_df[existing_df['Year/Month'].astype(int) != int(formatted_date)]

    # Ensure Month and Year are strings
    SSPR_history_df['Month'] = SSPR_history_df['Month'].astype(int)  # Ensure numeric first
    SSPR_history_df['Year'] = SSPR_history_df['Year'].astype(int)

    # Pad the month with leading zero if needed
    SSPR_history_df['Month_Format'] = SSPR_history_df['Month'].astype(int).astype(str).str.zfill(2)

    # Add in the Year/Month and Accounting Month data to SSPR_history_df    
    SSPR_history_df['Year/Month'] = SSPR_history_df['Year'].astype(str) + SSPR_history_df['Month_Format']
    # Build the Accounting Month as a real datetime
    SSPR_history_df['Accounting Month'] = pd.to_datetime(
        SSPR_history_df['Year'].astype(str) + '-' +
        SSPR_history_df['Month'].astype(int).astype(str).str.zfill(2) + '-01'
    )

    # Filter the Current Month/Year (last month's) to append on to existing dataframe
    SSPR_history_df = SSPR_history_df[SSPR_history_df['Year/Month'].astype(int) == int(formatted_date)]

    SSPR_history_df = SSPR_history_df.drop(columns=['Month_Format'])

    # Append last month's SSPR_history to existing data
    updated_df = pd.concat([filtered_df, SSPR_history_df], ignore_index=True)

    # Clear full range before re-writing cleaned data
    SSPR_Historic_Data_tab.range('A5:J500000').clear_contents()

    # Insert historic data into tab
    SSPR_Historic_Data_tab.range('A5').options(index=False, header=False).value = updated_df


    # Run Macro
    Run_Macro = Allocation_wb.macro("ExecuteMacros")
    Run_Macro()
    Allocation_wb.save()

    # Save and close the excel document(s)    
    if Allocation_wb:
        Allocation_wb.close()
    if app:
        app.quit()
    
    # Send file to Shared folder and trigger Power Automate
    # shutil.copyfile(Allocation_Tracker_File, r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Triggers\Allocation_File_Link\Allocation_Tracker.xlsm')

    # End Timer
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Script completed in {elapsed_time:.2f} seconds")

    return

#run function
if __name__ == '__main__':
    #Process_Daily_Sales_File()
    Allocation_Tracker_Update()