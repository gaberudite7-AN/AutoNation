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


def EV_Availability_Update():

    # Get current date
    today = datetime.today()

    # Begin timer
    start_time = time.time()

    # Current date
    today = datetime.today()

    # Equivalent to SQL: DECLARE @StartDate DATE = '2025-06-01' (We want to take the past 2 months);
    # Go to the last day of the previous month
    last_day_previous_month = today.replace(day=1) - timedelta(days=1)

    # Get the first day of the previous month
    first_day_previous_month = last_day_previous_month.replace(day=1)

    # Optional: format as string
    Start_Date = first_day_previous_month.strftime('%Y-%m-%d')

    print(f"Starting date is {Start_Date}")

    # Run SQL queries using SQL Alchemy and dump into Data tab
    Sales_Inv_Pipeline_query = f"""
DECLARE @GivenDate DATE = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0);
DECLARE @CurrentMonth DATE = @GivenDate;
DECLARE @StartDate DATE = '{Start_Date}';
DECLARE @Yesterday DATE = DATEADD(DAY, -1, GETDATE());
 
-- Aggregate sales by month and make
WITH SalesData AS (
    SELECT
        DATEADD(MONTH, DATEDIFF(MONTH, 0, AccountingMonth), 0) AS AccountingMonth,
        StoreName,
        StoreHyperion AS Hyperion,
        AllocationGroup AS Model,
        VehicleFuelType AS FuelType,
        -- Use more efficient string formatting if available in your SQL version
        CONCAT(UPPER(LEFT(VehicleMakeName, 1)), LOWER(SUBSTRING(VehicleMakeName, 2, LEN(VehicleMakeName)))) AS VehicleMakeName,
        SUM(VehicleSoldCount) AS SalesCount
    FROM NDD_ADP_RAW.NDDUsers.vSalesDetail_Vehicle
    WHERE AccountingMonth >= @StartDate
        AND DepartmentName = 'new'
        AND MarketName NOT IN ('market 97', 'market 98')
        AND VehicleMakeName IS NOT NULL
    GROUP BY 
        DATEADD(MONTH, DATEDIFF(MONTH, 0, AccountingMonth), 0),
        StoreName,
        StoreHyperion,
        AllocationGroup,
        VehicleFuelType,
        CONCAT(UPPER(LEFT(VehicleMakeName, 1)), LOWER(SUBSTRING(VehicleMakeName, 2, LEN(VehicleMakeName))))
),
InventoryBase AS (
    -- Base inventory query
    SELECT
        AccountingMonth,
        StoreName,
        AllocationGroup AS Model,
        Hyperion,
        FuelType,
        CONCAT(UPPER(LEFT(Make, 1)), LOWER(SUBSTRING(Make, 2, LEN(Make)))) AS VehicleMakeName,
        SUM(InventoryCount) AS InventoryCount,
        0 AS Not_Produced_Count,
        0 AS To_Be_Built_Count,
        0 AS Built_Count,
        0 AS InTransitCount
    FROM NDD_ADP_RAW.NDDUsers.vInventoryMonthEnd
    WHERE AccountingMonth >= @StartDate
        AND RegionName <> 'AND Corporate Management'
        AND MarketName NOT IN ('market 97', 'market 98')
        AND Department = '300'
        AND Status IS NOT NULL
        AND Make IS NOT NULL
        AND NOT (AccountingMonth = @CurrentMonth AND Status = 'G')
    GROUP BY 
        AccountingMonth,
        StoreName,
        AllocationGroup,
        Hyperion,
        FuelType,
        CONCAT(UPPER(LEFT(Make, 1)), LOWER(SUBSTRING(Make, 2, LEN(Make))))
),
OnOrderBase AS (
    -- Consolidated OnOrder queries
    SELECT
        AccountingMonth,
        StatusName AS StoreName,
        alloc_grp_map AS Model,
        hyperion_id AS Hyperion,
        StatusName AS FuelType,
        CONCAT(UPPER(LEFT(make, 1)), LOWER(SUBSTRING(make, 2, LEN(make)))) AS VehicleMakeName,
        0 AS InventoryCount,
        CASE WHEN an_status = 'NOT PRODUCED' THEN COUNT(*) ELSE 0 END AS Not_Produced_Count,
        CASE WHEN an_status = 'TO BE BUILT' THEN COUNT(*) ELSE 0 END AS To_Be_Built_Count,
        CASE WHEN an_status = 'BUILT' THEN COUNT(*) ELSE 0 END AS Built_Count,
        CASE WHEN an_status = 'in transit' THEN COUNT(*) ELSE 0 END AS InTransitCount
    FROM (
        SELECT 
            DATEADD(MONTH, DATEDIFF(MONTH, 0, CAST(CONCAT(LEFT(time_period, 4), '-', RIGHT(time_period, 2), '-01') AS DATE)), 0) AS AccountingMonth,
            an_status,
            CASE 
                WHEN an_status = 'NOT PRODUCED' THEN 'Not_Produced'
                WHEN an_status = 'TO BE BUILT' THEN 'To_Be_Built'
                WHEN an_status = 'BUILT' THEN 'Built'
                WHEN an_status = 'in transit' THEN 'In Transit'
            END AS StatusName,
            alloc_grp_map,
            hyperion_id,
            make
        FROM NDD_ADP_RAW.NDDUsers.vOnOrderInfo_SnapShot
        WHERE DATEADD(MONTH, DATEDIFF(MONTH, 0, CAST(CONCAT(LEFT(time_period, 4), '-', RIGHT(time_period, 2), '-01') AS DATE)), 0) >= @StartDate
            AND make IS NOT NULL
            AND an_status IN ('NOT PRODUCED', 'TO BE BUILT', 'BUILT', 'in transit')
        UNION ALL
        SELECT 
            @CurrentMonth AS AccountingMonth,
            an_status,
            CASE 
                WHEN an_status = 'NOT PRODUCED' THEN 'CM_Not_Produced'
                WHEN an_status = 'TO BE BUILT' THEN 'CM_To_Be_Built'
                WHEN an_status = 'BUILT' THEN 'CM_Built'
                WHEN an_status = 'in transit' THEN 'CM In Transit'
            END AS StatusName,
            alloc_grp_map,
            hyperion_id,
            make
        FROM NDD_ADP_RAW.NDDUsers.vOnOrderInfo_Dly_SnapShot
        WHERE update_date = @Yesterday
            AND make IS NOT NULL
            AND an_status IN ('NOT PRODUCED', 'TO BE BUILT', 'BUILT', 'in transit')
    ) AS OnOrder
    GROUP BY 
        AccountingMonth,
        StatusName,
        alloc_grp_map,
        hyperion_id,
        StatusName,
        an_status,
        CONCAT(UPPER(LEFT(make, 1)), LOWER(SUBSTRING(make, 2, LEN(make))))
),
Combined AS (
    SELECT 
        AccountingMonth,
        StoreName,
        Model,
        Hyperion,
        FuelType,
        VehicleMakeName,
        SUM(InventoryCount) AS InventoryCount,
        SUM(Not_Produced_Count) AS Not_Produced_Count,
        SUM(To_Be_Built_Count) AS To_Be_Built_Count,
        SUM(Built_Count) AS Built_Count,
        SUM(InTransitCount) AS InTransitCount
    FROM (
        SELECT * FROM InventoryBase
        UNION ALL
        SELECT * FROM OnOrderBase
    ) AS AllData
    GROUP BY AccountingMonth, StoreName, Model, Hyperion, FuelType, VehicleMakeName
)
SELECT
    c.AccountingMonth,
    YEAR(c.AccountingMonth) AS AccountingYear,
    c.StoreName,
    c.Model,
    c.Hyperion,
    c.FuelType,
    c.VehicleMakeName,
    CASE
        WHEN c.VehicleMakeName IN ('Chrysler','Dodge','Jeep','Ram','Fiat') THEN 'CDJR'
        WHEN c.VehicleMakeName IN ('Buick','GMC','Cadillac','Chevrolet') THEN 'GM'
        WHEN c.VehicleMakeName IN ('Ford','Lincoln') THEN 'Ford LM'
        WHEN c.VehicleMakeName IN ('Jaguar','Land Rover') THEN 'Jaguar Land Rover'
        WHEN c.VehicleMakeName IN ('Hyundai','Genesis') THEN 'Hyundai'
        ELSE c.VehicleMakeName
    END AS [Brand Group],
    CASE
        WHEN c.VehicleMakeName IN ('Chrysler','Dodge','Jeep','Ram','Fiat','Ford','Lincoln','Cadillac','Buick','GMC','Chevrolet') THEN 'Domestic'
        WHEN c.VehicleMakeName IN ('Toyota','Honda','Acura','Nissan','Infiniti','Subaru','Volkswagen','Volvo','Mazda','Hyundai','Genesis') THEN 'Import'
        WHEN c.VehicleMakeName IN ('Mercedes-Benz','Audi','BMW','Mini','Lexus','Jaguar','Land Rover','Porsche','Aston Martin','Bentley') THEN 'Premium Luxury'
        ELSE 'Other'
    END AS Segment,
    SUM(c.InventoryCount) AS InventoryCount,
    SUM(c.Not_Produced_Count) AS Not_Produced_Count,
    SUM(c.To_Be_Built_Count) AS To_Be_Built_Count,
    SUM(c.Built_Count) AS Built_Count,
    SUM(c.InTransitCount) AS InTransitCount,
    COALESCE(s.SalesCount, 0) AS SalesCount
FROM Combined c
LEFT JOIN SalesData s
    ON c.AccountingMonth = s.AccountingMonth
    AND c.Model = s.Model
    AND c.Hyperion = s.Hyperion
    AND c.FuelType = s.FuelType
    AND c.VehicleMakeName = s.VehicleMakeName
GROUP BY 
    c.AccountingMonth,
    c.StoreName,
    c.Model,
    c.Hyperion,
    c.FuelType,
    c.VehicleMakeName,
    s.SalesCount
ORDER BY 
    c.AccountingMonth DESC,
    c.VehicleMakeName;
    """

    Hyperion_Lookup_query = """
    SELECT [HYPERION_ID] as Hyperion
        ,[ENTITY_NAME] as StoreName
    FROM [NDD_ADP_RAW].[NDDUsers].[vHyperionDetail]
    """

    Fuel_Type_Lookup_query = """
SELECT
AllocationGroup as Model,
CASE
	WHEN FuelType = 'Electric Fuel System'
	THEN 'EV'
	ELSE 'ICE'
END AS EV_Flag



FROM
NDDUsers.vInventoryMonthEnd

WHERE
AccountingMonth > = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-6, 0)
and Department = 300

GROUP BY
AllocationGroup,
CASE
	WHEN FuelType = 'Electric Fuel System'
	THEN 'EV'
	ELSE 'ICE'
END
"""

    SalesMix_query = """
WITH DenominatorTable AS (
    SELECT
        StoreHyperion,
        SUM(VehicleSoldCount) AS SoldCount
    FROM
        NDDUsers.vSalesDetail
    WHERE
        AccountingMonth = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0)
        AND DepartmentName = 'NEW'
        AND RecordSource = 'Accounting'
        AND VehicleMakeName IS NOT NULL
    GROUP BY
        StoreHyperion
    HAVING
        SUM(VehicleSoldCount) <> 0
)

SELECT
    A.StoreHyperion,
    A.VehicleMakeName,
    A.AllocationGroup,
    SUM(A.VehicleSoldCount) AS SoldCount,
    B.SoldCount AS Denominator,
    SUM(A.VehicleSoldCount) * 1.0 / NULLIF(B.SoldCount, 0) AS SalesMix
FROM
    NDDUsers.vSalesDetail A
LEFT JOIN
    DenominatorTable B ON B.StoreHyperion = A.StoreHyperion
WHERE
    AccountingMonth = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0)
    AND A.DepartmentName = 'NEW'
    AND A.RecordSource = 'Accounting'
    AND A.VehicleMakeName IS NOT NULL
GROUP BY
    A.StoreHyperion,
    A.VehicleMakeName,
    B.SoldCount,
    A.AllocationGroup
HAVING
    SUM(A.VehicleSoldCount) <> 0    
"""    
    
    try:
        with pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=nddprddb01,48155;'
                'DATABASE=NDD_ADP_RAW;'
                'Trusted_Connection=yes;'
        ) as conn:
            print("Reading in SQL Queries...")
            Sales_Inv_Pipeline_df = pd.read_sql(Sales_Inv_Pipeline_query, conn)
            Hyperion_Lookup_df = pd.read_sql(Hyperion_Lookup_query, conn)
            SalesMix_df = pd.read_sql(SalesMix_query, conn)
        end_time = time.time()
        elapsed_time = end_time - start_time
        # Print time it took to load query
        print(f"Script read in Queries and converted to dataframes in {elapsed_time:.2f} seconds")
    except Exception as e:
        print("❌ Connection failed:", e)

################################################################################################################################
    '''DROP IN SQL QUERIES INTO EXCEL FILE AND RUN MACRO FOR UPDATE'''
################################################################################################################################

    # Read In Fuel_List file:
    Fuel_Lookup_df = pd.read_csv(r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Inventory_Availability\Fuel_List.csv')
    print(Fuel_Lookup_df)

    # Lookup 1: Hyperion into Lookup tab to complete StoreNames on In Transit/Pipeline vehicles
    # Convert Key to string for lookup processing
    Sales_Inv_Pipeline_df['Hyperion'] = Sales_Inv_Pipeline_df['Hyperion'].astype(str)
    Hyperion_Lookup_df['Hyperion'] = Hyperion_Lookup_df['Hyperion'].astype(str)
    
    # Bring in StoreName
    # Print row counts before merge
    print("Sales_Inv_Pipeline_df rows prior to merge:", len(Sales_Inv_Pipeline_df))
    Sales_Inv_Pipeline_df = Sales_Inv_Pipeline_df.merge(Hyperion_Lookup_df, on='Hyperion', how='left')
    print("Sales_Inv_Pipeline_df rows after merge:", len(Sales_Inv_Pipeline_df))
    Sales_Inv_Pipeline_df.drop(columns=['StoreName_x'], inplace=True)
    Sales_Inv_Pipeline_df.rename(columns={'StoreName_y': 'StoreName'}, inplace=True) 
    
    # Lookup 2: AllocationGroup/Model into Lookup tab to bring in EV Vehicles
    # Convert Key to string for lookup processing
    Sales_Inv_Pipeline_df['Model'] = Sales_Inv_Pipeline_df['Model'].astype(str)
    Fuel_Lookup_df['Model'] = Fuel_Lookup_df['Model'].astype(str)
    print("Sales_Inv_Pipeline_df rows prior to merge:", len(Sales_Inv_Pipeline_df))    
    Sales_Inv_Pipeline_df = Sales_Inv_Pipeline_df.merge(Fuel_Lookup_df, on='Model', how='left')
    print("Sales_Inv_Pipeline_df rows after merge:", len(Sales_Inv_Pipeline_df))

    # EV Flag 2 Formula to process if Fuel Type meets categories and EV query is EV then EV
    Sales_Inv_Pipeline_df['EV_Flag2'] = np.where(
        (Sales_Inv_Pipeline_df['EV_Flag'] == 'EV') &
        (Sales_Inv_Pipeline_df['FuelType'].isin([
            'Electric Fuel System',
            'Gas/Electric Hybrid',
            'Gasoline/Mild Electric Hy',
            'Plug-In Electric/Gas'
        ])),
        'EV',  # value if condition is True
        'Non-EV'  # value if condition is False
    )

    # Create formula to distinguish current month/prior month
    Sales_Inv_Pipeline_df['Month'] = np.where(Sales_Inv_Pipeline_df['AccountingMonth'] == Start_Date, "Prior Month", "Current Month")

    # filter down to only EVs
    Sales_Inv_Pipeline_df = Sales_Inv_Pipeline_df[Sales_Inv_Pipeline_df['EV_Flag'] == 'EV'].copy()
    # filter down to only EVs
    Sales_Inv_Pipeline_df = Sales_Inv_Pipeline_df[Sales_Inv_Pipeline_df['EV_Flag'] == 'EV'].copy()
    # Split up into 2 dataframes (Current and Prior Month)
    Sales_Inv_Pipeline_df_Current = Sales_Inv_Pipeline_df[Sales_Inv_Pipeline_df['Month'] == "Current Month"].copy()
    Sales_Inv_Pipeline_df_Prior = Sales_Inv_Pipeline_df[Sales_Inv_Pipeline_df['Month'] == "Prior Month"].copy()

    """Current Month"""
    # Group to remove duplicates for fuel type
    Sales_Inv_Pipeline_df_Current = Sales_Inv_Pipeline_df_Current.groupby(
        ['StoreName', 'Brand Group', 'Hyperion', 'Model', 'VehicleMakeName'],
        as_index=False
    ).agg({
        'InventoryCount': 'sum',
        'Not_Produced_Count': 'sum',
        'To_Be_Built_Count': 'sum',
        'Built_Count': 'sum',
        'InTransitCount': 'sum',
        'SalesCount': 'sum'
    })

    Sales_Inv_Pipeline_df_Current['Month'] = 'Current Month'

    # Bring in Pace from Success File
    Success_File = r'W:\Applications\PowerBI\Pricing and Inventory\Targets and Pace\Success Menu Dashboard Data Template_Sales Inventory.xlsx'  # Change to your actual filename    
    Success_df = pd.read_excel(Success_File, sheet_name='Unit Sales')

    # Convert Key to string for lookup processing
    # First dropnas
    Sales_Inv_Pipeline_df_Current = Sales_Inv_Pipeline_df_Current.dropna(subset=['Hyperion'])
    Success_df = Success_df.dropna(subset=['Hyperion'])
    # Convert to string
    Sales_Inv_Pipeline_df_Current['Hyperion'] = Sales_Inv_Pipeline_df_Current['Hyperion'].astype(float).astype(int).astype(str)
    Success_df['Hyperion'] = Success_df['Hyperion'].astype(float).astype(int).astype(str)

    # Limit Success df to only New data
    Success_df = Success_df[Success_df['Dept'] == 'New'].copy()
    # Merge Pace into Dataframe
    print("Sales_Inv_Pipeline_df rows prior to merge:", len(Sales_Inv_Pipeline_df_Current))
    Sales_Inv_Pipeline_df_Current = Sales_Inv_Pipeline_df_Current.merge(Success_df[['Hyperion', 'Unit Sales Month Pace']], on='Hyperion', how='left')
    print("Sales_Inv_Pipeline_df rows after merge:", len(Sales_Inv_Pipeline_df_Current))
    Sales_Inv_Pipeline_df_Current = Sales_Inv_Pipeline_df_Current.rename(columns={'Unit Sales Month Pace': 'Pace'})

    # Normalize values, fill nas with 0
    Sales_Inv_Pipeline_df_Current['Pace'] = pd.to_numeric(Sales_Inv_Pipeline_df_Current['Pace'], errors='coerce').fillna(0).astype(float)

    # Bring in prior month Sales Mix
    # Normalize Keys
    Sales_Inv_Pipeline_df_Current['SalesMixKey'] = Sales_Inv_Pipeline_df_Current['Hyperion'].astype(str).str.strip() + Sales_Inv_Pipeline_df_Current['Model'].astype(str).str.strip()
    SalesMix_df['SalesMixKey'] = SalesMix_df['StoreHyperion'].astype(str).str.strip() + SalesMix_df['AllocationGroup'].astype(str).str.strip()
    # Merge SalesMix into Dataframe
    print("Sales_Inv_Pipeline_df rows prior to merge:", len(Sales_Inv_Pipeline_df_Current))
    Sales_Inv_Pipeline_df_Current = Sales_Inv_Pipeline_df_Current.merge(SalesMix_df[['SalesMixKey', 'SalesMix']], on='SalesMixKey', how='left')
    Sales_Inv_Pipeline_df_Current['SalesMix'] = pd.to_numeric(Sales_Inv_Pipeline_df_Current['SalesMix'], errors='coerce').fillna(0).astype(float)
    print("Sales_Inv_Pipeline_df rows after to merge:", len(Sales_Inv_Pipeline_df_Current))

    # Multiply Pace by Sales Mix
    Sales_Inv_Pipeline_df_Current['Pace_By_Model (Based on Prior %)'] = Sales_Inv_Pipeline_df_Current['Pace']*Sales_Inv_Pipeline_df_Current['SalesMix']
    
    """Prior Month fill empty cells for appending"""
    # Group to remove duplicates for fuel type
    Sales_Inv_Pipeline_df_Prior = Sales_Inv_Pipeline_df_Prior.groupby(
        ['StoreName', 'Brand Group', 'Hyperion', 'Model', 'VehicleMakeName'],
        as_index=False
    ).agg({
        'InventoryCount': 'sum',
        'Not_Produced_Count': 'sum',
        'To_Be_Built_Count': 'sum',
        'Built_Count': 'sum',
        'InTransitCount': 'sum',
        'SalesCount': 'sum'
    })

    Sales_Inv_Pipeline_df_Prior['Month'] = 'Prior Month'
    Sales_Inv_Pipeline_df_Prior['Pace'] = 0
    Sales_Inv_Pipeline_df_Prior['SalesMixKey'] = 0
    Sales_Inv_Pipeline_df_Prior['SalesMix'] = 0
    Sales_Inv_Pipeline_df_Prior['Pace_By_Model (Based on Prior %)'] = 0


    Sales_Inv_Pipeline_df = pd.concat([Sales_Inv_Pipeline_df_Current, Sales_Inv_Pipeline_df_Prior])

    EV_File = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Inventory_Availability\EV.xlsm'
    app = xw.App(visible=True) 
    EV_wb = app.books.open(EV_File)

    EV_tab = EV_wb.sheets['Data']
    EV_tab.range("A1:P50000").clear_contents()
    EV_tab.range('A1').options(index=False).value = Sales_Inv_Pipeline_df

    # End Timer
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Script completed in {elapsed_time:.2f} seconds")

    return

#run function
if __name__ == '__main__':
    
    EV_Availability_Update()