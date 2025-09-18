# %%
import pandas as pd
import pyodbc
import time
import warnings
warnings.filterwarnings("ignore", category=UserWarning, message=".*pandas only supports SQLAlchemy.*")



def Clean_csv():

#     Query = """
# --Sales/Inv Dynamic Dates
# DECLARE @SnapshotDate DATE = GETDATE();
# DECLARE @StartDate DATE = DATEADD(MONTH, -37, DATEFROMPARTS(YEAR(@SnapshotDate), MONTH(@SnapshotDate), 1));
# DECLARE @EndDate DATE = DATEFROMPARTS(YEAR(@SnapshotDate), MONTH(@SnapshotDate), 1);
# --Pipeline Dynamic Dates
# DECLARE @GivenDate DATE = DATEFROMPARTS(YEAR(GETDATE()), MONTH(GETDATE()), 1);
# DECLARE @QuarterStart DATE = DATEFROMPARTS(YEAR(@GivenDate), ((DATEPART(QUARTER, @GivenDate) - 1) * 3) + 1, 1);
# DECLARE @CurrentMonth DATE = @GivenDate;
# DECLARE @Yesterday DATETIME = CAST(CONVERT(date, DATEADD(DAY, -1, GETDATE())) AS DATETIME);

# WITH InventoryCTE AS (
#     SELECT 
#         AccountingMonth       = AccountingMonth,
#         MarketName            = MarketName,
#         StoreName             = StoreName,
#         Department            = Department,
#         StyleName             = StyleName,
#         FuelType              = FuelType,
#         Make                  = Make,
#         Model                 = Model,
#         Trim                  = Trim,
#         Mileage               = Mileage,
#         Hyperion              = Hyperion,
#         DaysInInventoryAN     = DaysInInventoryAN,
#         DaysInInventoryStore  = DaysInInventoryStore,
#         Loaner                = Loaner,
#         MSRP_Adv              = MSRP_Adv,
#         AllocationGroup       = AllocationGroup,
#         Color                 = ExteriorColor,
#         Status                = [Status],
#         Balance               = Balance,
#         ExService             = CAST('' AS VARCHAR(50)),
#         Table_Source          = 'Inventory_Table',
#         InvCount              = SUM(inventorycount),
#         SumSold               = CAST(0 AS MONEY),
#         SumFrontGross         = CAST(0 AS MONEY),
#         SumFIGross            = CAST(0 AS MONEY),
#         SumIncentiveGross     = CAST(0 AS MONEY),
#         SumOVIGross           = CAST(0 AS MONEY),
#         InTransitCount        = CAST(0 AS INT),
#         PipelineCount         = CAST(0 AS INT)
#     FROM NDD_ADP_RAW.NDDUsers.vInventoryMonthEnd
#     WHERE 
#         AccountingMonth BETWEEN @StartDate AND @EndDate
#         AND MarketName NOT IN ('market 98', 'market 97')
# 		AND Department = '300'
#     GROUP BY 
#         AccountingMonth, MarketName, StoreName, Department, StyleName, FuelType,
#         Make, Model, Trim, Mileage, Hyperion, DaysInInventoryAN, DaysInInventoryStore,
#         Loaner, MSRP_Adv, AllocationGroup, ExteriorColor, [Status], Balance
#     HAVING SUM(inventorycount) > 0
# ), 
# SalesCTE AS (
#     SELECT 
#         AccountingMonth       = AccountingMonth,
#         MarketName            = MarketName,
#         StoreName             = StoreName,
#         Department            = DepartmentName,
#         StyleName             = VehicleStyleName,
#         FuelType              = VehicleFuelType,
#         Make                  = VehicleMakeName,
#         Model                 = VehicleModelName,
#         Trim                  = VehicleTrimName,
#         Mileage               = VehicleMileage,
#         Hyperion              = StoreHyperion, 
#         DaysInInventoryAN     = CAST(0 AS INT),
#         DaysInInventoryStore  = CAST(0 AS INT),
#         Loaner                = CAST('' AS VARCHAR(50)),
#         MSRP_Adv              = CAST(0 AS MONEY),
#         AllocationGroup       = CAST('' AS VARCHAR(50)),
#         Color                 = CAST('' AS VARCHAR(50)),
#         Status                = CAST('' AS VARCHAR(50)),
#         Balance               = CAST(0 AS MONEY),
#         ExService             = ExService,
#         Table_Source          = 'Sales_Table',
#         InvCount              = CAST(0 AS INT),
#         SumSold               = SUM(VehicleSoldCount),
#         SumFrontGross         = SUM(FrontGross),
#         SumFIGross            = SUM(FIGross),
#         SumIncentiveGross     = SUM(Incentives),
#         SumOVIGross           = Sum(OVI),
#         InTransitCount        = CAST(0 AS INT),
#         PipelineCount         = CAST(0 AS INT)
#     FROM NDD_ADP_RAW.NDDUsers.vSalesDetail_Vehicle
#     WHERE 
#         AccountingMonth BETWEEN @StartDate AND @EndDate
#         AND MarketName NOT IN ('market 98', 'market 97')
#         AND VehicleModelYear > 2020
# 		AND DepartmentName = 'New'
#     GROUP BY 
#         AccountingMonth, MarketName, StoreName, DepartmentName, VehicleStyleName,
#         VehicleFuelType, VehicleMakeName, VehicleModelName, VehicleTrimName,
#         VehicleMileage, StoreHyperion,  ExService
# ),
# PipelineCTE AS (
#     -- OnOrder Snapshot: In Transit
#     SELECT 
#         DATEFROMPARTS(LEFT(CAST(time_period AS VARCHAR), 4), RIGHT(CAST(time_period AS VARCHAR), 2), 1) AS AccountingMonth,
#         MarketName            = 'In Transit',
#         StoreName             = 'In Transit',
#         Department            = 'In Transit',
#         StyleName             = 'In Transit',
#         FuelType              = 'In Transit',
#         UPPER(LEFT(LOWER(make), 1)) + LOWER(SUBSTRING(make, 2, LEN(make))) AS Make,
#         Model                 = 'In Transit',
#         Trim                  = 'In Transit',
#         Mileage               = CAST('' AS VARCHAR(50)),
#         Hyperion              = hyperion_id, 
#         DaysInInventoryAN     = CAST(0 AS INT),
#         DaysInInventoryStore  = CAST(0 AS INT),
#         Loaner                = CAST('' AS VARCHAR(50)),
#         MSRP_Adv              = CAST(0 AS MONEY),
#         AllocationGroup       = CAST('' AS VARCHAR(50)),
#         Color                 = CAST('' AS VARCHAR(50)),
#         Status                = CAST('' AS VARCHAR(50)),
#         Balance               = CAST(0 AS MONEY),
#         ExService             = 'In Transit',
#         Table_Source          = 'In Transit',
#         InvCount              = CAST(0 AS INT),
#         SumSold               = CAST(0 AS MONEY),
#         SumFrontGross         = CAST(0 AS MONEY),
#         SumFIGross            = CAST(0 AS MONEY),
#         SumIncentiveGross     = CAST(0 AS MONEY),
#         SumOVIGross           = CAST(0 AS MONEY),
#         COUNT(*) AS InTransitCount,
#         PipelineCount         = CAST(0 AS INT)
#     FROM NDD_ADP_RAW.NDDUsers.vOnOrderInfo_SnapShot
#     WHERE 
#         an_status = 'in transit'
#         AND DATEFROMPARTS(LEFT(CAST(time_period AS VARCHAR), 4), RIGHT(CAST(time_period AS VARCHAR), 2), 1) >= @StartDate
#         AND (
#             DATEFROMPARTS(LEFT(CAST(time_period AS VARCHAR), 4), RIGHT(CAST(time_period AS VARCHAR), 2), 1) BETWEEN @QuarterStart AND DATEADD(MONTH, -1, @CurrentMonth)
#             OR (
#                 RIGHT(CAST(time_period AS VARCHAR), 2) IN ('03','06','09','12')
#                 AND DATEFROMPARTS(LEFT(CAST(time_period AS VARCHAR), 4), RIGHT(CAST(time_period AS VARCHAR), 2), 1) < @QuarterStart
#             )
#         )
#         AND make IS NOT NULL
#     GROUP BY 
#         DATEFROMPARTS(LEFT(CAST(time_period AS VARCHAR), 4), RIGHT(CAST(time_period AS VARCHAR), 2), 1),
#         UPPER(LEFT(LOWER(make), 1)) + LOWER(SUBSTRING(make, 2, LEN(make))),
#         hyperion_id

#     UNION ALL

#     -- OnOrder Snapshot: Pipeline
#     SELECT 
#         DATEFROMPARTS(LEFT(CAST(time_period AS VARCHAR), 4), RIGHT(CAST(time_period AS VARCHAR), 2), 1) AS AccountingMonth,
#         MarketName            = 'Pipeline',
#         StoreName             = 'Pipeline',
#         Department            = 'Pipeline',
#         StyleName             = 'Pipeline',
#         FuelType              = 'Pipeline',
#         UPPER(LEFT(LOWER(make), 1)) + LOWER(SUBSTRING(make, 2, LEN(make))) AS Make,
#         Model                 = 'Pipeline',
#         Trim                  = 'Pipeline',
#         Mileage               = CAST('' AS VARCHAR(50)),
#         Hyperion              = hyperion_id, 
#         DaysInInventoryAN     = CAST(0 AS INT),
#         DaysInInventoryStore  = CAST(0 AS INT),
#         Loaner                = CAST('' AS VARCHAR(50)),
#         MSRP_Adv              = CAST(0 AS MONEY),
#         AllocationGroup       = CAST('' AS VARCHAR(50)),
#         Color                 = CAST('' AS VARCHAR(50)),
#         Status                = CAST('' AS VARCHAR(50)),
#         Balance               = CAST(0 AS MONEY),
#         ExService             = 'Pipeline',
#         Table_Source          = 'Pipeline',
#         InvCount              = CAST(0 AS INT),
#         SumSold               = CAST(0 AS MONEY),
#         SumFrontGross         = CAST(0 AS MONEY),
#         SumFIGross            = CAST(0 AS MONEY),
#         SumIncentiveGross     = CAST(0 AS MONEY),
#         SumOVIGross           = CAST(0 AS MONEY),
#         InTransitCount        = CAST(0 AS INT),
#         COUNT(*) AS PipelineCount
#     FROM NDD_ADP_RAW.NDDUsers.vOnOrderInfo_SnapShot
#     WHERE 
#         an_status NOT IN ('ignore', 'cancelled', 'in transit')
#         AND DATEFROMPARTS(LEFT(CAST(time_period AS VARCHAR), 4), RIGHT(CAST(time_period AS VARCHAR), 2), 1) >= @StartDate
#         AND (
#             DATEFROMPARTS(LEFT(CAST(time_period AS VARCHAR), 4), RIGHT(CAST(time_period AS VARCHAR), 2), 1) BETWEEN @QuarterStart AND DATEADD(MONTH, -1, @CurrentMonth)
#             OR (
#                 RIGHT(CAST(time_period AS VARCHAR), 2) IN ('03','06','09','12')
#                 AND DATEFROMPARTS(LEFT(CAST(time_period AS VARCHAR), 4), RIGHT(CAST(time_period AS VARCHAR), 2), 1) < @QuarterStart
#             )
#         )
#         AND make IS NOT NULL
#     GROUP BY 
#         DATEFROMPARTS(LEFT(CAST(time_period AS VARCHAR), 4), RIGHT(CAST(time_period AS VARCHAR), 2), 1),
#         UPPER(LEFT(LOWER(make), 1)) + LOWER(SUBSTRING(make, 2, LEN(make))),
#         hyperion_id

#     UNION ALL

#     -- Daily Snapshot: Current Month In Transit
#     SELECT 
#         @CurrentMonth AS AccountingMonth,
#         MarketName            = 'CM In Transit',
#         StoreName             = 'CM In Transit',
#         Department            = 'CM In Transit',
#         StyleName             = 'CM In Transit',
#         FuelType              = 'CM In Transit',
#         UPPER(LEFT(LOWER(make), 1)) + LOWER(SUBSTRING(make, 2, LEN(make))) AS Make,
#         Model                 = 'CM In Transit',
#         Trim                  = 'CM In Transit',
#         Mileage               = CAST('' AS VARCHAR(50)),
#         Hyperion              = hyperion_id, 
#         DaysInInventoryAN     = CAST(0 AS INT),
#         DaysInInventoryStore  = CAST(0 AS INT),
#         Loaner                = CAST('' AS VARCHAR(50)),
#         MSRP_Adv              = CAST(0 AS MONEY),
#         AllocationGroup       = CAST('' AS VARCHAR(50)),
#         Color                 = CAST('' AS VARCHAR(50)),
#         Status                = CAST('' AS VARCHAR(50)),
#         Balance               = CAST(0 AS MONEY),
#         ExService             = 'CM In Transit',
#         Table_Source          = 'CM In Transit',
#         InvCount              = CAST(0 AS INT),
#         SumSold               = CAST(0 AS MONEY),
#         SumFrontGross         = CAST(0 AS MONEY),
#         SumFIGross            = CAST(0 AS MONEY),
#         SumIncentiveGross     = CAST(0 AS MONEY),
#         SumOVIGross           = CAST(0 AS MONEY),
#         COUNT(*) AS InTransitCount,
#         PipelineCount         = CAST(0 AS INT)
#     FROM NDD_ADP_RAW.NDDUsers.vOnOrderInfo_Dly_SnapShot
#     WHERE 
#         an_status = 'in transit'
#         AND update_date = @Yesterday
#         AND make IS NOT NULL
#     GROUP BY
#         hyperion_id,
#         UPPER(LEFT(LOWER(make), 1)) + LOWER(SUBSTRING(make, 2, LEN(make)))

#     UNION ALL

#     -- Daily Snapshot: Current Month Pipeline
#     SELECT 
#         @CurrentMonth AS AccountingMonth,
#         MarketName            = 'CM Pipeline',
#         StoreName             = 'CM Pipeline',
#         Department            = 'CM Pipeline',
#         StyleName             = 'CM Pipeline',
#         FuelType              = 'CM Pipeline',
#         UPPER(LEFT(LOWER(make), 1)) + LOWER(SUBSTRING(make, 2, LEN(make))) AS Make,
#         Model                 = 'CM Pipeline',
#         Trim                  = 'CM Pipeline',
#         Mileage               = CAST('' AS VARCHAR(50)),
#         Hyperion              = hyperion_id, 
#         DaysInInventoryAN     = CAST(0 AS INT),
#         DaysInInventoryStore  = CAST(0 AS INT),
#         Loaner                = CAST('' AS VARCHAR(50)),
#         MSRP_Adv              = CAST(0 AS MONEY),
#         AllocationGroup       = CAST('' AS VARCHAR(50)),
#         Color                 = CAST('' AS VARCHAR(50)),
#         Status                = CAST('' AS VARCHAR(50)),
#         Balance               = CAST(0 AS MONEY),
#         ExService             = 'CM Pipeline',
#         Table_Source          = 'CM Pipeline',
#         InvCount              = CAST(0 AS INT),
#         SumSold               = CAST(0 AS MONEY),
#         SumFrontGross         = CAST(0 AS MONEY),
#         SumFIGross            = CAST(0 AS MONEY),
#         SumIncentiveGross     = CAST(0 AS MONEY),
#         SumOVIGross           = CAST(0 AS MONEY),
#         InTransitCount        = CAST(0 AS INT),
#         COUNT(*) AS PipelineCount
#     FROM NDD_ADP_RAW.NDDUsers.vOnOrderInfo_Dly_SnapShot
#     WHERE 
#         an_status NOT IN ('ignore', 'cancelled', 'in transit')
#         AND update_date = @Yesterday
#         AND make IS NOT NULL
#     GROUP BY
#         hyperion_id,
#         UPPER(LEFT(LOWER(make), 1)) + LOWER(SUBSTRING(make, 2, LEN(make)))
# ),
# Combined AS (
#     SELECT * FROM SalesCTE
#     UNION ALL
#     SELECT * FROM InventoryCTE
#     UNION ALL
#     SELECT * FROM PipelineCTE
# )
# SELECT 
#     AccountingMonth,
#     MarketName,
#     StoreName,
#     Department,
#     StyleName,
#     FuelType,
#     Make,
#     Model,
#     Trim,
#     Mileage,
#     Hyperion,
#     DaysInInventoryAN,
#     DaysInInventoryStore,
#     Loaner,
#     MSRP_Adv,
#     AllocationGroup,
#     Color,
#     Status,
#     Balance,
#     ExService,
#     Table_Source,
#     InvCount,
#     SumSold,
#     SumFrontGross,
#     SumFIGross, 
#     SumIncentiveGross,
#     SumOVIGross,
#     InTransitCount,
#     PipelineCount
# FROM Combined
# WHERE AccountingMonth BETWEEN @StartDate AND @EndDate;"""
#     start_time = time.time()
    # try:
    #     with pyodbc.connect(
    #             'DRIVER={ODBC Driver 17 for SQL Server};'
    #             'SERVER=nddprddb01,48155;'
    #             'DATABASE=NDD_ADP_RAW;'
    #             'Trusted_Connection=yes;'
    #     ) as conn:
    #         df = pd.read_sql(Query, conn)
    # except Exception as e:
    #     print("❌ Connection failed:", e)
    #---------------------------------Load files to df--------------------------------------------------

    # File path
    file_path = r"C:\Development\PowerBI\Master_Dash\Portfolio.csv"

    # Load the CSV file
    df = pd.read_csv(file_path)

    ###FORMULAS
    # Extract Year from AccountingMonth (assuming format like '2025-08' or 'YYYY-MM')
    df['Year'] = df['AccountingMonth'].astype(str).str[:4]

    # Rename Columns to fit BI Model
    df = df.rename(columns={
        'Make': 'Brand',
        'AllocationGroup': 'Model',
        'AccountingMonth': 'Sales Date',
        'SumSold': 'Units Sold',
        'AverageSellingPrice': 'ASP',
        'InvCount': 'Inventory Levels',
        'DaysInInventoryStore': 'Days on Lot'})



    # Include/Exclude columns
    df = df[[
        'Brand', 'Model', 'Year', 'MarketName', 'StoreName',
        'Sales Date', 'Units Sold', 'SumFrontGross', 'SumCashPrice', 'Inventory Levels', 'Days on Lot'
    ]]

    # Save to a new CSV if needed
    df.to_csv(r"C:\Development\PowerBI\Master_Dash\Portfolio_final.csv", index=False)


    return

#run function
if __name__ == '__main__':
    
    Clean_csv()