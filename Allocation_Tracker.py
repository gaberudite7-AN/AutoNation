import xlwings as xw
import pandas as pd
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import shutil
import pyodbc
import time
import warnings
import os
import glob
warnings.filterwarnings("ignore", category=UserWarning, message=".*pandas only supports SQLAlchemy.*")

class AllocationTracker:
    def __init__(self, base_path):
        self.base_path = base_path
        self.daily_sales_filename = "Dynamic_Daily_Sales.xlsb"
        self.today = datetime.today()
        self.reference_date = None
        self.beginning_of_month_dt = None
        self.month_to_query = None

        # Excel files
        if self.today.day >= 6: # Use Regular Allocation File
            self.allocation_tracker_file = os.path.join(base_path, "Allocation_Tracker.xlsm")
            print("Using Regular Allocation File")
        else: # Use EOM Allocation File
            self.allocation_tracker_file = os.path.join(base_path, "Allocation_Tracker_EOM.xlsm")
            print("Using EOM Allocation File since day is less than 7th")
        self.dynamic_sor_file = r'W:\Corporate\Management Reporting Shared\Dynamic SOR.xlsb'

    def process_daily_sales_file(self):
        all_files = glob.glob(os.path.join(self.base_path, "*.xlsb"))
        latest_file = max(all_files, key=os.path.getmtime)
        print(f"Latest .xlsb file: {latest_file}")

        daily_files = glob.glob(os.path.join(self.base_path, "*Daily*.xlsb"))

        if len(daily_files) == 1:
            print(f"Only one file found: {daily_files[0]}")
            return daily_files[0]
        else:
            new_path = os.path.join(self.base_path, self.daily_sales_filename)
            shutil.copyfile(latest_file, new_path)
            os.remove(latest_file)
            print("Successfully removed old file and replaced current with new.")
            return new_path

    def calculate_dates(self):
        if self.today.day >= 6: # Use Current month for SSPR and Prior month for NDD
            self.reference_date = self.today
            first_day_this_month = self.today.replace(day=1)
            months_back = first_day_this_month - timedelta(days=30)
        else: # Use Previous month for SSPR and Two months prior for NDD
            self.reference_date = self.today.replace(day=1) - timedelta(days=1)
            first_day_this_month = self.today.replace(day=1)
            months_back = first_day_this_month - timedelta(days=60)

        self.beginning_of_month_dt = months_back.replace(day=1)
        self.month_to_query = self.reference_date.month
        self.beginning_of_month = f"{self.beginning_of_month_dt.month}/1/{self.beginning_of_month_dt.year}"

    def run_NDD_sql_queries(self, queries: dict):
        try:
            with pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=nddprddb01,48155;'
                'DATABASE=NDD_ADP_RAW;'
                'Trusted_Connection=yes;'
            ) as conn:
                start_time = time.time()
                results = {}
                for name, query in queries.items():
                    df = pd.read_sql(query, conn)
                    results[name] = df
                    elapsed = time.time() - start_time
                    print(f"Loaded {name} in {elapsed:.2f} seconds")
                return results
        except Exception as e:
            print("❌ Connection failed:", e)
            return {}

    def run_BAPRD_sql_queries(self, queries: dict):
        try:
            with pyodbc.connect(
                    r'DRIVER={ODBC Driver 17 for SQL Server};'
                    r'SERVER=BAPRDDB01\BAPRD,49174;'
                    r'DATABASE=SSPRv3;'
                    r'Trusted_Connection=yes;'
            ) as conn:
                start_time = time.time()
                results = {}
                for name, query in queries.items():
                    df = pd.read_sql(query, conn)
                    results[name] = df
                    elapsed = time.time() - start_time
                    print(f"Loaded {name} in {elapsed:.2f} seconds")
                return results
        except Exception as e:
            print("❌ Connection failed:", e)
            return {}

    def run_Marketing_queries(self, queries: dict):
        try:
            with pyodbc.connect(
                r'DRIVER={ODBC Driver 17 for SQL Server};'
                r'SERVER=S1WPVSQLMBI1,46160;'
                r'DATABASE=BIODS;'
                r'Trusted_Connection=yes;'
                r'Encrypt=yes;'
                r'TrustServerCertificate=yes;'
            ) as conn:
                start_time = time.time()
                results = {}
                for name, query in queries.items():
                    df = pd.read_sql(query, conn)
                    results[name] = df
                    elapsed = time.time() - start_time
                    print(f"Loaded {name} in {elapsed:.2f} seconds")
                return results
        except Exception as e:
            print("❌ Connection failed:", e)
            return {}

    def update_excel_NDD_SSPR_Data(self, allocation_wb, dataframes: dict):

        # Dump data into tabs
        allocation_wb.sheets['Store SSPR DATA'].range('P4:S10000').clear_contents()
        allocation_wb.sheets['Store SSPR DATA'].range('P3').options(index=False).value = dataframes.get('SSPR_df')

        allocation_wb.sheets['EV Store SSPR DATA'].range('S4:X10000').clear_contents()
        allocation_wb.sheets['EV Store SSPR DATA'].range('S3').options(index=False).value = dataframes.get('SSPR_EV_df')

        allocation_wb.sheets['EV Store SSPR DATA'].range('Z4:AA10000').clear_contents()
        allocation_wb.sheets['EV Store SSPR DATA'].range('Z3').options(index=False).value = dataframes.get('NDD_df_4')

        ndd_tab = allocation_wb.sheets['NDD DATA']
        ndd_tab.range('A4:F10000').clear_contents()
        ndd_tab.range('A3').options(index=False).value = dataframes.get('NDD_df')

        ndd_tab.range('H4:M10000').clear_contents()
        ndd_tab.range('H3').options(index=False).value = dataframes.get('NDD_EV_df')

        ndd_tab.range('O4:R10000').clear_contents()
        ndd_tab.range('O3').options(index=False).value = dataframes.get('NDD_df_3')


    def update_excel_Sales_Efficiency(self, allocation_wb, dataframes: dict):

        # Dump data into tabs
        allocation_wb.sheets['Sales Efficiency'].range('C4:J10000').clear_contents()
        allocation_wb.sheets['Sales Efficiency'].range('C3').options(index=False).value = dataframes.get('Sales_Efficiency_df')

    def run_macro_and_save_close(self, wb, macro_name="ExecuteMacros"):
        #wb = app.books.open(file_path)
        wb.macro(macro_name)()
        wb.save()
        wb.close()

    def update_sspr_history(self, sspr_history_df):
        formatted_date = (self.today.replace(day=1) - timedelta(days=1)).strftime("%Y%m")
        result_date = (self.today.replace(day=1) - timedelta(days=1)).replace(day=1)
        result_date_str = f"{result_date.month}/{result_date.day}/{result_date.year}"

        allocation_wb = xw.Book(self.allocation_tracker_file)
        history_tab = allocation_wb.sheets['SSPR Historic Data']

        existing_df = history_tab.range('A1').expand('down').resize(None, 9).options(pd.DataFrame, header=1).value
        existing_df.reset_index(inplace=True)
        filtered_df = existing_df[existing_df['Year/Month'].astype(int) != int(formatted_date)]

        sspr_history_df['Month'] = sspr_history_df['Month'].astype(int)
        sspr_history_df['Year'] = sspr_history_df['Year'].astype(int)
        sspr_history_df['Month_Format'] = sspr_history_df['Month'].astype(str).str.zfill(2)
        sspr_history_df['Year/Month'] = sspr_history_df['Year'].astype(str) + sspr_history_df['Month_Format']
        sspr_history_df['Accounting Month'] = pd.to_datetime(sspr_history_df['Year'].astype(str) + '-' + sspr_history_df['Month_Format'] + '-01')
        sspr_history_df = sspr_history_df[sspr_history_df['Year/Month'].astype(int) == int(formatted_date)].copy()
        sspr_history_df.drop(columns=['Month_Format'], inplace=True)

        updated_df = pd.concat([filtered_df, sspr_history_df], ignore_index=True)
        updated_df['Segment Filter'] = updated_df['Segment']

        history_tab.range('A2:J500000').clear_contents()
        history_tab.range('A2').options(index=False, header=False).value = updated_df

    def run_allocation_tracker(self):
        self.calculate_dates()


        # NDD queries
        NDD_query = f"""
    SELECT
    *

    FROM(

    SELECT					
    CONVERT(VARCHAR, AccountingMonth,101) AS AccountingMonth,								
    hyperion,					
    StoreName,					
    CASE
        WHEN Make IN ('CHRYSLER', 'DODGE','JEEP','RAM')
        THEN 'CDJR'
        WHEN Make in ('FORD', 'LINCOLN')
        THEN 'Ford'
        WHEN Make IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC')
        THEN 'GM'
        WHEN Make in ('HYUNDAI', 'GENESIS')
        THEN 'Hyundai'
        WHEN Make in ('JAGUAR', 'LAND ROVER')
        THEN 'JLR'
        Else Make
        END AS Brand,
    --AllocationGroup,
    SUM(CASE WHEN Type = 'Inv' THEN QTY ELSE 0 END) AS Inv,					
    SUM(CASE WHEN Type = 'Sales' THEN QTY ELSE 0 END) AS Sales					
                        
    FROM(					
                        
    SELECT					
    AccountingMonth,									
    hyperion,					
    StoreName,					
    Make,
    --AllocationGroup,
    'Inv' AS Type,					
    SUM(InventoryCount) AS QTY					
                        
    FROM					
    NDDUsers.vInventoryMonthEnd					
                        
    WHERE									
    AccountingMonth = '{self.beginning_of_month}'
    AND Department = '300'		
    --AND Make = 'Mercedes-Benz'
                        
    GROUP BY					
    AccountingMonth,									
    hyperion,					
    StoreName,					
    Make
    --AllocationGroup
                        
    HAVING 					
    SUM(InventoryCount) <> 0					
                        
    UNION ALL					
                        
    SELECT					
    AccountingMonth,								
    StoreHyperion,					
    StoreName,					
    VehicleMakeName,
    --AllocationGroup,
    'Sales' AS Type,					
    SUM(VehicleSoldCount) AS QTY					
                        
    FROM					
    NDDUsers.vSalesDetail_Vehicle					
                        
    WHERE										
    AccountingMonth = '{self.beginning_of_month}'			
    AND DepartmentName = 'NEW'					
    AND RecordSource = 'Accounting'
    --AND VehicleMakeName = 'Mercedes-Benz'
                
                        
    GROUP BY					
    AccountingMonth,								
    StoreHyperion,					
    StoreName,					
    VehicleMakeName
    --AllocationGroup
                        
    HAVING 					
    SUM(VehicleSoldCount) <> 0					
                        
    ) AS SQ					

    WHERE
    Make IS NOT NULL
        
    GROUP BY					
    AccountingMonth,									
    hyperion,					
    StoreName,					
    --AllocationGroup,
    CASE
        WHEN Make IN ('CHRYSLER', 'DODGE','JEEP','RAM')
        THEN 'CDJR'
        WHEN Make in ('FORD', 'LINCOLN')
        THEN 'Ford'
        WHEN Make IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC')
        THEN 'GM'
        WHEN Make in ('HYUNDAI', 'GENESIS')
        THEN 'Hyundai'
        WHEN Make in ('JAGUAR', 'LAND ROVER')
        THEN 'JLR'
        Else Make
        END

    ) SQ2

    WHERE
    Inv >5
        """

        NDD_EV_query = f"""
    SELECT					
    CONVERT(VARCHAR, AccountingMonth,101) AS AccountingMonth,								
    hyperion,					
    StoreName,					
    CASE
        WHEN Make IN ('CHRYSLER', 'DODGE','JEEP','RAM')
        THEN 'CDJR'
        WHEN Make in ('FORD', 'LINCOLN')
        THEN 'Ford'
        WHEN Make IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC')
        THEN 'GM'
        WHEN Make in ('HYUNDAI', 'GENESIS')
        THEN 'Hyundai'
        WHEN Make in ('JAGUAR', 'LAND ROVER')
        THEN 'JLR'
        Else Make
        END AS Brand,
    --AllocationGroup,
    SUM(CASE WHEN Type = 'Inv' THEN QTY ELSE 0 END) AS Inv,					
    SUM(CASE WHEN Type = 'Sales' THEN QTY ELSE 0 END) AS Sales					
                        
    FROM(					
                        
    SELECT					
    AccountingMonth,									
    hyperion,					
    StoreName,					
    Make,
    --AllocationGroup,
    'Inv' AS Type,					
    SUM(InventoryCount) AS QTY					
                        
    FROM					
    NDDUsers.vInventoryMonthEnd					
                        
    WHERE									
    AccountingMonth = '{self.beginning_of_month}'			
    AND Department = '300'		
    AND FuelType = 'Electric Fuel System'
    --AND Make = 'Mercedes-Benz'
                        
    GROUP BY					
    AccountingMonth,									
    hyperion,					
    StoreName,					
    Make
    --AllocationGroup
                        
    HAVING 					
    SUM(InventoryCount) <> 0					
                        
    UNION ALL					
                        
    SELECT					
    AccountingMonth,								
    StoreHyperion,					
    StoreName,					
    VehicleMakeName,
    --AllocationGroup,
    'Sales' AS Type,					
    SUM(VehicleSoldCount) AS QTY					
                        
    FROM					
    NDDUsers.vSalesDetail_Vehicle					
                        
    WHERE										
    AccountingMonth = '{self.beginning_of_month}'			
    AND DepartmentName = 'NEW'					
    AND RecordSource = 'Accounting'
    AND VehicleFuelType = 'Electric Fuel System'
    --AND VehicleMakeName = 'Mercedes-Benz'
                
                        
    GROUP BY					
    AccountingMonth,								
    StoreHyperion,					
    StoreName,					
    VehicleMakeName
    --AllocationGroup
                        
    HAVING 					
    SUM(VehicleSoldCount) <> 0					
                        
    ) AS Subquery					

    WHERE
    Make IS NOT NULL
        
    GROUP BY					
    AccountingMonth,									
    hyperion,					
    StoreName,					
    --AllocationGroup,
    CASE
        WHEN Make IN ('CHRYSLER', 'DODGE','JEEP','RAM')
        THEN 'CDJR'
        WHEN Make in ('FORD', 'LINCOLN')
        THEN 'Ford'
        WHEN Make IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC')
        THEN 'GM'
        WHEN Make in ('HYUNDAI', 'GENESIS')
        THEN 'Hyundai'
        WHEN Make in ('JAGUAR', 'LAND ROVER')
        THEN 'JLR'
        Else Make
        END    
    """

        NDD_EV_Flag_query = """
        SELECT
        AllocationGroup,
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

        NDD_query_3 = f"""
    SELECT					
    CONVERT(VARCHAR, AccountingMonth,101) AS AccountingMonth,								
    --hyperion,					
    --StoreName,					
    CASE
        WHEN Make IN ('CHRYSLER', 'DODGE','JEEP','RAM')
        THEN 'CDJR'
        WHEN Make in ('FORD', 'LINCOLN')
        THEN 'Ford'
        WHEN Make IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC')
        THEN 'GM'
        WHEN Make in ('HYUNDAI', 'GENESIS')
        THEN 'Hyundai'
        WHEN Make in ('JAGUAR', 'LAND ROVER')
        THEN 'JLR'
        Else Make
        END AS Brand,
    --AllocationGroup,
    SUM(CASE WHEN Type = 'Inv' THEN QTY ELSE 0 END) AS Inv,					
    SUM(CASE WHEN Type = 'Sales' THEN QTY ELSE 0 END) AS Sales					
                        
    FROM(					
                        
    SELECT					
    AccountingMonth,									
    --hyperion,					
    --StoreName,					
    Make,
    --AllocationGroup,
    'Inv' AS Type,					
    SUM(InventoryCount) AS QTY					
                        
    FROM					
    NDDUsers.vInventoryMonthEnd					
                        
    WHERE									
    AccountingMonth = '{self.beginning_of_month}'			
    AND Department = '300'		
    --AND Make = 'Mercedes-Benz'
                        
    GROUP BY					
    AccountingMonth,									
    --hyperion,					
    --StoreName,					
    Make
    --AllocationGroup
                        
    HAVING 					
    SUM(InventoryCount) <> 0					
                        
    UNION ALL					
                        
    SELECT					
    AccountingMonth,								
    --StoreHyperion,					
    --StoreName,					
    VehicleMakeName,
    --AllocationGroup,
    'Sales' AS Type,					
    SUM(VehicleSoldCount) AS QTY					
                        
    FROM					
    NDDUsers.vSalesDetail_Vehicle					
                        
    WHERE										
    AccountingMonth = '{self.beginning_of_month}'			
    AND DepartmentName = 'NEW'					
    AND RecordSource = 'Accounting'
    --AND VehicleMakeName = 'Mercedes-Benz'
                
                        
    GROUP BY					
    AccountingMonth,								
    --StoreHyperion,					
    --StoreName,					
    VehicleMakeName
    --AllocationGroup
                        
    HAVING 					
    SUM(VehicleSoldCount) <> 0					
                        
    ) AS Subquery					

    WHERE
    Make IS NOT NULL
        
    GROUP BY					
    AccountingMonth,									
    --hyperion,					
    --StoreName,					
    --AllocationGroup,
    CASE
        WHEN Make IN ('CHRYSLER', 'DODGE','JEEP','RAM')
        THEN 'CDJR'
        WHEN Make in ('FORD', 'LINCOLN')
        THEN 'Ford'
        WHEN Make IN ('BUICK', 'CADILLAC', 'CHEVROLET', 'GMC')
        THEN 'GM'
        WHEN Make in ('HYUNDAI', 'GENESIS')
        THEN 'Hyundai'
        WHEN Make in ('JAGUAR', 'LAND ROVER')
        THEN 'JLR'
        Else Make
        END
        """
        
        # Run NDD queries
        ndd_queries = {
            "NDD_df": NDD_query,    
        # Replace with actual query
            "NDD_EV_df": NDD_EV_query,
            "NDD_df_4": NDD_EV_Flag_query,
            "NDD_df_3": NDD_query_3
        }
        ndd_results = self.run_NDD_sql_queries(ndd_queries)

        # SSPR queries
        SSPR_query = f"""
        SELECT
        --Year,
        --Month,
        ParentBrand,
        Segment,
        Hyperion,
        EarnedM1,
        CASE 
            WHEN commitM1 > earnedM1 
            THEN earnedM1 
            ELSE commitM1 END AS CommitM1

        FROM(

        SELECT --ParentBrand
        
        Year,
        
        Month,
        
        Hyperion,
        
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
        
        --Make,
        
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
        
        DC.[Year] = YEAR(GETDATE())
        AND DC.Month = {self.month_to_query}
        
        --AND
        
        --DC.[Month] in ( Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-5, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-4, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-3, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-2, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0))  )

        UNION

        SELECT
        
        DF.[Year], DF.[Month], DF.[DealerVehicleID], DF.[Mth], DF.[ColumnCD], DF.[EnteredValue], DF.[StatusID]
        
        FROM
        
        [SSPRv3].[dbo].[DealerForecasts] AS DF
        
        WHERE
        
        DF.[Year] = YEAR(GETDATE())
        AND DF.Month = {self.month_to_query}
        
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
        
        DC.[Year] = YEAR(GETDATE())
        AND DC.Month = {self.month_to_query}
        
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
        
        --Make,
        
        --Model,
        
        Hyperion,
        
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

        ) AS FINALSQ

        WHERE
        Segment IS NOT NULL
            """
        SSPR_EV_query = f"""
    SELECT
    --Year,
    --Month,
    ParentBrand,
    Segment,
    Hyperion,
    Model,
    EarnedM1,
    CASE 
        WHEN commitM1 > earnedM1 
        THEN earnedM1 
        ELSE commitM1 END AS CommitM1

    FROM(

    SELECT --ParentBrand
    
    Year,
    
    Month,
    
    Hyperion,
    
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
    
    --Make,
    
    Model,
    
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
    
    DC.[Year] = YEAR(GETDATE())
    AND DC.Month = {self.month_to_query}
    
    --AND
    
    --DC.[Month] in ( Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-5, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-4, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-3, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-2, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-1, 0)), Month(DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0))  )

    UNION

    SELECT
    
    DF.[Year], DF.[Month], DF.[DealerVehicleID], DF.[Mth], DF.[ColumnCD], DF.[EnteredValue], DF.[StatusID]
    
    FROM
    
    [SSPRv3].[dbo].[DealerForecasts] AS DF
    
    WHERE
    
    DF.[Year] = YEAR(GETDATE())
    AND DF.Month = {self.month_to_query}
    
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
    
    DC.[Year] = YEAR(GETDATE())
    AND DC.Month = {self.month_to_query}
    
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
    
    --Make,
    
    Model,
    
    Hyperion,
    
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

    ) AS FINALSQ

    WHERE
    Segment IS NOT NULL
    """

        SSPR_History_query = f"""
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

        # Run SSPR queries
        sspr_queries = {
            "SSPR_df": SSPR_query,
            "SSPR_EV_df": SSPR_EV_query,
            "SSPR_history_df": SSPR_History_query
        }
        sspr_results = self.run_BAPRD_sql_queries(sspr_queries)

        # Sales Efficiency Query
        Sales_efficiency_query = f"""
            SELECT 	 
            [ORG] AS STORE_HYPERION_ID
            ,MONTH([MONTH]) AS [MONTH]
            ,YEAR([MONTH]) AS [YEAR]
            ,[SCENARIO]
            ,[MANUFACTURE]
            ,[DEPARTMENT]
            ,[ACCOUNT]
            ,[VALUE]
                
            FROM [BISTG].[STGNDD].[STG_BA_VSOR_ESSBASE_EXTRACT_TOT]
                
            WHERE 	 
            1=1 	 
            AND ACCOUNT = 'SALES EFFICIENCY %'
            AND CONSOLIDATEDSCENARIO = 'ACTUAL'
            AND DEPARTMENT ='NEW'
            AND [MONTH] = '{self.beginning_of_month}'
            AND MANUFACTURE ='PPR'
        """
        sales_efficiency_queries = {
            "Sales_Efficiency_df": Sales_efficiency_query
        }
        sales_efficiency_results = self.run_Marketing_queries(sales_efficiency_queries)
        sspr_results.update(sales_efficiency_results)
        # Combine results and update Excel
        all_data = {**ndd_results, **sspr_results, **sales_efficiency_results}
        
        # Open Excel files    
        app = xw.App(visible=True)    
        allocation_wb = app.books.open(self.allocation_tracker_file)
        daily_sales_wb = app.books.open(os.path.join(self.base_path, self.daily_sales_filename))
        dynamic_sor_wb = app.books.open(self.dynamic_sor_file, read_only=True, ignore_read_only_recommended=True)
        try:
            self.update_excel_NDD_SSPR_Data(allocation_wb, all_data)
            self.update_sspr_history(all_data.get('SSPR_history_df'))
            self.update_excel_Sales_Efficiency(allocation_wb, all_data)
            self.run_macro_and_save_close(allocation_wb)
        finally:
            # Close other workbooks without saving
            if daily_sales_wb:
                daily_sales_wb.close()
            if dynamic_sor_wb:
                dynamic_sor_wb.close()
            app.quit()

if __name__ == "__main__":

    base_path = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker'
    tracker = AllocationTracker(base_path)
    tracker.process_daily_sales_file()
    tracker.run_allocation_tracker()