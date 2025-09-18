# Imports
import xlwings as xw
import pandas as pd
import os
from datetime import datetime, timedelta
import shutil
import numpy as np
import gc
import pyodbc
from dateutil.relativedelta import relativedelta
import time
import glob
import warnings
warnings.filterwarnings("ignore", category=UserWarning, message=".*pandas only supports SQLAlchemy.*")


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


#main code
def Used_Car_Update():

    # Get current date
    today = datetime.today()

    # Begin timer
    start_time = time.time()

    if today.day >= 7:
        # Use current month
        reference_date = today
        formatted_date = reference_date.strftime("%Y%m")  # e.g. "202506"
        beginning_of_month_dt = reference_date.replace(day=1)
    else:
        # Use previous month
        reference_date = today.replace(day=1) - timedelta(days=1)  # last day of previous month
        formatted_date = reference_date.strftime("%Y%m")
        beginning_of_month_dt = reference_date.replace(day=1)

    # Format the dates
    current_day = f"{today.month}/{today.day}/{today.year}"  # actual current day
    beginning_of_month = f"{beginning_of_month_dt.month}/1/{beginning_of_month_dt.year}"  # e.g. "5/1/2025"

    # Get 3 months before the beginning of selected month
    three_months_prior_dt = beginning_of_month_dt - relativedelta(months=3)
    three_months_prior = f"{three_months_prior_dt.month}/1/{three_months_prior_dt.year}"

    # Optional: Print results for debugging
    print(f"Formatted Date (yyyymm): {formatted_date}")
    print(f"Current Day: {current_day}")
    print(f"Beginning of Month: {beginning_of_month}")
    print(f"3 Months Prior: {three_months_prior}")


################################################################################################################
    '''RUN NDD QUERIES'''
################################################################################################################

    # Run SQL queries using SQL Alchemy and dump into Data tab
    InvData_query = f"""
    SELECT month(AccountingMonth) as Month,
    --RegionName,
    --MarketName,
    /*Hyperion,
    StoreName,
    Vin,
    Year, 
    Make,
    Model,
    Mileage,
    Balance,
    PriceTier_93 AS Website,
    DaysInInventoryAN,
    DaysInInventoryStore,*/
    sum(InventoryCount) InvCount,
    --Status,

    case when DaysInInventoryAN is null then 'N/A'
    when DaysInInventoryAN <45 then 'Under 45 Days'
    when DaysInInventoryAN >=45 then 'Over/At 45 Days'
    else 'N/A' end as '45DayBucket',

    case when DaysInInventoryAN is null then 'N/A'
    when DaysInInventoryAN <120 then 'Under 120 Days'
    when DaysInInventoryAN >=120 then 'Over/At 120 Days'
    else 'N/A' end as '120DayBucket',

    case when PriceTier_93 is null then 'N/A'
    when PriceTier_93 <1 then 'N/A'
    when PriceTier_93 <20001 then '1:0-20,000'
    when PriceTier_93 <40001 then '2:20,001 - 40,000'
    else '3:Over 40,001' end as 'PriceBucket'

    FROM nddusers.vInventoryMonthEnd

    WHERE AccountingMonth >= '{three_months_prior}'
    AND RegionName <> 'AND Corporate Management'
    AND MarketName <> 'Eastern Market 99'
    And MarketName <> 'western market 99'
    And MarketName <> 'WS Auctions'
    AND Balance > 0
    AND InventoryCount > 0 
    AND Department = '320'
    and Status = 'S'

    Group by AccountingMonth,
    --RegionName,
    --MarketName, 
    case when DaysInInventoryAN is null then 'N/A'
    when DaysInInventoryAN <45 then 'Under 45 Days'
    when DaysInInventoryAN >=45 then 'Over/At 45 Days'
    else 'N/A' end,

    case when DaysInInventoryAN is null then 'N/A'
    when DaysInInventoryAN <120 then 'Under 120 Days'
    when DaysInInventoryAN >=120 then 'Over/At 120 Days'
    else 'N/A' end,

    case when PriceTier_93 is null then 'N/A'
    when PriceTier_93 <1 then 'N/A'
    when PriceTier_93 <20001 then '1:0-20,000'
    when PriceTier_93 <40001 then '2:20,001 - 40,000'
    else '3:Over 40,001' end

    Having sum(InventoryCount) <> 0
        """
    
    SalesData_query = f"""
    DECLARE @STARTDATE AS DATE
    DECLARE @ENDDATE AS DATE

    SET @STARTDATE = '{three_months_prior}'
    SET @ENDDATE   = '{beginning_of_month}'


    Select Year, Month, --Market, Market2,--Hyperion, Store, 
    Source, case when source in ('Auction/Rental','Other') then 'External' else 'Internal' end as SourceGroup,
    PriceBand, AN_Age as AgeBucket,
    sum(SoldCount) as SoldCount, sum(BaseGross) as BaseGross, sum(ICGross) as ICGross, sum(CFSGross) as CFSGross, sum(OVIGross) as OVIGross, sum(CashPrice) as CashPrice, sum(TargetPrice) as TargetPrice, sum(ReconCost) as ReconCost,
    --sum(SoldCountAdj) as SoldCountWithValidTarget,
    sum(InvCount) as InvCount, sum(InvPrice) as InvPrice, sum(InvBalance) as InvBalance, sum(InvTargetPrice) as InvTargetPrice, sum(InvOver45) as InvOver45, sum(InvOver120) as InvOver120
    --, sum(InventoryCountAdj) as InvCountwithValidTarget


    From 

    (Select year(accountingmonth) as Year,
    month(Accountingmonth) as Month,
    MarketName as Market,
    Case when left(storename,14) = 'AutoNation USA' then 'AN USA' else 'Franchise' end as 'Market2',
    StoreHyperion as Hyperion,
    StoreName as Store,
    VIN,
    sum(vehiclesoldcount) as SoldCount,
    sum(frontgross) as BaseGross,
    sum(wholesaleintercompanygross) as ICGross,
    sum(FIGross) as CFSGross,
    sum(OVI) as OVIGross,
    sum(case when ReconCostOriginatingStore is null or vehiclesoldcount <> 1 then 0 else ReconCostOriginatingStore end) + sum(case when NonROAmount is null or vehiclesoldcount <> 1 then 0 else NonROAmount end)  as ReconCost,
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
        ELSE 'Other' END AS Source,
    case when max(abs(VehicleAgeAN)) between 0 and 45 then '0-45' 
    when max(abs(VehicleAgeAN)) between 45 and 120 then '46-120' 
    when max(abs(VehicleAgeAN)) between 120 and 1200 then '120+' else 'Adjustments' end as AN_Age,
    case when max(abs(VehicleAge)) between 0 and 45 then '0-45' 
    when max(abs(VehicleAge)) between 45 and 120 then '46-120' 
    when max(abs(VehicleAge)) between 120 and 1200 then '120+' else 'Adjustments' end as StoreAge,
    Case when max(abs(FrontRevenue)) between 1 and 20000 then '$0-$20k'
    when max(abs(FrontRevenue)) between 20000 and 40000 then '$20-$40k'
    when max(abs(FrontRevenue)) between 40000 and 1000000 then '$40k+'
    when max(abs(Wholesaleintercompanyrevenue)) between 1 and 20000 then '$0-$20k'
    when max(abs(Wholesaleintercompanyrevenue)) between 20000 and 40000 then '$20-$40k'
    when max(abs(Wholesaleintercompanyrevenue)) between 40000 and 1000000 then '$40k+' else 'Adjustments' end as PriceBand,

    case when sum(vehiclesoldcount) = 0 then 0 when max(abs(FrontRevenue)) is null then 0 when max(abs(FrontRevenue)) = 0 then 0
    when max(abs(TargetPrice)) is null then 0 when max(abs(TargetPrice)) = 0 then 0  when max(abs(FrontRevenue)) > (max(abs(TargetPrice)) + 10000) then 0
    when max(abs(TargetPrice)) < 1000 then 0 when max(abs(FrontRevenue)) < 1000 then 0 when (max(abs(FrontRevenue)) + 10000) < max(abs(TargetPrice)) then 0
    else max(abs(FrontRevenue))  end as CashPrice,

    case when sum(vehiclesoldcount) = 0 then 0 when max(abs(FrontRevenue)) is null then 0 when max(abs(FrontRevenue)) = 0 then 0
    when max(abs(TargetPrice)) is null then 0 when max(abs(TargetPrice)) = 0 then 0  when max(abs(FrontRevenue)) > (max(abs(TargetPrice)) + 10000) then 0
    when max(abs(TargetPrice)) < 1000 then 0 when max(abs(FrontRevenue)) < 1000 then 0 when (max(abs(FrontRevenue)) + 10000) < max(abs(TargetPrice)) then 0
    else max(abs(TargetPrice)) end as TargetPrice,

    case when sum(vehiclesoldcount) = 0 then 0 when max(abs(FrontRevenue)) is null then 0 when max(abs(FrontRevenue)) = 0 then 0
    when max(abs(TargetPrice)) is null then 0 when max(abs(TargetPrice)) = 0 then 0  when max(abs(FrontRevenue)) > (max(abs(TargetPrice)) + 10000) then 0
    when max(abs(TargetPrice)) < 1000 then 0 when max(abs(FrontRevenue)) < 1000 then 0 when (max(abs(FrontRevenue)) + 10000) < max(abs(TargetPrice)) then 0
    else max(abs(vehiclesoldcount)) end as SoldCountAdj,

    0 as InvCount,
    0 as InvPrice,
    0 as InvBalance,
    0 as InvTargetPrice,
    0 as InventoryCountAdj,
    0 as InvOver45,
    0 as InvOver120

    From NDDUsers.vSalesDetail_Vehicle

    Where departmentname = 'Used'
    and accountingmonth between @Startdate and @Enddate
    and RegionName = 'Dealership'
    --and MarketName not in ('Market 98','Market 97')
    and recordsource = 'Accounting'

    Group by year(accountingmonth),
    month(Accountingmonth),
    MarketName,
    Case when left(storename,14) = 'AutoNation USA' then 'AN USA' else 'Franchise' end,
    StoreHyperion,
    StoreName,
    VIN,
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
        ELSE 'Other' END

    Having sum(vehiclesoldcount) <> 0 or sum(frontgross) <> 0 or sum(figross) <> 0 or sum(wholesaleintercompanygross) <> 0 or sum(OVI) <> 0

    union all


    Select year(accountingmonth) as Year,
    month(Accountingmonth) as Month,
    MarketName as Market,
    Case when left(storename,14) = 'AutoNation USA' then 'AN USA' else 'Franchise' end as 'Market2',
    Hyperion as Hyperion,
    StoreName as Store,
    VIN,
    0 as SoldCount,
    0 as BaseGross,
    0 as ICGross,
    0 as CFSGross,
    0 as OVIGross,
    0 as ReconCost,
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
        ELSE 'Other' END AS Source,
    case when max(abs(DaysInInventoryAN)) between 0 and 45 then '0-45' 
    when max(abs(DaysInInventoryAN)) between 45 and 120 then '46-120' 
    when max(abs(DaysInInventoryAN)) between 120 and 1200 then '120+' else 'Adjustments' end as AN_Age,
    case when max(abs(DaysInInventoryStore)) between 0 and 45 then '0-45' 
    when max(abs(DaysInInventoryStore)) between 45 and 120 then '46-120' 
    when max(abs(DaysInInventoryStore)) between 120 and 1200 then '120+' else 'Adjustments' end as StoreAge,
    Case when max(abs(InternetPrice)) between 1 and 20000 then '$0-$20k'
    when max(abs(InternetPrice)) between 20000 and 40000 then '$20-$40k'
    when max(abs(InternetPrice)) between 40000 and 1000000 then '$40k+'
    when max(abs(Pricetier_93)) between 1 and 20000 then '$0-$20k'
    when max(abs(Pricetier_93)) between 20000 and 40000 then '$20-$40k'
    when max(abs(Pricetier_93)) between 40000 and 1000000 then '$40k+' 
    when max(abs(Balance)) between 1 and 20000 then '$0-$20k'
    when max(abs(Balance)) between 20000 and 40000 then '$20-$40k'
    when max(abs(Balance)) between 40000 and 1000000 then '$40k+' else 'Adjustments' end as PriceBand,
    0 as CashPrice,
    0 as TargetPrice,
    0 as SoldCountAdj,
    sum(InventoryCount) as InvCount,

    case when sum(InventoryCount) = 0 then 0 when max(abs(InternetPrice)) is null then 0 when max(abs(InternetPrice)) = 0 then 0
    when max(abs(TargetPrice)) is null then 0 when max(abs(TargetPrice)) = 0 then 0 when max(abs(Balance)) is null then 0 when max(abs(Balance)) = 0 then 0
    when max(abs(InternetPrice)) > (max(abs(TargetPrice)) + 10000) then 0
    when max(abs(TargetPrice)) < 1000 then 0 when max(abs(InternetPrice)) < 1000 then 0 when (max(abs(InternetPrice)) + 10000) < max(abs(TargetPrice)) then 0
    else max(abs(InternetPrice))  end as InvPrice,

    case when sum(InventoryCount) = 0 then 0 when max(abs(InternetPrice)) is null then 0 when max(abs(InternetPrice)) = 0 then 0
    when max(abs(TargetPrice)) is null then 0 when max(abs(TargetPrice)) = 0 then 0 when max(abs(Balance)) is null then 0 when max(abs(Balance)) = 0 then 0
    when max(abs(InternetPrice)) > (max(abs(TargetPrice)) + 10000) then 0
    when max(abs(TargetPrice)) < 1000 then 0 when max(abs(InternetPrice)) < 1000 then 0 when (max(abs(InternetPrice)) + 10000) < max(abs(TargetPrice)) then 0
    else max(abs(Balance))  end as InvBalance,

    case when sum(InventoryCount) = 0 then 0 when max(abs(InternetPrice)) is null then 0 when max(abs(InternetPrice)) = 0 then 0
    when max(abs(TargetPrice)) is null then 0 when max(abs(TargetPrice)) = 0 then 0 when max(abs(Balance)) is null then 0 when max(abs(Balance)) = 0 then 0
    when max(abs(InternetPrice)) > (max(abs(TargetPrice)) + 10000) then 0
    when max(abs(TargetPrice)) < 1000 then 0 when max(abs(InternetPrice)) < 1000 then 0 when (max(abs(InternetPrice)) + 10000) < max(abs(TargetPrice)) then 0
    else max(abs(TargetPrice)) end as InvTargetPrice,

    case when sum(InventoryCount) = 0 then 0 when max(abs(InternetPrice)) is null then 0 when max(abs(InternetPrice)) = 0 then 0
    when max(abs(TargetPrice)) is null then 0 when max(abs(TargetPrice)) = 0 then 0 when max(abs(Balance)) is null then 0 when max(abs(Balance)) = 0 then 0
    when max(abs(InternetPrice)) > (max(abs(TargetPrice)) + 10000) then 0
    when max(abs(TargetPrice)) < 1000 then 0 when max(abs(InternetPrice)) < 1000 then 0 when (max(abs(InternetPrice)) + 10000) < max(abs(TargetPrice)) then 0
    else max(abs(InventoryCount)) end as InventoryCountAdj,

    case when case when max(abs(DaysInInventoryAN)) > 45 then sum(InventoryCount) end is null then 0 else case when max(abs(DaysInInventoryAN)) > 45 then sum(InventoryCount) end end as InvOver45,
    case when case when max(abs(DaysInInventoryAN)) > 120 then sum(InventoryCount) end is null then 0 else case when max(abs(DaysInInventoryAN)) > 120 then sum(InventoryCount) end end as InvOver120


    From NDDUsers.vInventoryMonthEnd

    Where department = '320'
    and accountingmonth between @Startdate and @Enddate
    and RegionName = 'Dealership'
    and status = 'S'
    and MarketName not in ('Market 98','Market 97')

    Group by year(accountingmonth),
    month(Accountingmonth),
    MarketName,
    Case when left(storename,14) = 'AutoNation USA' then 'AN USA' else 'Franchise' end,
    Hyperion,
    StoreName,
    VIN,
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
        ELSE 'Other' END

    Having sum(InventoryCount) <> 0 



    ) A

    Group by Year, Month, --Market, Market2,--Hyperion, Store, 
    Source, case when source in ('Auction/Rental','Other') then 'External' else 'Internal' end, PriceBand, AN_Age

    """

    New_and_Used_Sales_query = f"""
    Select month,
    year,
    Sum(vehiclesold) as vehiclesold,
    Sum(trades) as Trades

    from (

    Select 
    month(Accountingmonth) as Month,
    year(accountingmonth) as Year,
    sum(VehicleSoldCount) as VehicleSold, 
    case when Trade1Vin is not null then sum(vehiclesoldcount) else 0 end as Trades

    From NDDUsers.vSalesDetail_Vehicle

    Where
    accountingmonth between '{three_months_prior}' and '{beginning_of_month}'
    and RegionName = 'Dealership'
    and RecordSource = 'Accounting'

    Group by 
    month(Accountingmonth),
    year(accountingmonth),
    Trade1Vin

    Having sum(VehicleSoldCount) <> 0 ) A

    group by month,year

    order by Year, Month asc
    """

    Wholesale_Data_query = f"""
    Select 
    month(Accountingmonth) as Month,
    year(accountingmonth) as Year,
    sum(WholesaleAuctionCount) as WholesaleSold

    From NDDUsers.vSalesDetail_Vehicle

    Where departmentname = 'Used'
    and accountingmonth between '{three_months_prior}' and '{beginning_of_month}'
    and RegionName = 'Dealership'
    --and MarketName not in ('Market 98','Market 97')

    Group by 
    month(Accountingmonth),
    year(accountingmonth)

    Having sum(WholesaleAuctionCount) <> 0 

    order by Year, Month asc
    """

    TargetPrice_query = f"""
    With temptable1 as (Select ROW_NUMBER() Over (Partition by a.VIN, hyperion order by SnapshotDate asc) as RN,
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
    , Inventorysourcename, VIN, CONCAT(vin,hyperion) as HyperionVINFlag, targetprice, balance, OriginalACV

    From nddusers.vInventory_Daily_Snapshot a

    Where department = '320'
    and snapshotdate between '7/15/2024' and '{current_day}' and status <> 'G' 
    and targetprice is not null and targetprice <> 0 and balance is not null and balance >= OriginalACV
    ) 

    Select VIN, targetprice, OriginalACV, balance

    from temptable1 a

    where rn = 1

    group by VIN, targetprice, OriginalACV, balance

    order by vin desc
    """

    Org_Map_query = """
    SELECT  CAST ( a.hyperion_id AS CHAR(4)) as StoreHyperion
,a.entity_name as StoreName, r.entity_name as RegionName
,m.entity_name as MarketName
,case when [CC_Segment] is null then 'Other Entity' else [CC_Segment] end as Segment 
,case when [CC_Dominant_Brand] is null then 'Other Entity' else [CC_Dominant_Brand] end as BrandGroup
,case when [CC_Dominant_Brand] is null then 'Other Entity' when b.MAN_NAME='Acura' then 'Other Imports' when [CC_Dominant_Brand] ='Other Luxury' then b.MAN_NAME else [CC_Dominant_Brand] end as BrandGroup2
,case when b.CC_Dominant_Brand in ('Chrysler', 'Ford','GM') then b.CC_Dominant_Brand else b.MAN_NAME end as Brand --Used to Match DOC setup for brands
, m.hyperion_id as MarketHyperion, r.hyperion_id as RegionHyperion, b.mfrGroup, a.City, a.[State]
,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE( REPLACE(REPLACE( REPLACE( REPLACE (REPLACE
(a.entity_name , 'AutoNation' , 'AN' ) , 'Mercedes-Benz' , 'MB'), 'Chrysler ', 'C')
, 'Dodge ', 'D'),'Jeep ', 'J'), 'Ram ', 'R'),'Chevrolet', 'Chevy'), 'Volkswagen', 'VW')
,'Toyota Scion', 'Toyota'),'BMW of Houston North in The Woodlands / Mini of The Woodlands', 'BMW and MINI N The Woodlands')
,'North', 'N'),'BMW of Dallas and MINI of Dallas', 'BMW and MINI of Dallas') as StoreNameShort
,'NA' as GeoArea
FROM [NDD_ADP_RAW].[dbo].[ENTITIES] (nolock) a
join [NDD_ADP_RAW].[dbo].[Manufacturers] (nolock) b on a.manufacturer_code = b.man_id
join [NDD_ADP_RAW].[dbo].[ENTITIES] (nolock)  m  on a.parent_id = m.entity_id 
join [NDD_ADP_RAW].[dbo].[ENTITIES] (nolock) r on m.parent_ID = r.Entity_ID
--join (SELECT distinct [StoreHyperion]
--,[VOPSMarketName] FROM [NDD_ADP_RAW].[NDDUsers].[vEntityMake] (nolock)) g on a.hyperion_id = g.StoreHyperion
where 
a.entity_type = 'h' and 
a.[status] = 'open' and r.entity_name <> 'corporate'
--and a.manufacturer_code not in (175,199,200)
"""

    try:
        with pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=nddprddb01,48155;'
                'DATABASE=NDD_ADP_RAW;'
                'Trusted_Connection=yes;'
        ) as conn:
            InvData_df = pd.read_sql(InvData_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in InvData_df in {elapsed_time:.2f} seconds")
            SalesData_df = pd.read_sql(SalesData_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in SalesData_df in {elapsed_time:.2f} seconds")
            New_and_Used_Sales_df = pd.read_sql(New_and_Used_Sales_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in New_and_Used_Sales_df in {elapsed_time:.2f} seconds")
            Wholesale_Data_df = pd.read_sql(Wholesale_Data_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in Wholesale_Data_df in {elapsed_time:.2f} seconds")
            TargetPrice_df = pd.read_sql(TargetPrice_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in TargetPrice_df in {elapsed_time:.2f} seconds")
            Org_Map_df = pd.read_sql(Org_Map_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in Org_Map_df in {elapsed_time:.2f} seconds")

    except Exception as e:
        print("❌ Connection failed:", e)


################################################################################################################
    '''RUN MARKETING QUERIES'''
################################################################################################################

    AssociateData_Query = '''
    WITH StartDates AS (
        SELECT 
            CAST(DATEADD(MONTH, -3, DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)) AS DATE) AS START_DATE
    ),
    WBYC_DATA1 AS (
        SELECT
            YEAR(ACTIVITY_DATE) AS [YEAR],
            MONTH(ACTIVITY_DATE) AS [MONTH],
            R.REGION_NAME,
            CASE WHEN D.MARKET_NAME = 'AUTONATION DIRECT FLORIDA' THEN 'CAAD' ELSE D.MARKET_NAME END AS MARKET_NAME,
            C.STORE_HYPERION_ID,
            C.STORE_HYPERION_ID + '-' + C.STORE_NAME AS STORE,
            CASE 
                WHEN [WBYC_REPORTINGGROUP] IN ('TRADE OFFER (VOI)','RANGE TRADE (VOI)') THEN 'NON-WBYC LEAD'
                WHEN [WBYC_REPORTINGGROUP] IN ('SELL ONLY','RANGE - SELL ONLY') AND PROVIDER_NAME <> 'WBYC-SHOWROOM' THEN 'AN.COM SELL ONLY LEAD'
                WHEN [WBYC_REPORTINGGROUP] IN ('UNKNOWN') THEN 'NON-WBYC LEAD'
                ELSE [WBYC_REPORTINGGROUP]
            END AS [WBYC_REPORTINGGROUP],
            CASE 
                WHEN [WBYC_REPORTINGGROUP] IN ('SELL ONLY','RANGE - SELL ONLY') AND SUBPROVIDER_NAME IN ('WBYC Sell Only') THEN 'WBYC-AN.com'
                WHEN [WBYC_REPORTINGGROUP] IN ('SELL ONLY','RANGE - SELL ONLY') AND SUBPROVIDER_NAME IN ('Dealer Website – WBYC') THEN 'WBYC-OEM Sites'
                ELSE NULL 
            END AS [WBYC_REPORTINGGROUP_AN],
            MANUFACTURER_NAME,
            BPMS_ID,
            CAST(BPMS_MANAGER_NAME AS VARCHAR) AS BPMS_MANAGER,
            CAST(BPMS_ASSOCIATE1_NAME AS VARCHAR) AS BPMS_ASSOCIATE_1,
            CAST(BPMS_ASSOCIATE2_NAME AS VARCHAR) AS BPMS_ASSOCIATE_2,
            SUM([PURCHASE_COUNT]) AS PURCHASES,
            S.GROUP4 AS POD
        FROM
            BIDM.[FACT].[FACT_WBYC_TRAFFIC_DETAIL] A
            LEFT JOIN BIDM.DIM.DIM_STORE C ON C.STORE_EID = A.STORE_EID
            LEFT JOIN BIDM.DIM.DIM_MARKET D ON D.MARKET_EID = C.MARKET_EID
            LEFT JOIN BIDM.DIM.DIM_PROVIDER P ON P.PROVIDER_KEY = A.PROVIDER_KEY
            LEFT JOIN BIDM.DIM.DIM_REGION R ON R.REGION_EID = C.REGION_EID
            LEFT JOIN BIDM.DIM.DIM_MANUFACTURER M ON M.MANUFACTURER_KEY = C.MANUFACTURER_KEY
            LEFT JOIN [BIDM].[dim].[DIM_SUBPROVIDER] SP ON SP.SUBPROVIDER_PKEY = A.SUBPROVIDER_PKEY
            LEFT JOIN [BIDM].[dim].[DIM_STORE_GROUP_NEW] S ON C.store_hyperion_id = S.STORE_HYPERION_ID
            CROSS JOIN StartDates SD
        WHERE
            ACTIVITY_DATE BETWEEN SD.START_DATE AND GETDATE()
            AND D.MARKET_NAME IN ('Southern CA', 'Northern CA & NV', 'WA & AZ', 'CO & North TX', 'South TX', 'Midwest & Northeast', 'Southeast', 'North-Central FL', 'South FL')
        GROUP BY
            YEAR(ACTIVITY_DATE), MONTH(ACTIVITY_DATE),
            R.REGION_NAME,
            CASE WHEN D.MARKET_NAME = 'AUTONATION DIRECT FLORIDA' THEN 'CAAD' ELSE D.MARKET_NAME END,
            C.STORE_HYPERION_ID,
            C.STORE_NAME,
            MANUFACTURER_NAME,
            BPMS_ID,
            BPMS_MANAGER_NAME,
            BPMS_ASSOCIATE1_NAME,
            BPMS_ASSOCIATE2_NAME,
            CASE 
                WHEN [WBYC_REPORTINGGROUP] IN ('TRADE OFFER (VOI)','RANGE TRADE (VOI)') THEN 'NON-WBYC LEAD'
                WHEN [WBYC_REPORTINGGROUP] IN ('SELL ONLY','RANGE - SELL ONLY') AND PROVIDER_NAME <> 'WBYC-SHOWROOM' THEN 'AN.COM SELL ONLY LEAD'
                WHEN [WBYC_REPORTINGGROUP] IN ('UNKNOWN') THEN 'NON-WBYC LEAD'
                ELSE [WBYC_REPORTINGGROUP]
            END,
            CASE 
                WHEN [WBYC_REPORTINGGROUP] IN ('SELL ONLY','RANGE - SELL ONLY') AND SUBPROVIDER_NAME IN ('WBYC Sell Only') THEN 'WBYC-AN.com'
                WHEN [WBYC_REPORTINGGROUP] IN ('SELL ONLY','RANGE - SELL ONLY') AND SUBPROVIDER_NAME IN ('Dealer Website – WBYC') THEN 'WBYC-OEM Sites'
                ELSE NULL 
            END,
            S.GROUP4
    ),
    TEMP_ASSOCIATE AS (
        SELECT
            BPMS_ID,
            BPMS_ASSOCIATE_1,
            BPMS_ASSOCIATE_2,
            SUM(PURCHASES) AS PURCHASES
        FROM WBYC_DATA1
        WHERE BPMS_ASSOCIATE_2 IS NOT NULL
        GROUP BY BPMS_ID, BPMS_ASSOCIATE_1, BPMS_ASSOCIATE_2
    )

    -- ✅ FINAL SELECT — returnable via pandas.read_sql()
    SELECT 
        Year, 
        Month, 
        Region_Name, 
        BPMS_ASSOCIATE, 
        SUM(Purchases) AS Purchases
    FROM (
        SELECT 
            [YEAR], [MONTH], REGION_NAME,
            BPMS_ASSOCIATE_1 AS BPMS_ASSOCIATE,
            SUM(PURCHASES) AS PURCHASES
        FROM WBYC_DATA1
        WHERE BPMS_ASSOCIATE_2 IS NULL
        GROUP BY [YEAR], [MONTH], REGION_NAME, BPMS_ASSOCIATE_1

        UNION ALL

        SELECT 
            W.[YEAR], W.[MONTH], W.REGION_NAME,
            A.BPMS_ASSOCIATE_2 AS BPMS_ASSOCIATE,
            SUM(W.PURCHASES * 0.5) AS PURCHASES
        FROM WBYC_DATA1 W
        INNER JOIN TEMP_ASSOCIATE A ON W.BPMS_ID = A.BPMS_ID AND W.BPMS_ASSOCIATE_2 = A.BPMS_ASSOCIATE_2
        GROUP BY W.[YEAR], W.[MONTH], W.REGION_NAME, A.BPMS_ASSOCIATE_2

        UNION ALL

        SELECT 
            W.[YEAR], W.[MONTH], W.REGION_NAME,
            A.BPMS_ASSOCIATE_1 AS BPMS_ASSOCIATE,
            SUM(W.PURCHASES * 0.5) AS PURCHASES
        FROM WBYC_DATA1 W
        INNER JOIN TEMP_ASSOCIATE A 
            ON W.BPMS_ID = A.BPMS_ID 
            AND A.BPMS_ASSOCIATE_1 = W.BPMS_ASSOCIATE_1 
            AND A.BPMS_ASSOCIATE_2 = W.BPMS_ASSOCIATE_2
        GROUP BY W.[YEAR], W.[MONTH], W.REGION_NAME, A.BPMS_ASSOCIATE_1
    ) Z
    GROUP BY Year, Month, Region_Name, BPMS_ASSOCIATE
    '''

    Mkt_Data_query = """
    SELECT  [MONTH]
        ,[MARKET_NAME]
        ,[STORE]
        ,[DEPARTMENT_NAME] as DEPARTMENT
        ,SUM([OPPORTUNITIES]) AS [OPPORTUNITIES]
        ,SUM([INTERNET_OPPORTUNITIES]) AS [INTERNET_OPPORTUNITIES]
        ,SUM([SALES]) AS SALES
        ,SUM([COMPASS_SALES]) AS [COMPASS_SALES]
        ,SUM([CONTACT_MADE_COUNT]) AS [CONTACT_MADE_COUNT]
        ,SUM([INTERNET_CONTACT_MADE_COUNT]) AS [INTERNET_CONTACT_MADE_COUNT]
        ,SUM([TIME_TO_CONTACT_IN_MINS]) AS [TIME_TO_CONTACT_IN_MINS]
        ,SUM([LEADS_CONTACTED_WITHIN_24_HOURS_COUNT]) AS [LEADS_CONTACTED_WITHIN_24_HOURS_COUNT]
        ,SUM([FIRST_APPT_COUNT]) AS [FIRST_APPT_COUNT]
        ,SUM([FIRST_VISIT_COUNT]) AS [FIRST_VISIT_COUNT]
        ,SUM([FIRST_SHOW_COUNT]) AS [FIRST_SHOW_COUNT]
        ,SUM([FIRST_MENU_COUNT]) AS [FIRST_MENU_COUNT]

    FROM [BITESTDB].[FUENTESA1].[SALES_TRAFFIC_FUNNEL_FORECAST_VS_ACTUALS_T3_TREND] (nolock)

    WHERE DEPARTMENT_NAME = 'USED'

    GROUP BY [MONTH]
        ,[MARKET_NAME]
        ,[STORE]
        ,[DEPARTMENT_NAME]

    """

    Acq_data_query = f"""
    SELECT 
    MONTH([ACQUISITION_DATE]) as Month
        ,CASE WHEN [ACQUISITION_SOURCE_GROUPED] in ('Rental company','auction','Other CAAD','Street Buy') THEN 'Auction/Rental'
    WHEN [ACQUISITION_SOURCE_GROUPED] in ('LEASESTRPURCH','LEASEBUYOUT') THEN 'Lease Return'
    WHEN [ACQUISITION_SOURCE_GROUPED] in ('TRADE-IN-NEW','TRADE-IN-USED') THEN 'Trade-In'
    WHEN [ACQUISITION_SOURCE_GROUPED] in ('Loaner') THEN 'Service Loaner'
    WHEN [ACQUISITION_SOURCE_GROUPED] in ('WBYC') THEN 'WBYC'
    WHEN [ACQUISITION_SOURCE_GROUPED] in ('Other','Intercompany Trade') THEN 'Other'
    ELSE 'NONE' END AS Consolidated_Group
    ,VIN
        ,SUM([ACQUISITION_COUNT]) as Acquisitions
        ,SUM([VEHICLESOLDCOUNT]) as Units_Sold

    FROM [BIDM].[fact].[FACT_USED_ACQUISITIONS_PERFORMANCE]

    where [ACQUISITION_DATE] between '{three_months_prior}' and '{current_day}'

    group by 
    MONTH([ACQUISITION_DATE])
        ,CASE WHEN [ACQUISITION_SOURCE_GROUPED] in ('Rental company','auction','Other CAAD','Street Buy') THEN 'Auction/Rental'
    WHEN [ACQUISITION_SOURCE_GROUPED] in ('LEASESTRPURCH','LEASEBUYOUT') THEN 'Lease Return'
    WHEN [ACQUISITION_SOURCE_GROUPED] in ('TRADE-IN-NEW','TRADE-IN-USED') THEN 'Trade-In'
    WHEN [ACQUISITION_SOURCE_GROUPED] in ('Loaner') THEN 'Service Loaner'
    WHEN [ACQUISITION_SOURCE_GROUPED] in ('WBYC') THEN 'WBYC'
    WHEN [ACQUISITION_SOURCE_GROUPED] in ('Other','Intercompany Trade') THEN 'Other'
    ELSE 'NONE' END
    ,VIN

    Order by MONTH([ACQUISITION_DATE])
    """

    WBYC_query = '''
    DECLARE 
        @START_DATE AS DATE
    SET @START_DATE = DATEADD(MONTH, -5, DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)) -- LAST 12 MONTHS

    ;WITH WBYC_ACQ AS (
        SELECT 
            YEAR(TRX_DATE) AS [YEAR], 
            MONTH(TRX_DATE) AS [MONTH],
            A.VIN AS VIN,
            R.REGION_NAME,    
            CASE WHEN D.MARKET_NAME = 'AUTONATION DIRECT FLORIDA' THEN 'CAAD' ELSE D.MARKET_NAME END AS MARKET_NAME,
            A.STORE_HYPERION_ID,
            A.STORE_HYPERION_ID + '-' + S.STORE_NAME AS STORE,
            CAST(NULL AS VARCHAR) AS [WBYC_REPORTINGGROUP],
            CAST(NULL AS VARCHAR) AS [WBYC_REPORTINGGROUP_AN],
            M.MANUFACTURER_NAME,
            CAST(NULL AS VARCHAR) AS BPMS_MANAGER,
            CAST(NULL AS VARCHAR) AS BPMS_ASSOCIATE_1,
            CAST(NULL AS VARCHAR) AS BPMS_ASSOCIATE_2,
            CAST(NULL AS VARCHAR) AS KBB_STATUS,
            CASE WHEN SSK.SAME_STORE_FLAG = 'SAME STORE 2024' THEN 'Same Store' ELSE 'Not Same Store' END AS SAME_STORE_FLAG,
            0 AS LEAD_COUNT,  
            0 AS SALES_COUNT,
            0 AS SALES,
            0 AS TRADE_IN,
            0 AS ACQUISITIONS,
            0 AS PURCHASE_BY_ASSOCIATE,
            0 AS ACQ_SALES,
            0 AS ACQ_SALES_2,    
            0 AS FRONT_GROSS,
            0 AS TOTAL_GROSS,
            0 AS WHOLESALE_COUNT,
            0 AS WHOLESALE_GROSS,
            0 AS VEHICLE_AGE,
            SUM(A.[ACQUISITION_COUNT]) AS ACQUISITION_COUNT,
            SUM(CASE WHEN [ACQUISITION_SOURCE_RPTGRP1] = 'WBYC' THEN A.[ACQUISITION_COUNT] ELSE 0 END) AS WBYC_ACQUISITION,
            SUM(A.VEHICLESOLDCOUNT) AS USED_SALES,
            SUM(CASE WHEN [ACQUISITION_SOURCE_RPTGRP1] = 'WBYC' THEN A.VEHICLESOLDCOUNT ELSE 0 END) AS WBYC_SALES,
            0 AS ACQ_DWN_STREAM_3RDPARTY_COUNT,
            0 AS ACQ_DWN_STREAM_3RDPARTY_COUNT_2,
            0 AS ACQ_DWN_STREAM_3RDPARTY_GROSS
        FROM [BIDM].[FACT].[FACT_USED_ACQUISITIONS_PERFORMANCE] A (NOLOCK)
        LEFT JOIN BIDM.DIM.DIM_STORE S
            ON S.STORE_HYPERION_ID = A.STORE_HYPERION_ID
        LEFT JOIN BIDM.DIM.DIM_MARKET D
            ON D.MARKET_EID = S.MARKET_EID
        LEFT JOIN BIDM.DIM.DIM_REGION R
            ON R.REGION_EID = S.REGION_EID
        LEFT JOIN BIDM.DIM.DIM_MANUFACTURER M
            ON M.MANUFACTURER_KEY = S.MANUFACTURER_KEY
        LEFT JOIN BIDM.DIM.DIM_SAMESTORE SS WITH (READUNCOMMITTED)
            ON S.STORE_EID = SS.STORE_EID
            AND MONTH(TRX_DATE) = SS.CALENDAR_MONTH
        LEFT JOIN (SELECT * FROM BIDM.DIM.DIM_SSS_FLAG WHERE SAME_STORE_FLAG = 'SAME STORE 2024') SSK
            ON SS.SSS_KEY = SSK.SSS_KEY
        WHERE TRX_DATE BETWEEN @START_DATE AND GETDATE()
        GROUP BY
            YEAR(TRX_DATE),
            MONTH(TRX_DATE),
            A.VIN,
            R.REGION_NAME,    
            CASE WHEN D.MARKET_NAME = 'AUTONATION DIRECT FLORIDA' THEN 'CAAD' ELSE D.MARKET_NAME END,
            A.STORE_HYPERION_ID,
            A.STORE_HYPERION_ID + '-' + S.STORE_NAME,
            M.MANUFACTURER_NAME,
            CASE WHEN SSK.SAME_STORE_FLAG = 'SAME STORE 2024' THEN 'Same Store' ELSE 'Not Same Store' END
        HAVING SUM(A.[ACQUISITION_COUNT]) + SUM(A.VEHICLESOLDCOUNT) > 0
    ),
    TEMP_VOI_LEADS AS (
        SELECT
            LEAD_ID  
        FROM BIDM.[FACT].[FACT_WBYC_TRAFFIC_DETAIL] A
        WHERE [WBYC_REPORTINGGROUP] IN ('RANGE TRADE (VOI)', 'TRADE OFFER (VOI)')
            AND PURCHASE_COUNT >= 1
        GROUP BY LEAD_ID
    ),
    WBYC_DATA AS (
        SELECT
            YEAR(ACTIVITY_DATE) AS [YEAR],
            MONTH(ACTIVITY_DATE) AS [MONTH],
            A.ACQ_VIN AS VIN,
            R.REGION_NAME,    
            CASE WHEN D.MARKET_NAME = 'AUTONATION DIRECT FLORIDA' THEN 'CAAD' ELSE D.MARKET_NAME END AS MARKET_NAME,
            C.STORE_HYPERION_ID,
            C.STORE_HYPERION_ID + '-' + C.STORE_NAME AS STORE,
            CASE 
                WHEN [WBYC_REPORTINGGROUP] IN ('TRADE OFFER (VOI)', 'RANGE TRADE (VOI)') THEN 'NON-WBYC LEAD'
                WHEN [WBYC_REPORTINGGROUP] IN ('SELL ONLY', 'RANGE - SELL ONLY') AND PROVIDER_NAME <> 'WBYC-SHOWROOM' THEN 'AN.COM SELL ONLY LEAD'
                WHEN SP.SUBPROVIDER_NAME = 'Sell my car' AND P.PROVIDER_NAME = 'Cargurus' THEN 'CARGURUS-Sell My Car'
                WHEN [WBYC_REPORTINGGROUP] IN ('UNKNOWN') THEN 'NON-WBYC LEAD'
                ELSE [WBYC_REPORTINGGROUP] 
            END AS [WBYC_REPORTINGGROUP],
            CASE 
                WHEN [WBYC_REPORTINGGROUP] IN ('SELL ONLY', 'RANGE - SELL ONLY') AND SUBPROVIDER_NAME IN ('WBYC Sell Only') THEN 'WBYC-AN.com'
                WHEN [WBYC_REPORTINGGROUP] IN ('SELL ONLY', 'RANGE - SELL ONLY') AND SUBPROVIDER_NAME IN ('Dealer Website – WBYC', 'Dealer Website – WBYC') THEN 'WBYC-OEM Sites'
                ELSE NULL 
            END AS [WBYC_REPORTINGGROUP_AN],
            M.MANUFACTURER_NAME,
            CAST(BPMS_MANAGER_NAME AS VARCHAR) AS BPMS_MANAGER,
            CAST(BPMS_ASSOCIATE1_NAME AS VARCHAR) AS BPMS_ASSOCIATE_1,
            CAST(BPMS_ASSOCIATE2_NAME AS VARCHAR) AS BPMS_ASSOCIATE_2,
            K.[STATUS] AS KBB_STATUS,
            CASE WHEN SSK.SAME_STORE_FLAG = 'SAME STORE 2024' THEN 'Same Store' ELSE 'Not Same Store' END AS SAME_STORE_FLAG,
            CAST(SUM(CASE WHEN [WBYC_REPORTINGGROUP] IN ('RANGE TRADE (VOI)', 'TRADE OFFER (VOI)') AND VL.LEAD_ID IS NULL THEN 0 ELSE A.[OPPORTUNITIES] END) AS INT) AS LEAD_COUNT,
            CAST(SUM(CASE WHEN [WBYC_REPORTINGGROUP] IN ('RANGE TRADE (VOI)', 'TRADE OFFER (VOI)') AND VL.LEAD_ID IS NULL THEN 0 ELSE A.[SALES] END) AS INT) AS SALES_COUNT,
            CAST(SUM(CASE WHEN [WBYC_REPORTINGGROUP] IN ('RANGE TRADE (VOI)', 'TRADE OFFER (VOI)') AND VL.LEAD_ID IS NULL THEN 0 ELSE A.[COMPASS_SALES] END) AS INT) AS COMPASS_SALES,
            CAST(SUM(CASE WHEN [WBYC_REPORTINGGROUP] IN ('RANGE TRADE (VOI)', 'TRADE OFFER (VOI)') AND VL.LEAD_ID IS NULL THEN 0 ELSE A.[TRADEIN_COUNT] END) AS INT) AS TRADE_IN,
            SUM([PURCHASE_COUNT]) AS ACQUISITIONS,
            SUM(CASE WHEN BPMS_ASSOCIATE2_NAME IS NOT NULL THEN ([PURCHASE_COUNT] * 0.5) ELSE [PURCHASE_COUNT] END) AS PURCHASE_BY_ASSOCIATE,
            SUM([ACQ_DWN_STREAM_SALES]) AS ACQ_SALES,
            CAST(SUM(CASE WHEN [WBYC_REPORTINGGROUP] IN ('RANGE TRADE (VOI)', 'TRADE OFFER (VOI)', 'UNKNOWN') AND VL.LEAD_ID IS NULL THEN 0 ELSE A.[ACQ_DWN_STREAM_SALES] END) AS INT) AS ACQ_SALES_2,
            SUM([ACQ_DWN_STREAM_FRONT_GROSS]) AS FRONT_GROSS,
            SUM([ACQ_DWN_STREAM_TOTAL_GROSS]) AS TOTAL_GROSS,
            SUM([ACQ_DWN_STREAM_3RDPARTY_COUNT]) AS WHOLESALE_COUNT,
            SUM([ACQ_DWN_STREAM_3RDPARTY_GROSS]) AS WHOLESALE_GROSS,
            SUM([ACQ_DWN_STREAM_VEHICLE_ANAGE]) AS VEHICLE_AGE,
            0 AS ACQUISITION_COUNT,
            0 AS WBYC_ACQUISITION,
            0 AS USED_SALES,
            0 AS WBYC_SALES,
            SUM(ACQ_DWN_STREAM_3RDPARTY_COUNT) AS ACQ_DWN_STREAM_3RDPARTY_COUNT,
            CAST(SUM(CASE WHEN [WBYC_REPORTINGGROUP] IN ('RANGE TRADE (VOI)', 'TRADE OFFER (VOI)', 'UNKNOWN') AND VL.LEAD_ID IS NULL THEN 0 ELSE A.ACQ_DWN_STREAM_3RDPARTY_COUNT END) AS INT) AS ACQ_DWN_STREAM_3RDPARTY_COUNT_2,
            SUM(ACQ_DWN_STREAM_3RDPARTY_GROSS) AS ACQ_DWN_STREAM_3RDPARTY_GROSS
        FROM BIDM.[FACT].[FACT_WBYC_TRAFFIC_DETAIL] A 
        LEFT JOIN BIDM.DIM.DIM_STORE C
            ON C.STORE_EID = A.STORE_EID
        LEFT JOIN BIDM.DIM.DIM_MARKET D
            ON D.MARKET_EID = C.MARKET_EID
        LEFT JOIN BIDM.DIM.DIM_PROVIDER P
            ON P.PROVIDER_KEY = A.PROVIDER_KEY
        LEFT JOIN [BIDM].[dim].[DIM_SUBPROVIDER] SP
            ON SP.SUBPROVIDER_PKEY = A.SUBPROVIDER_PKEY
        LEFT JOIN BIDM.DIM.DIM_REGION R
            ON R.REGION_EID = C.REGION_EID
        LEFT JOIN BIDM.DIM.DIM_MANUFACTURER M
            ON M.MANUFACTURER_KEY = C.MANUFACTURER_KEY
        LEFT JOIN TEMP_VOI_LEADS VL
            ON A.LEAD_ID = VL.LEAD_ID 
        LEFT JOIN BITESTDB.FUENTESA1.KBB_ACTIVATION_NEW K
            ON K.Hype = C.STORE_HYPERION_ID 
        LEFT JOIN BIDM.DIM.DIM_SAMESTORE SS WITH (READUNCOMMITTED)
            ON A.STORE_EID = SS.STORE_EID
            AND MONTH(A.ACTIVITY_DATE) = SS.CALENDAR_MONTH
        LEFT JOIN (SELECT * FROM BIDM.DIM.DIM_SSS_FLAG WHERE SAME_STORE_FLAG = 'SAME STORE 2024') SSK
            ON SS.SSS_KEY = SSK.SSS_KEY
        WHERE ACTIVITY_DATE BETWEEN @START_DATE AND GETDATE()
        GROUP BY
            YEAR(ACTIVITY_DATE),
            MONTH(ACTIVITY_DATE),
            A.ACQ_VIN,
            R.REGION_NAME,
            D.MARKET_NAME,
            C.STORE_HYPERION_ID,
            C.STORE_NAME,
            MANUFACTURER_NAME,
            BPMS_MANAGER_NAME,
            BPMS_ASSOCIATE1_NAME,
            BPMS_ASSOCIATE2_NAME,
            K.[STATUS],
            CASE 
                WHEN [WBYC_REPORTINGGROUP] IN ('TRADE OFFER (VOI)', 'RANGE TRADE (VOI)') THEN 'NON-WBYC LEAD'
                WHEN [WBYC_REPORTINGGROUP] IN ('SELL ONLY', 'RANGE - SELL ONLY') AND PROVIDER_NAME <> 'WBYC-SHOWROOM' THEN 'AN.COM SELL ONLY LEAD'
                WHEN SP.SUBPROVIDER_NAME = 'Sell my car' AND P.PROVIDER_NAME = 'Cargurus' THEN 'CARGURUS-Sell My Car'
                WHEN [WBYC_REPORTINGGROUP] IN ('UNKNOWN') THEN 'NON-WBYC LEAD'
                ELSE [WBYC_REPORTINGGROUP] 
            END,
            CASE 
                WHEN [WBYC_REPORTINGGROUP] IN ('SELL ONLY', 'RANGE - SELL ONLY') AND SUBPROVIDER_NAME IN ('WBYC Sell Only') THEN 'WBYC-AN.com'
                WHEN [WBYC_REPORTINGGROUP] IN ('SELL ONLY', 'RANGE - SELL ONLY') AND SUBPROVIDER_NAME IN ('Dealer Website – WBYC', 'Dealer Website – WBYC') THEN 'WBYC-OEM Sites'
                ELSE NULL 
            END,
            CASE WHEN D.MARKET_NAME = 'AUTONATION DIRECT FLORIDA' THEN 'CAAD' ELSE D.MARKET_NAME END,
            CASE WHEN SSK.SAME_STORE_FLAG = 'SAME STORE 2024' THEN 'Same Store' ELSE 'Not Same Store' END
    )
    SELECT 
        Year, 
        Month, 
        VIN, 
        WBYC_REPORTINGGROUP, 
        LEAD_COUNT, 
        ACQUISITIONS, 
        ACQUISITION_COUNT,
        CASE 
            WHEN WBYC_REPORTINGGROUP IN ('AN.COM SELL ONLY LEAD', 'KBB INSTANT CASH OFFER') THEN WBYC_REPORTINGGROUP 
            ELSE 'WBYC-OTHER' 
        END AS WBYC_REPORTING_GROUP
    FROM WBYC_DATA w
    LEFT JOIN [BIDM].[dim].[DIM_STORE_GROUP_NEW] s
        ON w.STORE_HYPERION_ID = s.STORE_HYPERION_ID
    WHERE MARKET_NAME IN ('Southern CA', 'Northern CA & NV', 'WA & AZ', 'CO & North TX', 'South TX', 'Midwest & Northeast', 'Southeast', 'North-Central FL', 'South FL')
        AND (LEAD_COUNT <> 0 OR ACQUISITIONS <> 0 OR ACQUISITION_COUNT <> 0)

    UNION ALL

    SELECT 
        Year, 
        Month, 
        VIN, 
        WBYC_REPORTINGGROUP, 
        LEAD_COUNT, 
        ACQUISITIONS, 
        ACQUISITION_COUNT,
        CASE 
            WHEN WBYC_REPORTINGGROUP IN ('AN.COM SELL ONLY LEAD', 'KBB INSTANT CASH OFFER') THEN WBYC_REPORTINGGROUP 
            ELSE 'WBYC-OTHER' 
        END AS WBYC_REPORTING_GROUP
    FROM WBYC_ACQ w
    LEFT JOIN [BIDM].[dim].[DIM_STORE_GROUP_NEW] s
        ON w.STORE_HYPERION_ID = s.STORE_HYPERION_ID
    WHERE MARKET_NAME IN ('Southern CA', 'Northern CA & NV', 'WA & AZ', 'CO & North TX', 'South TX', 'Midwest & Northeast', 'Southeast', 'North-Central FL', 'South FL')
        AND (LEAD_COUNT <> 0 OR ACQUISITIONS <> 0 OR ACQUISITION_COUNT <> 0)
        '''

    Spend_Query = f'''
    WITH WBYC AS (
        SELECT 
            [STORE_EID],
            [ACTIVITY_DATE],
            [WBYC_REPORTINGGROUP],
            [CUST_ZIP_CODE],
            SUM([OPPORTUNITIES]) AS [OPPORTUNITIES],
            SUM([PURCHASE_COUNT]) AS [PURCHASE_COUNT]
        FROM [BIDM].[FACT].[FACT_WBYC_TRAFFIC_DETAIL] W
        WHERE ACTIVITY_DATE BETWEEN '{three_months_prior}' AND GETDATE() - 1
        GROUP BY 
            [STORE_EID],
            [ACTIVITY_DATE],
            [WBYC_REPORTINGGROUP],
            [CUST_ZIP_CODE]
    ),
    LATLONG AS (
        SELECT
            A.*,
            DS.ZIP AS STORE_ZIPCODE,
            ZP.LATITUDE AS CUSTLATITUDE,
            ZP.LONGITUDE AS CUSTLONGITUDE,
            ZP1.LATITUDE AS STORELATITUDE,
            ZP1.LONGITUDE AS STORELONGITUDE
        FROM WBYC A
        LEFT JOIN BIDM.DIM.DIM_STORE DS WITH (READUNCOMMITTED)
            ON A.STORE_EID = DS.STORE_EID
        LEFT JOIN BIDM.[DIM].[DIM_ZIPCODE] ZP1
            ON LEFT(DS.ZIP, 5) = ZP1.ZIP_CODE
        LEFT JOIN BIDM.[DIM].[DIM_ZIPCODE] ZP
            ON LEFT(A.CUST_ZIP_CODE, 5) = ZP.ZIP_CODE
    ),
    SPEND AS (
        SELECT 
            MNTH,
            YR,
            MARKET_NAME,
            WBYC_REPORTINGGROUP,
            SUM(ADV_SPEND) AS ADV_SPEND
        FROM BITESTDB.MELBOURNED.TEMP_WBYC_POD_PERFORMANCE_BY_VIN
        WHERE MARKET_NAME IN ('SOUTHERN CA', 'NORTHERN CA & NV', 'WA & AZ', 'CO & NORTH TX', 'SOUTH TX', 'MIDWEST & NORTHEAST', 'SOUTHEAST', 'NORTH-CENTRAL FL', 'SOUTH FL')
            AND YR IN ('2024', '2025')
        GROUP BY
            MNTH,
            YR,
            MARKET_NAME,
            WBYC_REPORTINGGROUP
        HAVING SUM(ADV_SPEND) > 0
    ),
    DISTANCE_CALC AS (
        SELECT
            A.*,
            ISNULL(CONVERT(NUMERIC(8,2), SQRT(SQUARE((A.CUSTLATITUDE - A.STORELATITUDE) * 69.1) + SQUARE((A.CUSTLONGITUDE - A.STORELONGITUDE) * 53))), -1) AS DISTANCE
        FROM LATLONG A
    )
    SELECT
        YEAR(ACTIVITY_DATE) AS YR,
        MONTH(ACTIVITY_DATE) AS MNTH,
        MARKET_NAME,
        [WBYC_REPORTINGGROUP],
        CASE 
            WHEN DISTANCE BETWEEN 0 AND 4.99 THEN '1: 0 - 4.99 MILES'
            WHEN DISTANCE BETWEEN 4.99 AND 9.99 THEN '2: 5 - 9.99 MILES'
            WHEN DISTANCE BETWEEN 9.99 AND 14.99 THEN '3: 10 - 14.99 MILES'
            WHEN DISTANCE BETWEEN 14.99 AND 19.99 THEN '4: 15 - 19.99 MILES'
            WHEN DISTANCE BETWEEN 19.99 AND 29.99 THEN '5: 20 - 29.99 MILES'
            WHEN DISTANCE BETWEEN 29.99 AND 39.99 THEN '6: 30 - 39.99 MILES'
            WHEN DISTANCE BETWEEN 39.99 AND 49.99 THEN '7: 40 - 49.99 MILES'
            WHEN DISTANCE BETWEEN 49.99 AND 74.99 THEN '8: 50 - 75 MILES'
            WHEN DISTANCE BETWEEN 75.00 AND 99.99 THEN '9: 75 - 100 MILES'  
            WHEN DISTANCE BETWEEN 100.00 AND 149.99 THEN '10: 100 - 150 MILES'
            WHEN DISTANCE > 149.99 THEN '11: 150+ MILES' 
            ELSE '12: UNKNOWN' 
        END AS DISTANCE_BAND,
        SUM(OPPORTUNITIES) AS OPPORTUNITIES,
        SUM([PURCHASE_COUNT]) AS PURCHASE_COUNT,
        SUM(0) AS SPEND
    FROM DISTANCE_CALC AA
    LEFT JOIN BIDM.DIM.DIM_STORE DS
        ON AA.STORE_EID = DS.STORE_EID
    LEFT JOIN BIDM.DIM.DIM_MARKET M
        ON M.MARKET_EID = DS.MARKET_EID
    GROUP BY
        YEAR(ACTIVITY_DATE),
        MONTH(ACTIVITY_DATE),
        MARKET_NAME,
        [WBYC_REPORTINGGROUP],
        CASE 
            WHEN DISTANCE BETWEEN 0 AND 4.99 THEN '1: 0 - 4.99 MILES'
            WHEN DISTANCE BETWEEN 4.99 AND 9.99 THEN '2: 5 - 9.99 MILES'
            WHEN DISTANCE BETWEEN 9.99 AND 14.99 THEN '3: 10 - 14.99 MILES'
            WHEN DISTANCE BETWEEN 14.99 AND 19.99 THEN '4: 15 - 19.99 MILES'
            WHEN DISTANCE BETWEEN 19.99 AND 29.99 THEN '5: 20 - 29.99 MILES'
            WHEN DISTANCE BETWEEN 29.99 AND 39.99 THEN '6: 30 - 39.99 MILES'
            WHEN DISTANCE BETWEEN 39.99 AND 49.99 THEN '7: 40 - 49.99 MILES'
            WHEN DISTANCE BETWEEN 49.99 AND 74.99 THEN '8: 50 - 75 MILES'
            WHEN DISTANCE BETWEEN 75.00 AND 99.99 THEN '9: 75 - 100 MILES'  
            WHEN DISTANCE BETWEEN 100.00 AND 149.99 THEN '10: 100 - 150 MILES'
            WHEN DISTANCE > 149.99 THEN '11: 150+ MILES' 
            ELSE '12: UNKNOWN' 
        END
    HAVING SUM(OPPORTUNITIES) <> 0 OR SUM([PURCHASE_COUNT]) <> 0
    UNION
    SELECT
        YR,
        MNTH,
        MARKET_NAME,
        [WBYC_REPORTINGGROUP],
        NULL AS DISTANCE_BAND,
        SUM(0) AS OPPORTUNITIES,
        SUM(0) AS PURCHASE_COUNT,
        SUM(ADV_SPEND) AS SPEND
    FROM SPEND
    GROUP BY
        YR,
        MNTH,
        MARKET_NAME,
        [WBYC_REPORTINGGROUP]
    HAVING SUM(ADV_SPEND) <> 0
    '''

    # right click on connection and go to properties to find the server name then select the database its looking at
    try:
        with pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=S1WPVSQLMBI1,46160;'
                'DATABASE=BIDM;'
                'Trusted_Connection=yes;'
        ) as conn:
                      
            AssociateData_df = pd.read_sql(AssociateData_Query, conn)        
            # Print time it took to load query
            end_time = time.time()
            elapsed_time = end_time - start_time
            print(f"Script read in AssociateData in {elapsed_time:.2f} seconds")
            # Print time it took to load query
            Mkt_Data_df = pd.read_sql(Mkt_Data_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            print(f"Script read in Mkt_Data in {elapsed_time:.2f} seconds")
            # Print time it took to load query
            Acq_Data_df = pd.read_sql(Acq_data_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in Acq_Data in {elapsed_time:.2f} seconds")
            WBYC_df = pd.read_sql(WBYC_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in WBYC_data in {elapsed_time:.2f} seconds")
            Spend_df = pd.read_sql(Spend_Query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in Spend_data in {elapsed_time:.2f} seconds")

    except Exception as e:
        print("❌ Connection failed:", e)

################################################################################################################
    '''RUN SPEED TO MARKET QUERIES'''
################################################################################################################

    Speed_to_Market_query = f"""
    select MONTH(A.Snapshot_Date) as MTH,
        A.Hyperion,
        count(distinct A.vin) as VinCount,
        case when datediff(day, A.Entry_Date, A.Fifteen_Photo_Date) <=5 then '0-5 Days'
                when datediff(day, A.Entry_Date, A.Fifteen_Photo_Date) > 5 and datediff(day, A.Entry_Date, A.Fifteen_Photo_Date) <= 10 then '6-10 Days'
                when datediff(day, A.Entry_Date, A.Fifteen_Photo_Date) > 10 then 'Above 10 Days'
        else 'NA' end as Flag
    from (
    select i1.Snapshot_Date, i1.Hyperion,i1.VIN, i1.Entry_Date,i1.Fifteen_Photo_Date
    from [NDDUsers].[vSpeed_To_Market_Trend_Daily_Snapshot] i1 (nolock)
    join [NDDUsers].[vInventory_Daily_Snapshot] ds(nolock) on ds.vin = i1.vin
    where --i1.Hyperion in ('2281','2140') and 
    i1.Snapshot_Date between '{current_day}' and '{current_day}'
    and ds.StockType = 'USED' and ds.ValidForAN = 1 and ds.Status = 'S'
    and ds.ImageType <> 'StockPhotos'
    and i1.Fifteen_Photo_Date is not null
    and i1.Fifteen_Photo_Date >= i1.Entry_Date
    and datediff(day, i1.Entry_Date, i1.Fifteen_Photo_Date) <= 30 
    group by i1.Snapshot_Date, i1.Hyperion,i1.VIN, i1.Entry_Date,i1.Fifteen_Photo_Date

    ) A
    group by MONTH(A.Snapshot_Date),
            A.Hyperion,
            case when datediff(day, A.Entry_Date, A.Fifteen_Photo_Date) <=5 then '0-5 Days'
                when datediff(day, A.Entry_Date, A.Fifteen_Photo_Date) > 5 and datediff(day, A.Entry_Date, A.Fifteen_Photo_Date) <= 10 then '6-10 Days'
                when datediff(day, A.Entry_Date, A.Fifteen_Photo_Date) > 10 then 'Above 10 Days'
            else 'NA' end
    """

    Speed_to_Market_query2 = f'''
    WITH FinalCTE AS (
        SELECT
            DATEADD(month, DATEDIFF(month, 0, i1.Snapshot_Date), 0) AS SnapshotDate,
            i1.Hyperion,

            -- Between 0-5 Days
            SUM(CASE 
                WHEN snapshot_date <= '{beginning_of_month}' THEN NULL 
                WHEN i1.Fifteen_Photo_Date IS NOT NULL 
                    AND i1.Entry_Date <= i1.Fifteen_Photo_Date 
                    AND DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) <= 5 
                THEN DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) 
                ELSE NULL 
            END) AS TotalDaysFifteenPhoto_under5,

            SUM(CASE 
                WHEN snapshot_date <= '{beginning_of_month}' THEN NULL 
                WHEN i1.Fifteen_Photo_Date IS NOT NULL 
                    AND i1.Entry_Date <= i1.Fifteen_Photo_Date 
                    AND DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) <= 5 
                THEN 1 ELSE 0 
            END) AS TotalRecordsFifteenPhoto_under5,

            -- Between 6-10 Days
            SUM(CASE 
                WHEN snapshot_date <= '{beginning_of_month}' THEN NULL 
                WHEN i1.Fifteen_Photo_Date IS NOT NULL 
                    AND i1.Entry_Date <= i1.Fifteen_Photo_Date 
                    AND DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) > 5 
                    AND DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) <= 10 
                THEN DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) 
                ELSE NULL 
            END) AS TotalDaysFifteenPhoto_under10,

            SUM(CASE 
                WHEN snapshot_date <= '{beginning_of_month}' THEN NULL 
                WHEN i1.Fifteen_Photo_Date IS NOT NULL 
                    AND i1.Entry_Date <= i1.Fifteen_Photo_Date 
                    AND DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) > 5 
                    AND DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) <= 10 
                THEN 1 ELSE 0 
            END) AS TotalRecordsFifteenPhoto_under10,

            -- Between 10-30 Days
            SUM(CASE 
                WHEN snapshot_date <= '{beginning_of_month}' THEN NULL 
                WHEN i1.Fifteen_Photo_Date IS NOT NULL 
                    AND i1.Entry_Date <= i1.Fifteen_Photo_Date 
                    AND DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) > 10 
                    AND DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) <= 30 
                THEN DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) 
                ELSE NULL 
            END) AS TotalDaysFifteenPhoto_Above10,

            SUM(CASE 
                WHEN snapshot_date <= '{beginning_of_month}' THEN NULL 
                WHEN i1.Fifteen_Photo_Date IS NOT NULL 
                    AND i1.Entry_Date <= i1.Fifteen_Photo_Date 
                    AND DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) > 10 
                    AND DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) <= 30 
                THEN 1 ELSE 0 
            END) AS TotalRecordsFifteenPhoto_Above10,

            -- Monthly Avg
            SUM(CASE 
                WHEN snapshot_date <= '{beginning_of_month}' THEN NULL 
                WHEN i1.Fifteen_Photo_Date IS NOT NULL 
                    AND i1.Entry_Date <= i1.Fifteen_Photo_Date 
                    AND DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) <= 30 
                THEN DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) 
                ELSE NULL 
            END) AS TotalDaysFifteenPhoto,

            SUM(CASE 
                WHEN snapshot_date <= '{beginning_of_month}' THEN NULL 
                WHEN i1.Fifteen_Photo_Date IS NOT NULL 
                    AND i1.Entry_Date <= i1.Fifteen_Photo_Date 
                    AND DATEDIFF(day, i1.Entry_Date, i1.Fifteen_Photo_Date) <= 30 
                THEN 1 ELSE 0 
            END) AS TotalRecordsFifteenPhoto

        FROM [NDDUsers].[vSpeed_To_Market_Trend_Daily_Snapshot] i1 (NOLOCK)
        JOIN NDD.vEntities_Hierarchy e (NOLOCK) ON i1.Hyperion = e.StoreHyperion
        WHERE i1.Snapshot_Date >= '{beginning_of_month}'
        AND i1.Snapshot_Date <= '{current_day}'
        GROUP BY DATEADD(month, DATEDIFF(month, 0, i1.Snapshot_Date), 0), i1.Hyperion
    )

    SELECT
        SnapshotDate,
        Hyperion,
        CAST(ISNULL(TotalDaysFifteenPhoto_under5, 0) AS FLOAT) AS AvgDaysToFifteenPhotos_under5_Num,
        CAST(ISNULL(TotalRecordsFifteenPhoto_under5, 0) AS FLOAT) AS AvgDaysToFifteenPhotos_under5_Den,
        CAST(ISNULL(TotalDaysFifteenPhoto_under10, 0) AS FLOAT) AS AvgDaysToFifteenPhotos_under10_Num,
        CAST(ISNULL(TotalRecordsFifteenPhoto_under10, 0) AS FLOAT) AS AvgDaysToFifteenPhotos_under10_Den,
        CAST(ISNULL(TotalDaysFifteenPhoto_Above10, 0) AS FLOAT) AS AvgDaysToFifteenPhotos_Above10_Num,
        CAST(ISNULL(TotalRecordsFifteenPhoto_Above10, 0) AS FLOAT) AS AvgDaysToFifteenPhotos_Above10_Den,
        CAST(ISNULL(TotalDaysFifteenPhoto, 0) AS FLOAT) AS AvgDaysToFifteenPhotos_Num,
        CAST(ISNULL(TotalRecordsFifteenPhoto, 0) AS FLOAT) AS AvgDaysToFifteenPhotos_Den
    FROM FinalCTE
    ORDER BY SnapshotDate ASC;
    '''

    try:
        with pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=NDDPRDDB03.us1.autonation.com, 48155;'
                'DATABASE=NDD_DW;'
                'Trusted_Connection=yes;'
        ) as conn:
            Speed_to_Market_df = pd.read_sql(Speed_to_Market_query, conn)
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in Speed_to_Market_df in {elapsed_time:.2f} seconds")
            Speed_to_Market_df2 = pd.read_sql(Speed_to_Market_query2, conn)            
            end_time = time.time()
            elapsed_time = end_time - start_time
            # Print time it took to load query
            print(f"Script read in Speed_to_Market_df2 in {elapsed_time:.2f} seconds")
            
    except Exception as e:
        print("❌ Connection failed:", e)

    # WE NEED TO USE THE PRIOR MONTH IF WITHIN THE FIRST 7 DAYS OF THE MONTH OTHERWISE SEARCH IN CURRENT MONTH FOLDER
    # Check if we're in the first 7 days of the month
    if today.day <= 7:
        # Use the last day of the previous month to get the correct format
        first_of_this_month = today.replace(day=1)
        prior_month_date = first_of_this_month - timedelta(days=1)
        month_formatted = str(prior_month_date.month)
    else:
        # Use current month
        month_formatted = str(today.month)

    print(f"We are using the following month folder {month_formatted}")

    Used_Car_Path = fr'W:\Corporate\Inventory\Reporting\Used Car Program - Strategy Sheet\{month_formatted}' 

    '''do we need this below??? just use latest used car file...then save to proper path'''


    # # Get all .xlsm files in the Speed to Market folder that do NOT contain 'HC'
    # Used_Car_files = [
    #     f for f in glob.glob(os.path.join(Used_Car_Path, "*.xlsx"))
    #     if "HC" not in os.path.basename(f)
    # ]

    # # Get the latest file from the filtered list
    # if Used_Car_files:
    #     latest_Used_car_file = max(Used_Car_files, key=os.path.getmtime)
    #     print(f"Latest .xlsx file (excluding 'HC'): {latest_Used_car_file}")
    # else:
    #     print("No matching .xlsx files found (excluding 'HC').")

    # # Save to OneDrive to Process Static File (no date adjustments needed)
    # shutil.copyfile(latest_Used_car_file, r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Used_Car_Program\Used Car Program - Strategy Sheet.xlsx')

    # # Wait 5 seconds so that file properly saves?
    # time.sleep(5)

    # Open up latest file and dump sql queries in
    app = xw.App(visible=True) 
    wb = app.books.open(r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Used_Car_Program\Used Car Program - Strategy Sheet.xlsx')

    # Dump data into tab section to process'
    # Process all the NDD Data
    InvData_tab = wb.sheets['InvData']
    InvData_tab.range('A2:E10000').clear_contents()
    InvData_tab.range('A2').options(index=False, header=False).value = InvData_df

    SalesData_tab = wb.sheets['SalesData']
    SalesData_tab.range('A2:T10000').clear_contents()
    SalesData_tab.range('A2').options(index=False, header=False).value = SalesData_df

    New_and_Used_tab = wb.sheets['New and Used Sales']
    New_and_Used_tab.range('A2:D10000').clear_contents()
    New_and_Used_tab.range('A2').options(index=False, header=False).value = New_and_Used_Sales_df    

    Wholesale_tab = wb.sheets['Wholesale Data']
    Wholesale_tab.range('A2:C10000').clear_contents()
    Wholesale_tab.range('A2').options(index=False, header=False).value = Wholesale_Data_df

    TargetPrice_tab = wb.sheets['TargetPrice']
    TargetPrice_tab.range('A2:D500000').clear_contents()
    TargetPrice_tab.range('A2').options(index=False, header=False).value = TargetPrice_df

    # Process all the Marketing Data
    AssociateData_tab = wb.sheets['AssociateData']
    AssociateData_tab.range('A2:E10000').clear_contents()
    AssociateData_tab.range('A2').options(index=False, header=False).value = AssociateData_df

    Mkt_Data_tab = wb.sheets['Mkt_Data']
    Mkt_Data_tab.range('A2:P10000').clear_contents()
    Mkt_Data_tab.range('A2').options(index=False, header=False).value = Mkt_Data_df        
    
    Acq_Data_tab = wb.sheets['Acq_Data']
    Acq_Data_tab.range('A2:I10000').clear_contents()
    Acq_Data_tab.range('A2').options(index=False, header=False).value = Acq_Data_df      
    
    WBYC_tab = wb.sheets['WBYC']
    WBYC_tab.range('A2:H500000').clear_contents()
    WBYC_tab.range('A2').options(index=False, header=False).value = WBYC_df 

    Spend_tab = wb.sheets['Spend']
    Spend_tab.range('A2:H500000').clear_contents()
    Spend_tab.range('A2').options(index=False, header=False).value = Spend_df

    OKR_tab = wb.sheets['OKR Score Card']
    OKR_tab.range('D5').value = current_day

    '''do we need? just use the most recent speed to market...'''

    # Speed_to_Market_Path = fr'W:\Corporate\Inventory\Reporting\Used Speed to Market' 

    # # Get all .xlsm files in the Speed to Market folder that do NOT contain 'HC'
    # Speed_to_Market_files = [
    #     f for f in glob.glob(os.path.join(Speed_to_Market_Path, "*.xlsm"))
    #     if "HC" not in os.path.basename(f)
    # ]

    # # Get the latest file from the filtered list
    # if Speed_to_Market_files:
    #     latest_Speed_to_Market_file = max(Speed_to_Market_files, key=os.path.getmtime)
    #     print(f"Latest .xlsm file (excluding 'HC'): {latest_Speed_to_Market_file}")
    # else:
    #     print("No matching .xlsm files found (excluding 'HC').")

    # # Save to OneDrive to Process Static File (no date adjustments needed)
    # shutil.copyfile(latest_Speed_to_Market_file, r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Used_Car_Program\Speed_to_Market.xlsm')

    # # Wait 5 seconds so that file properly saves?
    # time.sleep(5)

    # Open up latest Speed To Market file
    Speed_to_Market_wb = app.books.open(r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Used_Car_Program\Speed_to_Market.xlsm')

    # Process all the speed_to market data/update query
    # APPEND/REPLACE CURRENT MONTH QUERY DATA: LOGIC WILL SEARCH FOR CURRENT MONTH IN THE MTH COLUMN AND REPLACE OTHERWISE APPEND

    # Get current month as a string (e.g., "5" for May)
    current_month = str(datetime.today().month)
    print(current_month) 

    Speed_to_Market_tab1 = Speed_to_Market_wb.sheets['Data1']

    # Read existing data (assuming headers in row 1, data from B2)
    data_range = Speed_to_Market_tab1.range('B1').expand('down').resize(None, 4)
    existing_df = data_range.options(pd.DataFrame, header=1).value

    # Add headers if needed (assuming DataFrame has same column structure)
    existing_df.reset_index(inplace=True)

    # Filter out rows for the current month in MTH column
    filtered_df = existing_df[existing_df['MTH'].astype(int) != int(current_month)]

    # Combine filtered existing data with new data
    updated_df = pd.concat([filtered_df, Speed_to_Market_df], ignore_index=True)

    # Clear full range before re-writing cleaned data
    Speed_to_Market_tab1.range('B2:E500000').clear_contents()

    # Write back updated data
    Speed_to_Market_tab1.range('B2').options(index=False, header=False).value = updated_df

    # Connect to workbook and sheet
    Speed_to_Market_tab2 = Speed_to_Market_wb.sheets['Data2']

    # Find the first empty row in column B
    last_row = Speed_to_Market_tab2.range('B' + str(Speed_to_Market_tab2.cells.last_cell.row)).end('up').row
    start_row = last_row + 1  # Row to start appending

    # Write the DataFrame starting from that row
    Speed_to_Market_tab2.range(f'B{start_row}').options(index=False, header=False).value = Speed_to_Market_df2

    Speed_to_Market_Org_Map_tab = Speed_to_Market_wb.sheets['Org Map']
    Speed_to_Market_Org_Map_tab.range('A2:H500000').clear_contents()
    Speed_to_Market_Org_Map_tab.range('A2').options(index=False, header=False).value = Org_Map_df

    # function that processes latest daily_sales file
    Process_Daily_Sales_File()

    # Open up Dynamic Daily Sales for macro to process file
    Daily_Sales_wb = app.books.open(r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker\Dynamic_Daily_Sales.xlsb')

    # Run Macro that refreshed pivot table and moves data to Used Car File
    Run_Macro = Speed_to_Market_wb.macro("Update_File")
    Run_Macro()
    Speed_to_Market_wb.save()
    Speed_to_Market_wb.close()
    Daily_Sales_wb.close()

   # Save the file in the directory as a new name with todays date
    safe_date_str = today.strftime(r"%Y-%m-%d")  # '2025-05-27'
    new_file_name = f'Used Car Program - Strategy Sheet {safe_date_str}.xlsx'
    new_file = os.path.join(Used_Car_Path, new_file_name)
    wb.save(new_file)

    # Save and close the excel document    
    if wb:
        wb.close()
    if app:
        app.quit()

    # End Timer
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Script completed in {elapsed_time:.2f} seconds")

    return

#run function
if __name__ == '__main__':
    
    Used_Car_Update()