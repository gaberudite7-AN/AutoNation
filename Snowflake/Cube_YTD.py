# %%
import snowflake.connector
import pandas as pd
import numpy as np
from datetime import datetime, timedelta


# Step 1: Load Current and historical Data then append properly

def load_raw_and_historical_data():
    # Connect to Snowflake
    conn = snowflake.connector.connect(
        account = "HTGXNLD-AN",
        user = "besadag@autonation.com",
        authenticator = "externalbrowser",
        role = "FINANCIAL_PLANNING_ANALYTICS_ANALYST",
        warehouse = "FINANCIAL_PLANNING_ANALYTICS_WH",
        database = "WORKSPACES",
        schema = "FINANCIAL_PLANNING_ANALYTICS"
    )

    # Create a cursor object
    cur = conn.cursor()
    Current_query = """
    select * from workspaces.financial_planning_analytics.daily_sales_current
    """

    # Historical Query
    Historical_query = """
    select * from workspaces.financial_planning_analytics.urban_science_historical
    """

    # Allowed DMA Makes Dimension Table Query
    Dim_Allowed_DMA_Makes_query = """
    select * from workspaces.financial_planning_analytics.dim_allowed_dma_make
    """

    # AutoNation Market Dimension Table Query
    Dim_AutoNation_Market_query = """
    select * from workspaces.financial_planning_analytics.dim_autonation_market
    """

    # AN Brand Dimension Table Query
    Dim_AN_Brand_query = """
    select * from workspaces.financial_planning_analytics.dim_an_brand
    """

    # Dim Store override Query
    Dim_store_override_query = """
    select * from workspaces.financial_planning_analytics.dim_store_overrides
    """

    # Execute and fetch into a DataFrame
    try:
        cur.execute(Current_query)
        Current_df = pd.DataFrame(cur.fetchall(), columns=[col[0] for col in cur.description])

        # Step 33: Filter Current_df for specific columns
        selected_columns = [
            'DMA_CODE', 'MARKET_NAME', 'AN_STORE', 'DEALERID',
            'SITECODE', 'SELLING_DEALER', 'SELLING_DEALER_CITY',
            'SELLING_DEALER_STATE', 'MAKE', 'SEGMENT',
            'SALE_YEAR', 'SALE_MONTH', 'MTD_SALES', 'FILEDATE'
        ]
        
        Current_df = Current_df[selected_columns]

        # Get current date
        today = datetime.today()
        
        # Calculate last year
        last_year = today.year - 1
        
        # Calculate prior two months
        month_1 = (today.replace(day=1) - timedelta(days=15)).month  # last month
        month_2 = (today.replace(day=1) - timedelta(days=45)).month  # two months ago

        # Print out dynamic dates utilized for debugging
        print(f"Last year being used is {last_year}, current months being used are {month_1} and {month_2}")
        
        # Filter out last year's data for the two prior months
        Current_df = Current_df[~((Current_df['SALE_YEAR'] == last_year) & 
                                 (Current_df['SALE_MONTH'].isin([month_2, month_1])))]
        
        # Drop year 2025 and current months from the daily file (WE WONT NEED THIS IN NEXT MONTH)
        # Current_df = Current_df[~((Current_df['SALE_YEAR'] == 2025) & 
        #                          (Current_df['SALE_MONTH'].isin([10])))]

        # Execute historical query
        cur.execute(Historical_query)
        Historical_df = pd.DataFrame(cur.fetchall(), columns=[col[0] for col in cur.description])
        # Execute allowed DMA Dimension table query
        cur.execute(Dim_Allowed_DMA_Makes_query)
        AllowedDMA_df = pd.DataFrame(cur.fetchall(), columns=[col[0] for col in cur.description])

        # Execute AutoNation Market Dimension table query
        cur.execute(Dim_AutoNation_Market_query)
        AutoNation_Market_df = pd.DataFrame(cur.fetchall(), columns=[col[0] for col in cur.description])

        # Execute AN Brand Dimension table query
        cur.execute(Dim_AN_Brand_query)
        AN_Brand_df = pd.DataFrame(cur.fetchall(), columns=[col[0] for col in cur.description])

        # Execute Store override Dimension table query
        cur.execute(Dim_store_override_query)
        Store_override_df = pd.DataFrame(cur.fetchall(), columns=[col[0] for col in cur.description])

        return Current_df, Historical_df, AllowedDMA_df, AutoNation_Market_df, AN_Brand_df, Store_override_df

    finally:
        cur.close()
        conn.close()

def create_ytd_summary(df):
    """
    Generates YTD summary per brand/market/store combination
    """
    from datetime import date
    import calendar

    # 1: Create Store type to discern between AN Store and Competition
    Combined_df['STORE_TYPE'] = np.where(Combined_df['AN_STORE'].isna(), "Competition", 'AN')
    Combined_df['Store_Num'] = np.where(Combined_df['AN_STORE'].isna(), "Competition", Combined_df['AN_STORE'])
    Combined_df['MakeMarketCombo'] = Combined_df['MAKE'].astype(str) + Combined_df['MARKET_NAME'].astype(str)

    # 2: Prepare base monthly MTD summary
    # Create Store_Num column
    df['STORE_NUM'] = np.where(
        (df['AN_STORE'].isna()) | (df['AN_STORE'] == ''), 
        'Competition', 
        df['AN_STORE']
    )
    
    # Group by monthly data
    Table1_YTD = df.groupby([
        'STORE_NUM', 'SELLING_DEALER', 'AUTONATION_MARKET', 'Brand_Group',
        'SEGMENT', 'SALE_MONTH', 'SALE_YEAR', 'MakeMarketCombo', 'MARKET_NAME'
    ], dropna=False)['MTD_SALES'].sum().reset_index()
    
    current_month_num = date.today().month
    
    def calc_ytd_for_month(m: int):
        # Create time label
        month_name = calendar.month_abbr[m]
        TimeLabel_YTD = f"YTD{month_name}"
        
        # Filter data up to month m and aggregate YTD sales
        Table_YTD_Filtered = Table1_YTD[Table1_YTD['SALE_MONTH'] <= m].copy()
        
        Table_YTD_Aggregated = Table_YTD_Filtered.groupby([
            'STORE_NUM', 'SELLING_DEALER', 'AUTONATION_MARKET', 'Brand_Group',
            'MakeMarketCombo', 'SALE_YEAR', 'MARKET_NAME', 'SEGMENT'
        ], dropna=False)['MTD_SALES'].sum().reset_index()
        
        Table_YTD_Aggregated = Table_YTD_Aggregated.rename(columns={'MTD_SALES': 'YTD_SALES'})
        Table_YTD_Aggregated['TIME'] = TimeLabel_YTD
        
        # Calculate MakeMarketCombo aggregation
        MakeMarketComboAgg = Table_YTD_Aggregated.groupby([
            'MakeMarketCombo', 'SALE_YEAR', 'TIME'
        ], dropna=False)['YTD_SALES'].sum().reset_index()
        
        MakeMarketComboAgg = MakeMarketComboAgg.rename(columns={'YTD_SALES': 'MAKEMARKETCOMBO_YTD_SALES'})
        
        # Join back to get final result
        Final = Table_YTD_Aggregated.merge(
            MakeMarketComboAgg, 
            on=['MakeMarketCombo', 'SALE_YEAR', 'TIME'], 
            how='left'
        )
        
        # Filter out empty store numbers and calculate final metrics
        Final = Final[(Final['STORE_NUM'] != '') & (Final['STORE_NUM'].notna())].copy()
        
        Final['SHARE_PCT'] = (Final['YTD_SALES'] / Final['MAKEMARKETCOMBO_YTD_SALES']).round(3)
        
        Final['COMP_SALES'] = np.where(
            Final['STORE_NUM'] == 'Competition', 
            Final['YTD_SALES'], 
            0
        )
        
        Final['AN_SALES'] = np.where(
            Final['STORE_NUM'] != 'Competition', 
            Final['YTD_SALES'], 
            0
        )
        
        # Select final columns
        Final = Final[[
            'STORE_NUM', 'Brand_Group', 'AUTONATION_MARKET', 'MARKET_NAME',
            'SEGMENT', 'SHARE_PCT', 'TIME', 'SALE_YEAR', 'MAKEMARKETCOMBO_YTD_SALES',
            'YTD_SALES', 'COMP_SALES', 'AN_SALES'
        ]].rename(columns={'SALE_YEAR': 'YEAR'})
        
        return Final
    
    # Generate YTD data for all months up to current month
    ytd_dfs = []
    for m in range(1, current_month_num + 1):
        ytd_df = calc_ytd_for_month(m)
        ytd_dfs.append(ytd_df)
    
    # Combine all YTD dataframes
    YTD_AllMonths_Output = pd.concat(ytd_dfs, ignore_index=True)
    
    return YTD_AllMonths_Output

#run functions step by step
if __name__ == '__main__':
    # Step 1: Load Current and historical Data then append properly
    Current_df, Historical_df, AllowedDMA_df, AutoNation_Market_df, AN_Brand_df, Store_override_df = load_raw_and_historical_data()

    # Step 2: Append the two dataframes
    Combined_df = pd.concat([Current_df, Historical_df], ignore_index=True)

    # Drop Segment
    Combined_df = Combined_df.drop(columns=['SEGMENT'])
    # Step 3: Transformations
    # 1. Bring in the Allowed DMAS using DIM Table
    # Create DMA_Make Key
    Combined_df['DMA_MAKE'] = Combined_df['MARKET_NAME'].astype(str) + Combined_df['MAKE']
    Combined_df = Combined_df.merge(AllowedDMA_df, on='DMA_MAKE', how='inner')

    # 2. Bring in Autonation Market using DIM Table
    Combined_df = Combined_df.merge(AutoNation_Market_df, on='MARKET_NAME', how='left')

    # 3: Bring in Brand group using DIM Table
    Combined_df = Combined_df.merge(AN_Brand_df, on='MAKE', how='left')
    Combined_df = Combined_df.rename(columns={'AN_BRAND': 'Brand_Group'})

    # 4: Bring in Store Override using DIM Table
    # Convert SITECODE to same data type in both DataFrames
    Combined_df['SITECODE'] = Combined_df['SITECODE'].astype(str)
    Store_override_df['SITECODE'] = Store_override_df['SITECODE'].astype(str)
    Combined_df = Combined_df.merge(Store_override_df, on='SITECODE', how='left')

    # Adjust AN_Store to bring in Override Data
    Combined_df['AN_STORE'] = np.where(
        Combined_df['AN_STORE'] == '', 
        Combined_df['AN_STORE_OVERRIDE'], 
        Combined_df['AN_STORE']
    )

    Combined_df['SELLING_DEALER'] = np.where(
        Combined_df['SELLING_DEALER'] == '', 
        Combined_df['SELLING_DEALER_OVERRIDE'], 
        Combined_df['SELLING_DEALER']
    )

    # Drop old AN Store and Selling Dealer
    Combined_df = Combined_df.drop(columns=['AN_STORE_OVERRIDE', 'SELLING_DEALER_OVERRIDE'])

    # 5: Create Segment
    domestic = ["Ford", "Chrysler", "GM"]
    import_brands = [
        "Acura", "Honda", "Hyundai", "Infiniti", "Mazda", "Nissan",
        "Subaru", "Toyota", "Volkswagen", "Volvo"
    ]
    luxury = [
        "Aston Martin", "Audi", "BMW", "Bentley", "Jaguar Land Rover",
        "Lexus", "Mercedes", "Mini", "Other Lux", "Porsche"
    ]

    # Using nested numpy.where
    Combined_df['SEGMENT'] = np.where(
        Combined_df['Brand_Group'].isin(domestic), 'Domestic',
        np.where(
            Combined_df['Brand_Group'].isin(import_brands), 'Import',
            np.where(
                Combined_df['Brand_Group'].isin(luxury), 'Luxury',
                'Other'
            )
        )
    )

    # Step 4: Create YTD
    final_output = create_ytd_summary(Combined_df)

    final_output.to_csv(r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Market_Share\Cube\YTD_df.csv')