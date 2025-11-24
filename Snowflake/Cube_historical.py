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

def create_historical(Combined_df):

    # Step 4: Create historical
    # 1: Create Store type to discern between AN Store and Competition
    Combined_df['STORE_TYPE'] = np.where(Combined_df['AN_STORE'].isna(), "Competition", 'AN')
    Combined_df['Store_Num'] = np.where(Combined_df['AN_STORE'].isna(), "Competition", Combined_df['AN_STORE'])
    Combined_df['MakeMarketCombo'] = Combined_df['MAKE'].astype(str) + Combined_df['MARKET_NAME'].astype(str)
    
    # 2: Create MTD Sales by Store_Num table
    Grouped_df1 = Combined_df.groupby([
        'Store_Num', 'SELLING_DEALER', 'AUTONATION_MARKET', 'Brand_Group',
        'SEGMENT', 'SALE_MONTH', 'SALE_YEAR', 'MakeMarketCombo', 'MARKET_NAME'
    ], dropna=False)['MTD_SALES'].sum().reset_index()

    Grouped_df1 = Grouped_df1.rename(columns={'SALE_MONTH': 'Time', 'SALE_YEAR': 'Year'})

    # 3: Make second MakeMarketCombo grouping
    Base_output_df = Combined_df.groupby(
        ['Store_Num', 'Brand_Group', 'AUTONATION_MARKET', 'SEGMENT', 'SALE_MONTH', 'SALE_YEAR', 'MakeMarketCombo', 'SELLING_DEALER', 'MARKET_NAME'],
        dropna=False
    ).agg({'MTD_SALES': 'sum'}).reset_index()
    Base_output_df = Base_output_df.rename(columns={'SALE_MONTH': 'Time', 'SALE_YEAR': 'Year'})

    # Compute MakeMarketCombo total MTD_Sales
    totals_df = Base_output_df.groupby(['MakeMarketCombo', 'Time', 'Year'], dropna=False)['MTD_SALES'].sum().reset_index()
    totals_df = totals_df.rename(columns={'MTD_SALES': 'MakeMarketCombo_MTD_Sales'})

    Base_output_df = Base_output_df.merge(
        totals_df,
        on=['MakeMarketCombo', 'Time', 'Year'],
        suffixes=('', '_Total')
    )

    # Step 5: Final calculations
    Base_output_df['Share_Percent'] = (Base_output_df['MTD_SALES'] / Base_output_df['MakeMarketCombo_MTD_Sales']).round(3)

    Base_output_df['Comp_Sales'] = np.where(
        Base_output_df['Store_Num'] == "Competition", 
        Base_output_df['MTD_SALES'], 
        0
    )

    Base_output_df['AN_Sales'] = np.where(
        Base_output_df['Store_Num'] != "Competition", 
        Base_output_df['MTD_SALES'], 
        0
    )

    Base_output_df['Store_Num'] = np.where(
        Base_output_df['Store_Num'] == "Competition", 
        "", 
        Base_output_df['Store_Num']
    )

    # Select final columns and rename
    final_output = Base_output_df[[
        'Store_Num',
        'Brand_Group', 
        'AUTONATION_MARKET',
        'MARKET_NAME',
        'SEGMENT',
        'Share_Percent',
        'Time',
        'Year',
        'MakeMarketCombo_MTD_Sales',
        'MTD_SALES',
        'Comp_Sales',
        'AN_Sales'
    ]].rename(columns={
        'Share_Percent': 'Share_%',
        'MTD_SALES': 'MTD_Sales'
    })
    return final_output

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
    # Step 4: Create historical
    final_output = create_historical(Combined_df)

    final_output.to_csv(r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Market_Share\Cube\historical_df.csv')