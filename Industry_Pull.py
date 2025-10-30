# %%
import shutil
import snowflake.connector
import pandas as pd
import xlwings as xw
from pathlib import Path
from snowflake.connector.pandas_tools import write_pandas

def find_latest_file_in_dir(directory):
    p = Path(directory)
    if not p.exists() or not p.is_dir():
        return None
    files = [f for f in p.iterdir() if f.is_file()]
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime)

def Industry_Load():
    # Load latest Industry data to Snowflake
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

    # Grab most recent Industry file
    Industry_File_path = r'W:\Corporate\Inventory\Urban Science\Historics\Industry\CSV_Formatted'
    latest_file = find_latest_file_in_dir(Industry_File_path)
    if latest_file is None:
        print("No files found.")
        return
    print("Latest file found:", latest_file)

    # Load the latest file into a DataFrame
    df = pd.read_csv(latest_file)

    # Optional: normalize column names to match Snowflake (uppercase)
    df.columns = [c.upper() for c in df.columns]

    # Use write_pandas to upload/appended data to Snowflake table.
    try:
        success, nchunks, nrows, _ = write_pandas(
            conn,
            df,
            table_name="URBAN_SCIENCE_INDUSTRY",
            database="WORKSPACES",
            schema="FINANCIAL_PLANNING_ANALYTICS"
        )
        if success:
            print(f"Data appended successfully: {nrows} rows in {nchunks} chunks.")
        else:
            print("write_pandas reported failure.")
    finally:
        conn.close()

    print("Data loaded successfully.")

def Update_Industry_UrbanScience():
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

    # We will be pulling the latest data from snowflake for Cube Historical and YTD. 
    # Snowflake query is dynamic and will update latest daily file and append to historical while making the necessary transformations.
    cur = conn.cursor()
    Industry_query = """
    select * from workspaces.financial_planning_analytics.Urban_Science_Industry
    """

    # Execute and fetch into a DataFrame
    try:
        cur.execute(Industry_query)
        Industry_df = pd.DataFrame(cur.fetchall(), columns=[col[0] for col in cur.description])

        # Create AN_Brand and AN_Segment columns

        # 1. AN BRAND
        # Dictionary mapping Make codes to brand names
        make_to_brand = {
            "HOND": "Honda",
            "FORD": "Ford",
            "MMNA": "Mitsubishi",
            "CADI": "GM",
            "CHEV": "GM",
            "GMCT": "GM",
            "DODG": "Chrysler",
            "JEEP": "Chrysler",
            "RAM": "Chrysler",
            "TOY": "Toyota",
            "NISS": "Nissan",
            "KIA": "Kia",
            "BUIC": "GM",
            "MAZD": "Mazda",
            "HYUN": "Hyundai",
            "VOLK": "Volkswagen",
            "MB": "Mercedes",
            "BMW": "BMW",
            "LINC": "Ford",
            "CHRY": "Chrysler",
            "ALFA": "Other Lux",
            "AUDI": "Audi",
            "VOLV": "Volvo",
            "SUBA": "Subaru",
            "GEN": "Hyundai",
            "ROV": "Jaguar Land Rover",
            "INF": "Infiniti",
            "SPRT": "Mercedes",
            "MINI": "MINI",
            "LEX": "Lexus",
            "ACUR": "Acura",
            "PORS": "Porsche",
            "TESL": "Tesla",
            "FIAT": "Fiat",
            "AML": "Aston Martin",
            "BEN": "Bentley",
            "FERR": "Other Lux",
            "JAG": "Jaguar Land Rover",
            "MASR": "Other Lux",
            "LAMB": "Other Lux",
            "LOT": "Other Lux",
            "MCLA": "Other Lux",
            "POLE": "Polestar",
            "INEO": "Other Lux",
            "RR": "Other Lux",
            "LCID": "Other Lux",
            "RIVN": "Rivian"
        }

        # 2. AN SEGMENT    
        # Define brand categories
        domestic_brands = {"Ford", "Chrysler", "GM"}
        import_brands = {
            "Acura", "Honda", "Hyundai", "Infiniti", "Mazda", "Nissan",
            "Subaru", "Toyota", "Volkswagen", "Volvo"
        }
        luxury_brands = {
            "Aston Martin", "Audi", "BMW", "Bentley", "Jaguar Land Rover",
            "Lexus", "Mercedes", "Mini", "Other Lux", "Porsche"
        }

        # Function to classify brand
        def classify_brand(brand):
            if brand in domestic_brands:
                return "Domestic"
            elif brand in import_brands:
                return "Import"
            elif brand in luxury_brands:
                return "Luxury"
            else:
                return "Other"


        # Apply mapping to create a new column
        Industry_df['AN_Brand'] = Industry_df['MAKE'].map(make_to_brand).fillna('Unknown')

        # Apply classification
        Industry_df['AN_Segment'] = Industry_df['AN_Brand'].apply(classify_brand)

        # Open file and process macro/Sql
        Industry_File = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Market_Share\Industry\Urban_Industry_Data.xlsb'
        app = xw.App(visible=True)
        wb = app.books.open(Industry_File)

        Industry_tab = wb.sheets['Data']
        Industry_tab.range("A1:H100000").clear_contents()
        Industry_tab.range('A1').options(index=False, header=True).value = Industry_df

        # Run Macro
        Run_Macro = wb.macro("Refresh_Workbook")
        Run_Macro()
        wb.save()

        # Save and close the excel document    
        if wb:
            wb.close()
        if app:
            app.quit()
        
        # Send a copy to archive
        Archive_File = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Market_Share\Industry\Archive\Urban_Industry_Data_' + pd.Timestamp.now().strftime('%Y%m%d') + '.xlsb'
        shutil.copy(Industry_File, Archive_File)

        # Send a copy to W Drive
        shutil.copy(Industry_File, r'W:\Corporate\Inventory\Reporting\JDPower Industry vs AN\Urban_Industry_Data.xlsb')

    finally:
        cur.close()
        conn.close()

if __name__ == "__main__":
    
    # Load latest make csv file to snowflake
    Industry_Load()
    # Update Urban Science Industry Excel
    Update_Industry_UrbanScience()