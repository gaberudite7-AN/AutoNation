#Import libraries
import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import re
import pyodbc
import time
import shutil
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings("ignore", category=UserWarning, message=".*pandas only supports SQLAlchemy.*")

# Selenium packages
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import undetected_chromedriver as uc
import traceback
# Python human simulation
import pyautogui


# Function to download the latest Referall Sharepoint to local
def Download_sharepoint_file():

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
    destination_folder = r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Referrals"
    Sharepoint_Link = "https://autonation-my.sharepoint.com/:x:/r/personal/smithj42_autonation_com/_layouts/15/Doc.aspx?sourcedoc=%7BA68EB55F-02C9-4AAC-842D-E7B83A04BD5A%7D&file=Referral%20Form.xlsx&wdLOR=c487E5101-B465-4479-8BD7-5D0243D0EAFF&fromShare=true&action=default&mobileredirect=true&wdOrigin=TEAMS-MAGLEV.p2p_ns.rwc&wdExp=TEAMS-TREATMENT&wdhostclicktime=1753968148934&web=1"
    filename = "Referral form.xlsx"

    # Start browser
    try:
        #automatically use compatible chrome
        driver = uc.Chrome(
            options=chrome_options,
            use_subprocess=True
        )

        actions = ActionChains(driver)
        driver.set_page_load_timeout(20)
        driver.get(Sharepoint_Link)

        # Define wait AFTER driver is initialized
        wait = WebDriverWait(driver, 20)
        time.sleep(5)

        Email = "besadag@autonation.com"

        time.sleep(2)
        
        # Step 1: Enter email
        email_input = wait.until(EC.presence_of_element_located((By.ID, "i0116")))
        email_input.send_keys(Email)
        time.sleep(2)

        # Step 2: Click the "Next" button using XPath
        next_button = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="idSIButton9"]')))
        next_button.click()
        time.sleep(5)

        # Click the "Continue" button after entering email or password
        continue_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        continue_button.click()
        time.sleep(10)


        # Simulate click file, open, open recent...
        location = pyautogui.locateOnScreen(r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Images_tesseract\Click_File_Referrals.png')
        if location:
            pyautogui.click(location)
        print("Clicked File")
        time.sleep(3)

        # Simulate click file, open, open recent...
        location = pyautogui.locateOnScreen(r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Images_tesseract\Click_Create_a_copy2.png')
        if location:
            pyautogui.click(location)
        print("Clicked create a copy")
        time.sleep(3)

        # Simulate click file, open, open recent...
        location = pyautogui.locateOnScreen(r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Images_tesseract\Click_download_a_copy2.png')
        if location:
            pyautogui.click(location)
        print("Clicked download")
        time.sleep(3)

        # Move latest file to destination folder with adjusted name
        source_file = os.path.join(downloads_folder, filename)
        destination_file = os.path.join(destination_folder, filename)

        shutil.move(source_file, destination_file)
        print(f"Successfully moved file to: {destination_file}")

    except Exception as e:
        error_log_path = os.path.join(destination_folder, "Referral_error_log.txt")
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

def Update_Referral():

    # Select the Excel file using the file selection dialog
    file = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Referrals\Referral form.xlsx'

    # Call the function to read the Excel file
    data = pd.read_excel(file)

    # If the data was successfully loaded, you can manipulate it
    if data is not None:
        # Example operation: displaying the available columns
        #print("Columns available in the file:")
        #print(data.columns)



        df_renamed = data.rename(columns={
        'Email Address\xa0(of the person you are referring)': 'email_customer',
        'Last name\xa0(of the person you are referring)': 'last_name',
        'First name\xa0(of the person you are referring)': 'first_name',
        'Phone Number\xa0(of the person you are referring)': 'phone_number',
        'Based on the location provided, please select the corresponding AutoNation market': 'Market'
    })

    # Remove rows where 'Market' is NaN before creating 'market_prefix'
    df_renamed = df_renamed[df_renamed['Market'].notna()].copy()
    # Create a new column with the first 4 characters of the 'market' column
    df_renamed['market_prefix'] = df_renamed['Market'].astype(str).str[:4]
    # Display the new DataFrame
    #print(df_renamed.head())


    # Reset the WHERE clause before use
    where_clause = ""
    fname_case = ""
    email_case = ""
    phone_case = ""

    # Create a list for the WHERE conditions
    conditions = []
    name_conditions = []
    email_conditions = []
    phone_conditions = []

    # Iterate over the rows of the DataFrame to build the conditions
    for _, row in df_renamed.iterrows():
        
        last_name = row['last_name']
        last_name = last_name.replace("'","")    #remove ' if any  
        cus_first_name = row['first_name']
        cus_first_name = cus_first_name.replace("'","")   #remove ' if any
        market = row['market_prefix']
        email_address = row['email_customer']  # Email Address column
        #phone_number = row['phone_number']    # Phone Number column
        phone_number = re.sub(r'\D', '', row['phone_number'])


        # Generate the condition for each pair of LastName, Market, Email, and Phone
        condition = f"""
        (MarketHyperion = '{market}'
        AND (b.LastName LIKE '%{last_name}%'       
            OR b.email LIKE '%{email_address}%' 
            OR b.telephone LIKE '%{phone_number}%'))
        """
        conditions.append(condition)

        # Conditions for name, email and phone match
        email_conditions.append(f"WHEN b.email = '{email_address}' THEN 'Email_match'")
        phone_conditions.append(f"WHEN b.telephone = '{phone_number}' THEN 'Phone_match'")
        name_conditions.append(f"WHEN (b.FirstName LIKE '%{cus_first_name}%' AND b.LastName LIKE '%{last_name}%') THEN 'name_match'")


    # Join all the conditions with 'OR'
    where_clause = " OR ".join(conditions)

    # Construct CASE statements for email and phone matches
    email_case = "CASE " + " ".join(email_conditions) + " ELSE NULL END AS Email_match" if email_conditions else "NULL AS Email_match"
    phone_case = "CASE " + " ".join(phone_conditions) + " ELSE NULL END AS Phone_match" if phone_conditions else "NULL AS Phone_match"
    first_name_case = "CASE " + " ".join(name_conditions) + " ELSE NULL END AS Name_Match" if phone_conditions else "NULL AS Name_Match"


    # Get current date
    today = datetime.today()

    # take previous month
    reference_date = today.replace(day=1) - timedelta(days=1)  # last day of previous month
    beginning_of_month_dt = reference_date.replace(day=1)

    # Format the dates
    beginning_of_month = f"{beginning_of_month_dt.month}/1/{beginning_of_month_dt.year}"  # e.g. "5/1/2025"

    # Optional: Print results for debugging
    print(f"Accounting month to use is {beginning_of_month}")


    # Build the complete SQL
    sql_query = f"""
    WITH salesquery AS (
        SELECT 
            *
        FROM NDDUsers.vSalesDetail_vehicle
        WHERE accountingmonth >= '{beginning_of_month}'
        AND SaleTrxType = 'VehicleSale'
    )

    SELECT

        {email_case},
        {phone_case},
        {first_name_case},
        CustomerName,
        B.email,
        B.telephone,
        B.CustNo,
        A.DealNo,
        A.AccountingDate,
        A.RegionName,
        A.MarketName,
        A.StoreHyperion,
        A.StoreName,
        A.RecordSource,
        A.FrontGross,
        A.TotalGross,
        A.VehicleSoldCount,
        A.SaleTrxType,
        A.Vin,
        A.VehicleMakeName,
        A.VehicleModelName,
        A.VehicleModelYear,
        A.DepartmentName

        
    FROM salesquery A
    LEFT JOIN BA.vCustomer B
    on A.CustNo = B.CustNo
    WHERE 
        {where_clause}
        
    ORDER BY  Email_match DESC, Phone_match DESC, Name_Match DESC, CustomerName ASC
        
        ;
    """

    # Display the generated SQL (Optional)
    #print(sql_query)

    # Begin timer
    start_time = time.time()
    try:
        with pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=nddprddb01,48155;'
                'DATABASE=NDD_ADP_RAW;'
                'Trusted_Connection=yes;'
        ) as conn:
            Discounted_Inventory_df = pd.read_sql(sql_query, conn)
        end_time = time.time()
        elapsed_time = end_time - start_time
        # Print time it took to load query
        print(f"Script read in Queries and converted to dataframes in {elapsed_time:.2f} seconds")
    except Exception as e:
            print("❌ Connection failed:", e)


    """ NExt step is to compre what is already in last weeks file against the the sql query (avoid duplicating) and update referal form..."""

    # 1. Modify dataframe to only have data if we have a match in at least 2 out of 3 options
    match_count = Discounted_Inventory_df[['Email_match', 'Phone_match', 'Name_Match']].notna().sum(axis=1)

    Discounted_Inventory_df = Discounted_Inventory_df[match_count>=2]
    Discounted_Inventory_df.to_csv(r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Referrals\total_referrals_filtered.csv')

    # 2. Compare to last weeks file to see if we already added them to the list
    Compare_File_Directory = r"W:\Corporate\Inventory\EGadea\WBYC Requests\Referrals"
    latest_referral_file = None
    latest_time = None


    for filename in os.listdir(Compare_File_Directory):
        if filename.endswith('.xlsx') and "Results_referral" in filename:
            filepath = os.path.join(Compare_File_Directory, filename)
            file_time = os.path.getmtime(filepath)

            if latest_time is None or file_time > latest_time:
                latest_time = file_time
                latest_referral_file = filepath

    
    print(f"latest referral file is {latest_referral_file}")
    
    # Load the latest referral file into a DataFrame
    referral_df = pd.read_excel(latest_referral_file)
    referral_df['telephone'] = referral_df['telephone'].astype(str)
    Discounted_Inventory_df['telephone'] = Discounted_Inventory_df['telephone'].astype(str)

    # Filter out rows from current_df where phone_number matches referral_df
    filtered_df = Discounted_Inventory_df[~Discounted_Inventory_df['telephone'].isin(referral_df["telephone"].astype(str))]
    filtered_df.to_csv(r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Referrals\total_referrals_remove_last_month.csv')

    # 3. Bring in the user email from the main referral form file
    # Read in Referral Form file...again...match up the email to pull in the (Please enter your AutoNation email address to confirm your eligibility for the referral program.)
    data = pd.read_excel(file)
    data.columns = data.columns.str.strip()

    data = data.rename(columns={"Email Address\xa0(of the person you are referring)": 'email'})
    

    # Pull in the autonation email address based on email
    filtered_df = filtered_df.merge(data[['email', 'Please enter your AutoNation email address to confirm your eligibility for the referral program.']], on='email', how='left')
    filtered_df = filtered_df.rename(columns={"Please enter your AutoNation email address to confirm your eligibility for the referral program.": 'Email of employee'})

    # move email of employee to front    
    cols = list(filtered_df.columns)
    cols.insert(0, cols.pop(cols.index('Email of employee')))
    filtered_df = filtered_df[cols]

    # Format today's date
    today_str = datetime.today().strftime('%Y.%m.%d')


    filtered_df.to_excel(fr'W:\Corporate\Inventory\EGadea\WBYC Requests\Referrals\Results_referral {today_str}.xlsx', index = False)

if __name__ == '__main__':
    
    # Download_sharepoint_file()
    Update_Referral()