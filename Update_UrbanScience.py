# %%
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc
import time
from datetime import datetime, timedelta
import shutil
import os
import traceback
import pandas as pd
import numpy as np
from datetime import datetime
from openpyxl import load_workbook
import xlwings as xw
import csv


def Move_Current_to_Historics():
    
    # Determine how many days to subtract based on the day of the week
    today = datetime.today()
    if today.weekday() == 0:  # Monday
        delta_days = 3
    else:
        delta_days = 1

    # Calculate the adjusted date
    adjusted_date = today - timedelta(days=delta_days)
    date_str = f"{adjusted_date.year}{adjusted_date.month:02d}{adjusted_date.day:02d}"
    historics_folder = r"\\us1.autonation.com\workgroups\Corporate\Inventory\Urban Science\Historics"
    file_to_move = r"\\us1.autonation.com\workgroups\Corporate\Inventory\Urban Science\AutoNation_SalesFile_NationalSales.txt"
    filename_modified = f"AutoNation_SalesFile_NationalSales_{date_str}.txt"
    filename_final_file = os.path.join(historics_folder, filename_modified)
    print(f"Copied latest AutoNation_SalesFile_NationalSales.txt file to {filename_final_file}")
    
    shutil.copyfile(file_to_move, filename_final_file)

def Move_Current_to_Historics_Industry():
    

    # Get last week's date as we are moving current that was pulled last week to historics with yesterdays date
    yesterday = datetime.today() - timedelta(days=7)   
    historics_folder = r"\\us1.autonation.com\workgroups\Corporate\Inventory\Urban Science\Historics\Industry"
    date_str = f"{yesterday.year}{yesterday.month:02d}{yesterday.day:02d}"
    file_to_move = r"\\us1.autonation.com\workgroups\Corporate\Inventory\Urban Science\AutoNation_SalesFile_NationalSales_Make.txt"
    filename_modified = f"AutoNation_SalesFile_NationalSales_Make_{date_str}.txt"
    filename_final_file = os.path.join(historics_folder, filename_modified)
    print(f"Copied latest AutoNation_SalesFile_NationalSales_Make.txt file to {filename_final_file}")
    
    shutil.copyfile(file_to_move, filename_final_file)

def Update_Daily_UrbanScience():

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
    #destination_folder = r"W:\Corporate\Inventory\Urban Science\Historics"

    destination_folder = r"\\us1.autonation.com\workgroups\Corporate\Inventory\Urban Science"

    # Target URL
    url = "https://na-ftp.urbanscience.com/ThinClient/WTM/public/index.html#/login"

    # Get today's date in the format YYYYMMDD
    today_str = datetime.today().strftime('%Y%m%d')

    # Construct the expected filename
    filename = f"AutoNation_SalesFile_NationalSales_{today_str}.txt"
    filename_rename = "AutoNation_SalesFile_NationalSales.txt"

    # Start browser
    try:
        # use downloaded chrome path
        driver = uc.Chrome(
            driver_executable_path=chrome_driver_path,
            options=chrome_options,
            use_subprocess=True
        )

        #automatically use compatible chrome
        # driver = uc.Chrome(
        #     options=chrome_options,
        #     use_subprocess=True
        # )

        driver.set_page_load_timeout(20)
        driver.get(url)

        # Credentials for logging in
        username = "Stuartm"
        password = "f$w8Q)$z%pt)"

        # Login
        time.sleep(5)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "inputUsername"))).send_keys(username)
        time.sleep(1)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "password"))).send_keys(password)
        time.sleep(1)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "signIn"))).click()
        time.sleep(10)

        # Wait for the row containing the filename
        row = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located(
                (By.XPATH, f"//tr[.//div[@class='table-name' and normalize-space(text())='{filename}']]")
            )
        )

        # Scroll the row into view
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row)
        time.sleep(1)

        # Inside that row, find the checkbox and click it
        checkbox = row.find_element(By.XPATH, ".//input[@type='checkbox']")
        checkbox.click()
        print(f"{filename} clicked...")
        time.sleep(5)

        # Click download
        download_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'toolbar-button') and .//span[text()='Download']]"))
        )
        download_button.click()
        print(f"Download button clicked. Waiting for file to download...")
        time.sleep(30)

        # Move latest file to destination folder with adjusted name
        source_file = os.path.join(downloads_folder, filename)
        destination_file = os.path.join(destination_folder, filename_rename)

        if os.path.exists(source_file):
            shutil.move(source_file, destination_file)
            print(f"Successfully moved file to: {destination_file}")
        else:
            raise FileNotFoundError(f"Expected file not found: {source_file}")
        
        # Construct the expected filename
        filename = f"AutoNation_SalesFile_NationalSales_Make.txt"
        
        now = datetime.now()
        
        if now.weekday() == 2: # if its wednesday...process the make file
            print("Day is wednesday..downloading make file")            
            
            # Wait for the row containing the filename
            row = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, f"//tr[.//div[@class='table-name' and normalize-space(text())='{filename}']]")
                )
            )

            # Scroll the row into view
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row)
            time.sleep(1)

            # Inside that row, find the checkbox and click it
            checkbox = row.find_element(By.XPATH, ".//input[@type='checkbox']")
            checkbox.click()
            print(f"{filename} clicked...")
            time.sleep(2)

            # Click download
            download_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'toolbar-button') and .//span[text()='Download']]"))
            )
            download_button.click()
            print(f"Download button clicked. Waiting for file to download...")
            time.sleep(30)            

            filename = "AutoNation_SalesFile_NationalSales_Make.txt"

            # Move latest MAKE file to destination folder with same name
            source_file = os.path.join(downloads_folder, filename)
            destination_file = os.path.join(destination_folder, filename)

            if os.path.exists(source_file):
                shutil.move(source_file, destination_file)
                print(f"Successfully moved Make file to: {destination_file}")
            else:
                raise FileNotFoundError(f"Expected file not found: {source_file}")
        else:
            print("Day is not wednesday, finishing up Web Scraping function.")
    except Exception as e:
        error_log_path = os.path.join(destination_folder, "UrbanScienceScrape_error_log.txt")
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
    
    return

# Convert AN_Store from decimal to integer string, preserving nulls
def clean_store(val):
    if pd.isna(val):
        return np.nan
    try:
        return str(int(float(val)))
    except:
        return val  # If it can't convert, keep original

def Update_Historicals():
    now = datetime.now()
    if now.weekday() == 2:
        # if its wenesday update historics for make and current national sales file
        print(f"Today's date is {now.strftime('%A')}, so we will update Make and regular file")
        Move_Current_to_Historics()
        Move_Current_to_Historics_Industry()
    else: #otherwise simply update historics for national sales file
        print(f"todays date is {now.strftime('%A')}, so we update ONLY regular file")
        Move_Current_to_Historics()
    
    return

def Refresh_MarketShare_Excels():

    # Prior Week File
    UrbanScience_Main = r"W:\Corporate\Inventory\Urban Science\Excel_Update\MarketShare_FOR INTERNAL VARIABLE OPS ONLY.xlsm"

    app = xw.App(visible=True) 
    UrbanScience_Main_wb = app.books.open(UrbanScience_Main)

    # Run Macro to refresh all data connections
    Run_Macro = UrbanScience_Main_wb.macro("Refresh_File")
    Run_Macro()

    # Save and close the excel document(s)
    # Save both workbooks
    UrbanScience_Main_wb.save()
    time.sleep(5)
    UrbanScience_Main_wb.save()
    time.sleep(5)

    # Close workbooks and quit excel app
    if UrbanScience_Main_wb:
        UrbanScience_Main_wb.close()
    if app:
        app.quit()

    return

def Update_Make_File():

    # Move Make file to historicals
    Move_Current_to_Historics_Industry()

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
    #destination_folder = r"W:\Corporate\Inventory\Urban Science\Historics"

    destination_folder = r"\\us1.autonation.com\workgroups\Corporate\Inventory\Urban Science"

    # Target URL
    url = "https://na-ftp.urbanscience.com/ThinClient/WTM/public/index.html#/login"

    # Get today's date in the format YYYYMMDD
    today_str = datetime.today().strftime('%Y%m%d')

    # Construct the expected filename
    filename = f"AutoNation_SalesFile_NationalSales_{today_str}.txt"
    filename_rename = "AutoNation_SalesFile_NationalSales.txt"

    # Start browser
    try:
        # use downloaded chrome path
        driver = uc.Chrome(
            driver_executable_path=chrome_driver_path,
            options=chrome_options,
            use_subprocess=True
        )

        driver.set_page_load_timeout(20)
        driver.get(url)

        # Credentials for logging in
        username = "Stuartm"
        password = "f$w8Q)$z%pt)"

        # Login
        time.sleep(5)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "inputUsername"))).send_keys(username)
        time.sleep(1)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "password"))).send_keys(password)
        time.sleep(1)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "signIn"))).click()
        time.sleep(10)
        
        # Construct the expected filename
        filename = f"AutoNation_SalesFile_NationalSales_Make.txt"
        
        now = datetime.now()
        
        if now.weekday() == 2: # if its wednesday...process the make file
            print("Day is wednesday..downloading make file")            
            
            # Wait for the row containing the filename
            row = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, f"//tr[.//div[@class='table-name' and normalize-space(text())='{filename}']]")
                )
            )

            # Scroll the row into view
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row)
            time.sleep(1)

            # Inside that row, find the checkbox and click it
            checkbox = row.find_element(By.XPATH, ".//input[@type='checkbox']")
            checkbox.click()
            print(f"{filename} clicked...")
            time.sleep(2)

            # Click download
            download_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'toolbar-button') and .//span[text()='Download']]"))
            )
            download_button.click()
            print(f"Download button clicked. Waiting for file to download...")
            time.sleep(30)            

            filename = "AutoNation_SalesFile_NationalSales_Make.txt"
            file_name_rename = f"AutoNation_SalesFile_NationalSales_Make_{today_str}.txt"

            # Move latest MAKE file to destination folder with same name
            source_file = os.path.join(downloads_folder, filename)
            destination_file = os.path.join(destination_folder, filename)
            industry_folder = r"W:\Corporate\Inventory\Urban Science\Historics\Industry"


            if os.path.exists(source_file):
                shutil.move(source_file, destination_file)
                print(f"Successfully moved Make file to: {destination_file}")
                shutil.copy(destination_file, os.path.join(industry_folder, file_name_rename))
                print("Successfully copied today's make file to Industry for Snowflake load processing")
                time.sleep(2)
                convert_textfiles_to_csv()
                print("Successfully converted Industry text files to CSV format")
            else:
                raise FileNotFoundError(f"Expected file not found: {source_file}")
        else:
            print("Day is not wednesday, finishing up Web Scraping function.")
    except Exception as e:
        error_log_path = os.path.join(destination_folder, "UrbanScienceScrape_error_log.txt")
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
    
    return

def convert_textfiles_to_csv():

    # Mode: 'copy' = copy/rename .txt -> .csv (keeps file contents unchanged)
    #       'parse' = read lines, split by commas and write proper CSV rows
    mode = 'copy'  # set to 'parse' if you want to split lines into CSV cells

    # Define source and destination folders
    source_folder = r'W:\Corporate\Inventory\Urban Science\Historics\Industry'
    destination_folder = r'W:\Corporate\Inventory\Urban Science\Historics\Industry\CSV_Formatted'

    # Create destination folder if it doesn't exist
    os.makedirs(destination_folder, exist_ok=True)

    # Iterate through all .txt files in the source folder
    for filename in os.listdir(source_folder):
        if filename.endswith('.txt'):
            txt_path = os.path.join(source_folder, filename)
            csv_filename = os.path.splitext(filename)[0] + '.csv'
            csv_path = os.path.join(destination_folder, csv_filename)

            if mode == 'copy':
                # Copy the file and give it a .csv extension (same content as .txt)
                shutil.copy2(txt_path, csv_path)
            else:
                # Read the text file and write to CSV (split by commas)
                with open(txt_path, 'r', encoding='utf-8-sig') as txt_file:
                    lines = txt_file.readlines()

                with open(csv_path, 'w', newline='', encoding='utf-8') as csv_file:
                    writer = csv.writer(csv_file)
                    for line in lines:
                        writer.writerow([cell.strip() for cell in line.strip().split(',')])

    print("All .txt files have been converted to .csv and copied to the destination folder.")


#run function
if __name__ == '__main__':

    #Update_Historicals()    
    #Update_Daily_UrbanScience()
    #Update_Industry_UrbanScience()
    Update_Make_File()
    #Refresh_MarketShare_Excels()
# %%
