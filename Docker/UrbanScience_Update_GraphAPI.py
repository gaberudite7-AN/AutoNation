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
import xlwings as xw

#ClientSecret10
from msal import ConfidentialClientApplication
import requests

# App credentials
client_id = '3696431a-ff37-4cda-942d-300444982fdf'
client_secret = 'Aib8Q~-IXkotWA0mUzbSzbfhIQf5bKdESCHMWbIX'
tenant_id = 'bd54fbce-74dd-4b5a-8d71-2b978c6d210d'
app = ConfidentialClientApplication(
    client_id,
    authority=f"https://login.microsoftonline.com/{tenant_id}",
    client_credential=client_secret
)
token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
print(token_response)
access_token = token_response['access_token']

# Using graph api for accessing AutoNationSales.txt file in OneDrive
def download_file_from_onedrive(access_token, remote_path, local_path):
    headers = {'Authorization': f'Bearer {access_token}'}
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:{remote_path}:/content"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        with open(local_path, 'wb') as f:
            f.write(response.content)
        print(f"Downloaded {remote_path} to {local_path}")
    else:
        raise Exception(f"Failed to download file: {response.status_code} - {response.text}")

# Using graph api to archive historical files in OneDrive
def upload_file_to_onedrive(access_token, local_path, remote_path):
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/octet-stream'
    }
    with open(local_path, 'rb') as f:
        data = f.read()
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:{remote_path}:/content"
    response = requests.put(url, headers=headers, data=data)
    if response.status_code in [200, 201]:
        print(f"Uploaded {local_path} to {remote_path}")
    else:
        raise Exception(f"Failed to upload file: {response.status_code} - {response.text}")

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
    historics_folder = r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Urban_Science\Historics"
    
    # replacing file_to_move
    #file_to_move = r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Urban_Science\AutoNation_SalesFile_NationalSales.txt"
    download_file_from_onedrive(access_token,
        "/PowerAutomate/Urban_Science/AutoNation_SalesFile_NationalSales.txt",
        "AutoNation_SalesFile_NationalSales.txt")

    filename_modified = f"AutoNation_SalesFile_NationalSales_{date_str}.txt"
    filename_final_file = os.path.join(historics_folder, filename_modified)
    print(f"Copied latest AutoNation_SalesFile_NationalSales.txt file to {filename_final_file}")
    
    shutil.copyfile("AutoNation_SalesFile_NationalSales.txt", filename_final_file)
    upload_file_to_onedrive(access_token,
        filename_final_file,
        f"/PowerAutomate/Urban_Science/Historics/{filename_final_file}")


def Move_Current_to_Historics_Industry():
    

    # Get last week's date as we are moving current that was pulled last week to historics with yesterdays date
    yesterday = datetime.today() - timedelta(days=7)   
    historics_folder = r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Urban_Science\Historics"
    date_str = f"{yesterday.year}{yesterday.month:02d}{yesterday.day:02d}"
    file_to_move = r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Urban_Science\AutoNation_SalesFile_NationalSales_Make.txt"
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
        #automatically use compatible chrome
        driver = uc.Chrome(
            options=chrome_options,
            use_subprocess=True
        )

        driver.set_page_load_timeout(20)
        driver.get(url)

        # Credentials for logging in
        username = "Stuartm"
        password = "f$w8Q)$z%pt)"

        # Login
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "inputUsername"))).send_keys(username)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "password"))).send_keys(password)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "signIn"))).click()
        time.sleep(5)

        # Wait for file checkbox
        checkbox = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, f"input[id*='{filename}']"))
        )
        checkbox.click()
        print(f"Clicked checkbox for file: {filename}")
        time.sleep(5)

        # Click download
        download_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'toolbar-button') and .//span[text()='Download']]"))
        )
        download_button.click()
        print(f"Download button clicked. Waiting for file to download...")
        time.sleep(45)

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
            
            # Wait for file checkbox
            checkbox = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, f"input[id*='{filename}']"))
            )
            checkbox.click()
            print(f"Clicked checkbox for file: {filename}")
            time.sleep(5)

            # Click download
            download_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'toolbar-button') and .//span[text()='Download']]"))
            )
            download_button.click()
            print(f"Download button clicked. Waiting for file to download...")
            time.sleep(60)            

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

try:
    print("bump.")
    Move_Current_to_Historics()
    Update_Historicals() 
except Exception as e:
    with open(r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\error_log.txt", "a") as f:
        f.write(traceback.format_exc())
    # %%
