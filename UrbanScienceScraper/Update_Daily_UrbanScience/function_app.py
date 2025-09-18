import azure.functions as func
from playwright.sync_api import sync_playwright
from datetime import datetime
import shutil
import os

def Update_Daily_UrbanScience():

    # downloads_folder = r"C:\Users\BesadaG\Downloads"

    # destination_folder = r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Urban_Science"


    downloads_folder = "/tmp/downloads"  # Use temp folder in Azure
    destination_folder = "/tmp/processed"  # Replace with OneDrive upload later

    os.makedirs(downloads_folder, exist_ok=True)
    os.makedirs(destination_folder, exist_ok=True)


    # Target URL
    url = "https://na-ftp.urbanscience.com/ThinClient/WTM/public/index.html#/login"

    # Get today's date in the format YYYYMMDD
    today_str = datetime.today().strftime('%Y%m%d')

    # Construct the expected filename
    filename = f"AutoNation_SalesFile_NationalSales_{today_str}.txt"
    filename_rename = "AutoNation_SalesFile_NationalSales.txt"


    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # Go to login page
        page.goto("https://na-ftp.urbanscience.com/ThinClient/WTM/public/index.html#/login")

        # Login
        page.fill("#inputUsername", "Stuartm")
        page.fill("#password", "f$w8Q)$z%pt)")
        page.click("#signIn")
        page.wait_for_timeout(5000)

        # Click checkbox and download
        page.check(f"input[id*='{filename}']")
        with page.expect_download() as download_info:
            page.click("text=Download")
        download = download_info.value
        download_path = download.path()
        download.save_as(os.path.join(downloads_folder, filename))


        # Move and rename
        source_file = os.path.join(downloads_folder, filename)
        destination_file = os.path.join(destination_folder, filename_rename)
        shutil.move(source_file, destination_file)

        browser.close()

    return

if __name__ == '__main__':
   
    Update_Daily_UrbanScience()
# %%