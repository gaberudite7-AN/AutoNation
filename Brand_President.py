from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc
import time
from datetime import datetime, timedelta
import shutil
import os
import traceback
import pyautogui
import pytesseract
import pygetwindow as gw
import xlwings as xw
import pyodbc
import pandas as pd
import warnings
import cv2
import numpy as np
import glob
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
        return

    else:   
        # Rename file to Dynamic_Daily_Sales.xlsb    
        shutil.copyfile(Dynamic_Daily_Sales_File, r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Allocation_Tracker\Dynamic_Daily_Sales.xlsb')
                
        # Delete old file
        os.remove(Dynamic_Daily_Sales_File)
        print("Successfully removed old file and replaced current with new.")
        return

def click_image_with_opencv_and_pyautogui(image_path, confidence=0.95, grayscale=False):

    """
    Searches for an image on the screen using OpenCV and clicks it if found.

    Parameters:
    - image_path (str): Full path to the image file to locate.
    - confidence (float): Matching confidence threshold (default is 0.95).
    - grayscale (bool): Whether to use grayscale matching (default is False).

    Returns:
    - bool: True if image was found and clicked, False otherwise.
    """
    try:
        # Take a screenshot and convert to OpenCV format
        screenshot = pyautogui.screenshot()
        screenshot = cv2.cvtColor(np.array(screenshot), cv2.COLOR_RGB2BGR)

        # Load the template image
        template = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE if grayscale else cv2.IMREAD_COLOR)
        if template is None:
            print(f"Error: Could not load image from {image_path}")
            return False

        # Convert screenshot to grayscale if needed
        if grayscale:
            screenshot = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)

        # Match template
        result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(result)

        if max_val >= confidence:
            h, w = template.shape[:2]
            center_x = max_loc[0] + w // 2
            center_y = max_loc[1] + h // 2
            pyautogui.click(center_x, center_y)
            print(f"Clicked on image: {image_path}")
            return True
        else:
            print(f"Image not found (confidence={max_val:.3f})")
            return False

    except Exception as e:
        print(f"Error during image search: {e}")
        return False    


def Update_DOC_AND_BUDGET_file():

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
    filename = f"Build.ica"
    destination = r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Brand_President_Tracker\Build.ica"

    # Target URL
    url = "https://www.dealercentral.net/Pages/home.aspx"

    # Start browser
    try:
        driver = uc.Chrome(
            driver_executable_path=chrome_driver_path,
            options=chrome_options,
            use_subprocess=True
        )

        driver.set_page_load_timeout(20)
        driver.get(url)

        # Credentials
        username = "besadag"
        password = "T0ttenh@mG-DZ@ndy10!"

        # Login
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "signInControl_UserName"))).send_keys(username)
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "signInControl_password"))).send_keys(password)
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "signInControl_login"))).click()
        time.sleep(5)
 
        time.sleep(5)

        # Click EPM - Excel 2019
        download_button = driver.find_element(By.LINK_TEXT, "EPM - Excel 2019")
        download_button.click()
        print(f"EPM Clicked. Waiting for file to download...")
        time.sleep(15)

        # Open source file and interact with it using PyAutoGUI or AUtoIT
        source_file = os.path.join(downloads_folder, filename)
        print("Opening up file")
        os.startfile(source_file)
        time.sleep(30) # wait for excel to open

        # Simulate click file, open, open recent...
        image_path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Images_tesseract\Click_File.png'
        click_image_with_opencv_and_pyautogui(image_path, confidence=0.95, grayscale=True)
        time.sleep(2)
        
        # Simulate click file, open, open recent...
        image_path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Images_tesseract\Recent_DOC_And_Budget_Click2.png'
        click_image_with_opencv_and_pyautogui(image_path, confidence=0.85, grayscale=True)
        time.sleep(3)


        image_path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Images_tesseract\Click_Smart_View.png'
        click_image_with_opencv_and_pyautogui(image_path, confidence=0.85, grayscale=True)
        time.sleep(3)

        # Adjust Month, type in current month
        current_month = datetime.now().strftime('%b')
        print(current_month)

        yesterday = datetime.now() - timedelta(1)
        Doc_Input = "DOCFcst" + str(yesterday.day)
        print(Doc_Input)

        # Click Docfcst and adjust current day
        pyautogui.click(x=675, y=265)
        time.sleep(2)
        pyautogui.write(Doc_Input)
        # Click Docfcst go to month with pyautogui
        pyautogui.click(x=675, y=400)
        time.sleep(2)
        pyautogui.write(current_month)
        # Click Budget with pyautogui
        time.sleep(2)
        pyautogui.click(x=1100, y=400)
        time.sleep(2)
        pyautogui.write(current_month)
                
        # # Escape
        time.sleep(2)
        pyautogui.press('esc')
        pyautogui.click(x=200, y=400) 

        "Utilize OpenCV to more accurately find images"
        # Click Refresh button        
        image_path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Images_tesseract\Click_Refresh.png'
        click_image_with_opencv_and_pyautogui(image_path, confidence=0.95, grayscale=True)

        # Step 2: Wait a moment for the field to activate
        time.sleep(5)

        # Step 3: Type the password
        pyautogui.write('T0ttenh@mG-DZ@ndy10!', interval=0.1)
        time.sleep(2)

        "Utilize OpenCV to more accurately find images" 
        # Click Connect       
        image_path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Images_tesseract\Click_Connect.png'
        click_image_with_opencv_and_pyautogui(image_path, confidence=0.95, grayscale=True)
        time.sleep(3) 
        # Click File (before saving)
        image_path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Images_tesseract\Click_File.png'
        click_image_with_opencv_and_pyautogui(image_path, confidence=0.95, grayscale=True)
        time.sleep(3) 
        # Click Save
        image_path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Images_tesseract\Click_Save.png'
        click_image_with_opencv_and_pyautogui(image_path, confidence=0.95, grayscale=True)
        time.sleep(3)
        # Click File (before closing)
        image_path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Images_tesseract\Click_File.png'
        click_image_with_opencv_and_pyautogui(image_path, confidence=0.95, grayscale=True)
        time.sleep(3) 
        # Click Close
        image_path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Images_tesseract\Click_Close.png'
        click_image_with_opencv_and_pyautogui(image_path, confidence=0.95, grayscale=True)
        time.sleep(3)      

    finally:
        print("Completed auto-update of Essbase data from the visual")
        # Patch the __del__ method to suppress the OSError
        def safe_del(self):
            try:
                self.quit()
            except Exception:
                pass  # Silently ignore all errors
        uc.Chrome.__del__ = safe_del
    
    return

def Update_BPU_File():

    
    # Run SQL queries using SQL Alchemy and dump into Data tab
    NDD_query = """
    SELECT -- PIVOT TABLE
        CASE
            WHEN AccountingMonth = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 1, 0) THEN 'PM'
            WHEN AccountingMonth = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0) THEN 'CM'
            WHEN AccountingMonth = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 12, 0) THEN 'PY'
            ELSE 'ERROR'
        END AS Period,
        AccountingMonth,
        Hyperion,
        StoreName,
        Brand,
        SUM(CASE WHEN Period = 'Ending' THEN InvQTY ELSE 0 END) AS InvTotal,
        SUM(AgedInv) AS AgedInvCount,
        SUM(SoldCount) AS SoldCount,
        SUM(TotalBasePVR) AS BaseGross,
        SUM(TotalCorePVR) AS FrontGross,
        CASE 
            WHEN SUM(SoldCount) = 0 THEN NULL 
            ELSE SUM(TotalBasePVR) / SUM(SoldCount) 
        END AS BasePVR,
        CASE 
            WHEN SUM(SoldCount) = 0 THEN NULL 
            ELSE SUM(TotalCorePVR) / SUM(SoldCount) 
        END AS CorePVR,
        SUM(AccountingSoldCount) AS AccountingSoldCount
    FROM (
        -- First Query: Ending Inventory Data
        SELECT 
            AccountingMonth,
            Hyperion,
            StoreName,
            CASE 
                WHEN Make IN ('Buick', 'Cadillac', 'Chevrolet', 'GMC') THEN 'GM'
                WHEN Make IN ('Dodge', 'Fiat', 'Jeep', 'Ram', 'Chrysler') THEN 'Chrysler'
                WHEN Make = 'Genesis' THEN 'Hyundai'
                WHEN Make = 'Lincoln' THEN 'Ford'
                WHEN Make IN ('Jaguar', 'Land Rover') THEN 'JLR'
                ELSE Make
            END AS Brand,
            SUM(InventoryCount) AS InvQTY,
            SUM(CASE WHEN DaysInInventoryAN > 90 THEN InventoryCount ELSE 0 END) AS AgedInv,
            'Ending' AS Period,
            NULL AS InvoiceTotal,
            NULL AS CashPrice,
            NULL AS TotalCorePVR,
            NULL AS TotalBasePVR,
            NULL AS SoldCount,
            NULL AS TotalMSRP,
            NULL AS AccountingSoldCount
        FROM NDDUsers.vInventoryMonthEnd
        WHERE 
            (AccountingMonth = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 1, 0)
            AND AccountingMonth <> DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) + 1, 0)
            AND Department = 300
            AND MarketName NOT IN ('Market 98', 'Market 97'))
            OR (AccountingMonth = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 12, 0)
                AND Department = 300
                AND MarketName NOT IN ('Market 98', 'Market 97'))
            OR (AccountingMonth = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()), 0)
                AND Department = 300
                AND MarketName NOT IN ('Market 98', 'Market 97')
                AND Status <> 'G')
        GROUP BY AccountingMonth, Hyperion, StoreName, Make
        HAVING SUM(InventoryCount) <> 0

        UNION ALL

        -- Second Query: Sales Data
        SELECT 
            AccountingMonth,
            StoreHyperion AS Hyperion,
            StoreName,
            Brand,
            NULL AS InvQTY,
            NULL AS AgedInv,
            'Ending' AS Period,
            SUM(InvoiceTotal) AS InvoiceTotal,
            SUM(CashPrice) AS CashPrice,
            SUM(CorePVR) AS TotalCorePVR,
            SUM(BasePVR) AS TotalBasePVR,
            SUM(SoldCount) AS SoldCount,
            SUM(CASE WHEN SoldCount = 0 THEN 0 ELSE MSRP END) AS TotalMSRP,
            SUM(AccountingSoldCount) AS AccountingSoldCount
        FROM (
            SELECT 
                Q2.*,
                Q1.Brand
            FROM (
                -- First part of sales data
                SELECT
                    Vin,
                    StoreHyperion,
                    AccountingMonth,
                    StoreName,
                    VehicleMakeName,
                    NULL AS SoldCount,
                    SUM(InvoicePrice) AS InvoiceTotal,
                    SUM(CashPrice) AS CashPrice,
                    SUM(FrontGross) AS CorePVR,
                    SUM(Baseretail) AS BasePVR,
                    MAX(MSRP_Adv) AS MSRP,
                    SUM(VehicleSoldCount) AS AccountingSoldCount
                FROM NDDUsers.vSalesDetail_Vehicle
                WHERE 
                    AccountingMonth >= DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 1, 0)
                    AND AccountingMonth <> DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) + 1, 0)
                    AND DepartmentName = 'NEW'
                    AND RecordSource = 'Accounting'
                GROUP BY Vin, StoreHyperion, AccountingMonth, StoreName, VehicleMakeName
                
                UNION ALL
                
                -- Second part of sales data
                SELECT
                    Vin,
                    StoreHyperion,
                    AccountingMonth,
                    StoreName,
                    VehicleMakeName,
                    SUM(VehicleSoldCount) AS SoldCount,
                    NULL AS InvoiceTotal,
                    NULL AS CashPrice,
                    NULL AS CorePVR,
                    NULL AS BasePVR,
                    NULL AS MSRP,
                    NULL AS AccountingSoldCount
                FROM NDDUsers.vSalesDetail_Vehicle
                WHERE 
                    (AccountingMonth >= DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 1, 0)
                    AND AccountingMonth <> DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) + 1, 0)
                    AND DepartmentName = 'NEW')
                    OR (AccountingMonth = DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE()) - 12, 0)
                        AND DepartmentName = 'NEW')
                GROUP BY Vin, StoreHyperion, AccountingMonth, StoreName, VehicleMakeName
            ) Q2
            LEFT JOIN (
                SELECT 
                    HYPERION_ID,
                    CASE 
                        WHEN MAN_NAME IN ('Chevrolet', 'GMC', 'Cadillac') THEN 'GM'
                        WHEN MAN_NAME = 'Dodge' THEN 'Chrysler'
                        WHEN MAN_NAME = 'Lincoln-Mercury' THEN 'Ford'
                        WHEN MAN_NAME = 'Mercedes' THEN 'Mercedes-Benz'
                        WHEN MAN_NAME IN ('Jaguar', 'Land Rover') THEN 'JLR'
                        ELSE MAN_NAME
                    END AS Brand
                FROM NDDUsers.vHyperionDetail
                WHERE STORE_STATUS = 'OPEN'
                    AND REGION IN ('Region 1', 'Region 2')
                    AND ENTITY_NAME NOT LIKE '%Waymo%'
                    AND ENTITY_NAME NOT LIKE '%WBYC%'
                    AND HYP_TYPE = 'Store'
            ) Q1 ON Q1.HYPERION_ID = Q2.StoreHyperion
        ) AS SQ2
        GROUP BY AccountingMonth, StoreHyperion, StoreName, Brand
        HAVING SUM(SoldCount) <> 0
    ) AS SQ
    WHERE Brand NOT IN ('Lamborghini', 'Bentley', 'Aston Martin')
    GROUP BY AccountingMonth, Hyperion, StoreName, Brand;
    """

    try:
        with pyodbc.connect(
                'DRIVER={ODBC Driver 17 for SQL Server};'
                'SERVER=S2WPPSQL03NDD;'
                'DATABASE=NDD_ADP_RAW;'
                'Trusted_Connection=yes;'
        ) as conn:
            NDD_df = pd.read_sql(NDD_query, conn)

    except Exception as e:
        print("❌ Connection failed:", e)

    # Open file and process macro/Sql
    BPU_File = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Brand_President_Tracker\BP_Tracker.xlsm'
    DOC_AND_BUDGET_file = r'W:\Corporate\Inventory\Reporting\Brand President Tracker\DATA SOURCE FILES ESSBASE\DOC AND BUDGET.xlsx'
    app = xw.App(visible=True)
    wb = app.books.open(BPU_File)
    # Need to open up workbook for Macro
    # Doc_and_Budget_wb = app.books.open(DOC_AND_BUDGET_file)

    NDD_tab = wb.sheets['Data']
    NDD_tab.range("A6:M10000").clear_contents()
    NDD_tab.range('A6').options(index=False, header=False).value = NDD_df 

    # Run Macro
    Run_Macro = wb.macro("Execute_VBAs")
    Run_Macro()
    wb.save()

    # Save and close the excel document    
    if wb:
        wb.close()
    if app:
        app.quit()

    """ Prepare email files"""
    # Create the filename
    today = datetime.today().strftime('%m.%d.%y')
    filename = f"BP Tracker {today}.xlsm"
    destination_folder = r'C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Brand_President_Tracker\Reports'
    full_filename = os.path.join(destination_folder, filename)
    # Send copy of full file with todays date
    shutil.copy(BPU_File, full_filename)

    return

def BP_Tracker_Email():
    
    BP_Tracker_File = r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Brand_President_Tracker\BP_Tracker.xlsm"
    app = xw.App(visible=True)
    BP_wb = app.books.open(BP_Tracker_File)
    
    Run_Macro = BP_wb.macro("Create_BP_Tracker_Email")
    Run_Macro()
    
    # wait 10 seconds
    time.sleep(10)

    # Save and close the excel document(s)    
    if BP_wb:
        BP_wb.close()
    if app:
        app.quit()

    return

#run function
if __name__ == '__main__':
    #Process_Daily_Sales_File()
    #Update_DOC_AND_BUDGET_file()
    Update_BPU_File()
    # BP_Tracker_Email()