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

        # Credentials
        username = "besadag"
        password = "T0ttenh@mG-DZ@ndy10!"
        DOC_AND_BUDGET_file_path = r'W:\Corporate\Inventory\Reporting\Brand President Tracker\DATA SOURCE FILES ESSBASE\DOC AND BUDGET.xlsb'

        # Need to open up workbook for Macro
        app = xw.App(visible=True)
        Doc_and_Budget_wb = app.books.open(DOC_AND_BUDGET_file_path)
        # Wait for file to open
        time.sleep(5)

        # Select specific cell (e.g., 'B2') and input Doc_Input
        # Adjust Month, type in current month
        current_month = datetime.now().strftime('%b')
        yesterday = datetime.now() - timedelta(days=1)  # Yesterday's date
        Doc_Input = "DOCFcst" + str(yesterday.day)
        print(Doc_Input)
        sheet = Doc_and_Budget_wb.sheets[0]  # Adjust if not the first sheet
        sheet.range('F3').select()
        sheet.range('F3').value = Doc_Input
        time.sleep(1)
        sheet.range('F8').select()
        sheet.range('F8').value = current_month
        time.sleep(1)
        sheet.range('G8').value = current_month
        time.sleep(1)

        # Bring Excel window to foreground
        excel_windows = [w for w in gw.getWindowsWithTitle('DOC AND BUDGET') if w.visible]
        if excel_windows:
            excel_windows[0].activate()
            time.sleep(1)

        # Click Smart View ribbon button by coordinates (example: x=300, y=120)
        pyautogui.click(2656, 61)
        time.sleep(1)

        # Click Refresh button        
        pyautogui.click(2293, 101)
        time.sleep(1)

        # After refresh, save and close the workbook
        Doc_and_Budget_wb.save()
        Doc_and_Budget_wb.close()
        app.quit()

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
                'SERVER=nddprddb01,48155;'
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