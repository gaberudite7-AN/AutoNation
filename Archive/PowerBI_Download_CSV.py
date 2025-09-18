from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException
import undetected_chromedriver as uc
import time
from datetime import datetime
import shutil
import os
import traceback
import pyautogui
import pytesseract
import pygetwindow as gw
from pathlib import Path
import glob




def Get_Quarterly_Data():

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
    filename = "powerbi.csv"


    # Target URL
    url = "https://app.powerbi.com/groups/9115e1a0-32f3-4ec6-926c-12a271724a7c/reports/1a5868c8-bf57-489f-b9b5-fa1b37e5769b/e22a07b483d46189bd84?ctid=bd54fbce-74dd-4b5a-8d71-2b978c6d210d&experience=power-bi"

    # Start browser
    try:
        driver = uc.Chrome(
            driver_executable_path=chrome_driver_path,
            options=chrome_options,
            use_subprocess=True
        )
        actions = ActionChains(driver)
        driver.set_page_load_timeout(20)
        driver.get(url)

        # Define wait AFTER driver is initialized
        wait = WebDriverWait(driver, 20)
        time.sleep(5)

        Email = "besadag@autonation.com"

        time.sleep(1)
        
        # Step 1: Enter email
        email_input = wait.until(EC.presence_of_element_located((By.ID, "email")))
        email_input.send_keys(Email)
        time.sleep(2)

        # Click the Submit button
        submit_button = wait.until(EC.element_to_be_clickable((By.ID, "submitBtn")))
        submit_button.click()
        time.sleep(3)

        # Click the "Continue" button after entering email or password
        continue_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        continue_button.click()
        time.sleep(3)

        # Click the "Yes" button on the "Stay signed in?" prompt
        yes_button = wait.until(EC.element_to_be_clickable((By.ID, "idSIButton9")))
        yes_button.click()
        time.sleep(3)

        # Click Segment
        Segment_button = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//span[text()='Segment']")
        ))
        Segment_button.click()

        print(f"Segment Clicked")
        time.sleep(2)

        # XPath of the visual element to click
        xpath = '//h3[@class="slicer-header-text" and @title="Month_"]'
        element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        element.click()
        print("Clicked the Month_ header element")
        time.sleep(2)

        clear_buttons = driver.find_elements(By.CSS_SELECTOR, '[aria-label="Clear selections"]')
        visible_button = None
        for btn in clear_buttons:
            if btn.is_displayed():
                visible_button = btn
                break

        if visible_button:
            visible_button.click()
            print("Clicked the visible 'Clear selections' button in Month_!")
            time.sleep(2)
        else:
            print("No visible 'Clear selections' button found.")

        # XPath of the visual element to click
        xpath = '//h3[@class="slicer-header-text" and @title="Quarter"]'
        element = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        element.click()
        print("Clicked the Quarter header element")
        time.sleep(2)

        clear_buttons = driver.find_elements(By.CSS_SELECTOR, '[aria-label="Clear selections"]')
        visible_button = None
        for btn in clear_buttons:
            if btn.is_displayed():
                visible_button = btn
                break

        if visible_button:
            visible_button.click()
            print("Clicked the visible 'Clear selections' button in Quarter!")
            time.sleep(2)
        else:
            print("No visible 'Clear selections' button found.")

        """Begin exporting by quarter"""
        quarters = ['Select all', '1', '2', '3', '4']

        for quarter in quarters:
            print(f"Exporting quarter {quarter}")
            found = False
            dropdown = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div.slicer-dropdown-menu[aria-label='Quarter']")))
            dropdown.click()
            time.sleep(1)

            try: 
                # Wait for the dropdown options to appear and select the one with text "1"
                option = wait.until(EC.element_to_be_clickable((By.XPATH, f"//div[@class='slicerItemContainer' and @title='{quarter}']")))
                option.click()
                time.sleep(2)
                print("Clicked quarter we need")
                found = True

            except TimeoutException:
                print(f"❌ Quarter '{quarter}' not found in dropdown. Skipping.")
                continue

            # Get all visual containers (this may vary slightly based on your Power BI structure)
            actions = ActionChains(driver)
            target_keywords = ["AN SEGMENT"]
            visuals = driver.find_elements(By.CSS_SELECTOR, '[data-testid="visual-container"]')


            actions = ActionChains(driver)
            target_visual = None

            for idx, visual in enumerate(visuals):
                try:
                    visual_text = visual.text

                    if any(keyword in visual_text for keyword in target_keywords):
                        print(f"✅ Target visual found at index {idx}")

                        # Scroll to it and click (this helps trigger internal hover/activate behavior)
                        # Move mouse to it (hover)
                        ActionChains(driver).move_to_element(visual).pause(1).perform()
                        time.sleep(2)
                        # Click it separately (optional if hover is enough)
                        ActionChains(driver).move_to_element(visual).click().perform()
                        time.sleep(2)
                        # Once Clicked
                        wait = WebDriverWait(driver, 20)
                        more_options_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@aria-label='More options']")))
                        more_options_btn.click()
                        time.sleep(2)
                        print("Clicked more options")
                        break

                except Exception as e:
                    print(f"Visual {idx} failed: {e}")

            # Wait for the "Export data" option to appear and click it
            export_data_option = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[text()='Export data']")
            ))
            export_data_option.click()
            print("Clicked export data from options")
            time.sleep(2)

            # Wait for the Summarized data radio option and click it
            summarized_radio = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Summarized data')]"))
            )
            summarized_radio.click()
            print("Clicked 'Summarized data' option.")
            time.sleep(2)

            # STEP 1: Click the dropdown to open the export format options
            dropdown_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="pbi-dropdown"]'))
            )
            dropdown_button.click()
            print("Clicked export format dropdown.")

            # STEP 2: Wait for and click the .csv option
            csv_option = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), '.csv (30,000-row max)')]"))
            )
            csv_option.click()
            print("Selected .csv export format.")

            # Click the Export button
            export_button = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-testid="export-btn"]'))
            )
            export_button.click()
            print(f"✅ Export initiated for quarter {quarter}")
            
            print("Resetting the visual")
            # XPath of the visual element to click
            xpath = '//h3[@class="slicer-header-text" and @title="Quarter"]'
            element = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            element.click()
            print("Clicked the Quarter header element")
            time.sleep(2)

            clear_buttons = driver.find_elements(By.CSS_SELECTOR, '[aria-label="Clear selections"]')
            visible_button = None
            for btn in clear_buttons:
                if btn.is_displayed():
                    visible_button = btn
                    break

            if visible_button:
                visible_button.click()
                print("Clicked the visible 'Clear selections' button in Quarter!")
                time.sleep(2)
            else:
                print("No visible 'Clear selections' button found.")
            time.sleep(2)
            print("Done. Moving on to next quarter")

    finally:
        print("Completed exporting all the quarterly data from the visual")
        # Patch the __del__ method to suppress the OSError
        def safe_del(self):
            try:
                self.quit()
            except Exception:
                pass  # Silently ignore all errors
        uc.Chrome.__del__ = safe_del    
    return

def Move_Quarterly_Data():

    # Get path to user's Downloads folder
    downloads_path = str(Path.home() / "Downloads")
    target_folder = r"C:\Users\BesadaG\OneDrive - AutoNation\PowerAutomate\Market_Share\Quarter_Data_from_PowerBI"

    # Find all CSV files in Downloads containing "data" in the filename
    csv_files = glob.glob(os.path.join(downloads_path, "*data*.csv"))

    # Sort them by modified time (oldest first so quarter1 is the oldest)
    csv_files.sort(key=os.path.getmtime)

    # Rename and copy them to the target folder
    for i, file_path in enumerate(csv_files, start=0):
        dest_path = os.path.join(target_folder, f"quarter{i}.csv")
        shutil.copy(file_path, dest_path)
        print(f"✅ Copied: {file_path} → {dest_path}")

    return

def Update_Quarterly_Report():

    return  

#run function
if __name__ == '__main__':
    
    #Get_Quarterly_Data()
    #print("Done scraping quarterly data from PowerBI, moving files to desired destination")
    #time.sleep(2)
    #Move_Quarterly_Data()
    Update_Quarterly_Report()