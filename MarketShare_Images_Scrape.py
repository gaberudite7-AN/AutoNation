from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import undetected_chromedriver as uc
import time
import shutil
import os
from selenium.webdriver.common.action_chains import ActionChains


def export_markets_create_screenshot(driver, wait, page_label):
    # Click the segment tab
    Page_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, f"//span[text()='{page_label}']")
    ))
    Page_button.click()
    print(f"{page_label} clicked")
    time.sleep(2)

    # Define markets
    Markets = [
        'MK01 - Southern CA', 'MK02 - Northern CA & NV', 'MK03 - WA & AZ',
        'MK04 - CO & North TX', 'MK05 - South TX', 'MK06 - Midwest & Northeast',
        'MK07 - Southeast', 'MK08 - North-Central Fl', 'MK09 - South FL'
    ]

    for market in Markets:
        print(f"Exporting market {market}")
        try:
            # Open dropdown
            dropdown_toggle = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@class, 'slicer-dropdown') and not(contains(@class, 'slicer-dropdown-popup'))]")
            ))
            dropdown_toggle.click()

            # Wait for dropdown popup
            wait.until(EC.visibility_of_element_located(
                (By.XPATH, "//div[contains(@class, 'slicer-dropdown-popup') and contains(@class, 'visual')]")
            ))

            # Click target market
            slicer_item = wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//div[@class='slicerItemContainer' and @title='{market}']")
            ))
            slicer_item.click()

            # Close dropdown
            dropdown_toggle = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//div[contains(@class, 'slicer-dropdown') and not(contains(@class, 'slicer-dropdown-popup'))]")
            ))
            dropdown_toggle.click()

            time.sleep(2)
                                
            image_path = fr"W:\Corporate\Inventory\Urban Science\Excel_Update\Images\Weekly_Page1\{market}_visual.png"
            pdf_path = fr"W:\Corporate\Inventory\Urban Science\Excel_Update\Images\Weekly_Page1\{market}_visual.pdf"

            driver.save_screenshot(image_path)
            print("Saved Screenshot")

        except Exception as e:
            print(f"Could not find element for market {market}: {e}")
    
    return

def export_markets_to_pdf(driver, wait, page_label):
    # Click the segment tab
    Page_button = wait.until(EC.element_to_be_clickable(
        (By.XPATH, f"//span[text()='{page_label}']")
    ))
    Page_button.click()
    print(f"{page_label} clicked")
    time.sleep(2)

    # Define markets
    Markets = [
        'MK01 - Southern CA', 'MK02 - Northern CA & NV', 'MK03 - WA & AZ',
        'MK04 - CO & North TX', 'MK05 - South TX', 'MK06 - Midwest & Northeast',
        'MK07 - Southeast', 'MK08 - North-Central Fl', 'MK09 - South FL'
    ]

    for market in Markets:
        print(f"Exporting market {market}")
        try:
            label_text = "AN Market"
            
            # Locate the dropdown toggle for AN Market
            dropdown_toggle = wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//div[contains(@class,'slicer-dropdown') "
                        f"and not(contains(@class,'slicer-dropdown-popup')) "
                        f"and @aria-label='{label_text}']"))
            )

            # Get the popup id from the toggle's aria-controls
            popup_id = dropdown_toggle.get_attribute("aria-controls")
            dropdown_toggle.click()

            # Wait for the popup to appear
            popup = wait.until(EC.visibility_of_element_located((By.ID, popup_id)))

            # Click the desired option inside the AN Market popup
            slicer_item = wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//div[@id='{popup_id}']//div[@class='slicerItemContainer' and @title='{market}']"))
            )
            slicer_item.click()

            # Locate the dropdown toggle for AN Market
            dropdown_toggle = wait.until(EC.element_to_be_clickable(
                (By.XPATH, f"//div[contains(@class,'slicer-dropdown') "
                        f"and not(contains(@class,'slicer-dropdown-popup')) "
                        f"and @aria-label='{label_text}']"))
            )

            # Click again to close toggle
            popup_id = dropdown_toggle.get_attribute("aria-controls")
            dropdown_toggle.click()

            time.sleep(1)

            # Click export button
            export_button = wait.until(EC.element_to_be_clickable((By.ID, "exportMenuBtn")))
            export_button.click()
            time.sleep(1)
            
            # Click the 3rd option in the export menu
            export_option = wait.until(EC.element_to_be_clickable(
                (By.XPATH, "//*[@id='mat-menu-panel-2']/div/button[3]")
            ))
            export_option.click()
            time.sleep(1)

            # quick debug so you can see whether the dialog exists at all
            # dlg_ids = [e.get_attribute("id") for e in driver.find_elements(
            #     By.XPATH, "//*[@id and starts-with(@id,'mat-mdc-dialog-')]"
            # )]
            # print("Dialogs in DOM:", dlg_ids)
            # time.sleep(1)

            # click the actual square of the 2nd checkbox inside *the current dialog*
            checkbox_square = wait.until(EC.element_to_be_clickable((
                By.XPATH,
                "//*[@id and starts-with(@id,'mat-mdc-dialog-')]//pbi-checkbox[2]"
                "//div[@data-testid='pbi-checkbox-checkbox']"
            )))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", checkbox_square)
            checkbox_square.click()
            time.sleep(1)

            print("Clicked the 'Only export current page' checkbox.")
            time.sleep(2)                        
            # 3. Wait for the confirmation dialog and click OK
            ok_button = wait.until(EC.element_to_be_clickable((By.ID, "okButton")))
            ok_button.click()
            # wait for file to download
            time.sleep(45)

            # shutil to move download file to correct path.
            # Move latest file to destination folder with adjusted name
            original_filename = "MarketShare_Weekly_Summaries.pdf"
            new_filename = fr"{market}_visual.pdf"
            if page_label == "Weekly Market Summary":
                destination_folder = fr"W:\Corporate\Inventory\Urban Science\Excel_Update\Images\Weekly_Page1"
            else:
                destination_folder = fr"W:\Corporate\Inventory\Urban Science\Excel_Update\Images\Weekly_Page2"
            
            downloads_folder = r"C:\Users\BesadaG\Downloads"
            
            source_file = os.path.join(downloads_folder, original_filename)
            destination_file = os.path.join(destination_folder, new_filename)

            shutil.move(source_file, destination_file)
            print("Finished moving downloaded file to destination")
            print("Done. Moving on to next market")
            time.sleep(2)

        except Exception as e:
            print(f"Could not find element for market {market}: {e}")
    
    return


def Create_Images():

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


    # Target URL
    url = "https://app.powerbi.com/groups/9115e1a0-32f3-4ec6-926c-12a271724a7c/reports/d0a39c40-4918-402f-9ecf-6354b7552673/d97b03ba51a46550a4b7?ctid=bd54fbce-74dd-4b5a-8d71-2b978c6d210d&experience=power-bi"

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


        # Functions to get images from Weekly Pages 1 and Page 2
        export_markets_to_pdf(driver, wait, "Weekly Market Summary PG2")
        export_markets_to_pdf(driver, wait, "Weekly Market Summary")

    finally:
        print("Completed exporting all the Market data from the visual")
        # Patch the __del__ method to suppress the OSError
        def safe_del(self):
            try:
                self.quit()
            except Exception:
                pass  # Silently ignore all errors
        uc.Chrome.__del__ = safe_del    
    return



#run function
if __name__ == '__main__':
    start_time = time.time()
    Create_Images()
    # End Timer
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Script completed in {elapsed_time:.2f} seconds")