from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc
import time
from datetime import datetime
import shutil
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.action_chains import ActionChains


# Setup Chrome options
chrome_options = uc.ChromeOptions()
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36")

# Helper to get current and previous quarter
def get_quarter_options():
    now = datetime.now()
    year = now.year
    month = now.month
    # Calculate current quarter (1-4)
    current_quarter = (month - 1) // 3 + 1
    # Previous quarter logic
    if current_quarter == 1:
        prev_quarter = 4
        prev_year = year - 1
    else:
        prev_quarter = current_quarter - 1
        prev_year = year
    return [
        f"Qtr {prev_quarter}, {prev_year}",
        f"Qtr {current_quarter}, {year}"
    ]

def subaru_update():
    # Path to ChromeDriver
    chrome_driver_path = r"C:\Development\Chrome_Driver\chromedriver-win64\chromedriver.exe"
    url = "https://subarunet.com"


    # Get today's date in the format YYYYMMDD
    today_str = datetime.today().strftime('%Y%m%d')

    # Launch browser
    driver = uc.Chrome(
        driver_executable_path=chrome_driver_path,
        options=chrome_options,
        use_subprocess=True
    )

    # Open the login page
    driver.set_page_load_timeout(20)
    driver.get(url)

    # Override website detection
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")


    # Your credentials
    username = "mpiz3500"
    password = "Auto!Nation40"

    try:
        time.sleep(2)
        # Wait for username field and enter username

        username_field = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.NAME, "username"))
        )
        for char in username:
            username_field.send_keys(char)
            time.sleep(0.2)
        time.sleep(1)

        # Wait for password field and enter password
        password_field = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.NAME, "password"))
        )
        for char in password:
            password_field.send_keys(char)
            time.sleep(0.2)

        # Click the login button
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.NAME, "submit")) # Adjust selector as needed
        )
        ActionChains(driver).move_to_element(login_button).click().perform()

        print("Current URL:", driver.current_url)
        time.sleep(5)

        # 1. Click the store dropdown
        store_dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='app']/div[1]/header/div/button[3]"))
        )
        store_dropdown.click()

        time.sleep(2)
        
        store_element = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[@data-cy='item-AutoNation Subaru Scottsdale (090597)']"))
        )
        store_element.click()


        time.sleep(2)

        # 2. Expand the "OLP/Customer Commitment Award" section
        owner_loyalty_program_section = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'v-list-group__header')]//div[contains(text(), 'OLP/Customer Commitment Award')]"))
        )
        owner_loyalty_program_section.click()
        time.sleep(2)

        # 3. Click the "Owner Loyalty Program (OLP)" link
        owner_loyalty_program_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(., 'Owner Loyalty Program (OLP)')]"))
        )
        owner_loyalty_program_link.click()
        time.sleep(2)

        # New website activated
        # Switch to the newest tab
        driver.switch_to.window(driver.window_handles[-1])

        # 4. Click Purchase QTD
        Purchase_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_Main_hlPurchaseDetail']/div[1]/div"))
        )
        Purchase_link.click()
        time.sleep(2)


        # 1. Hover over the "Net Promoter Score" menu item
        Net_Promoter_Score_hover = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//*[@id='ctl00_Main_RadMenu1']/ul/li[3]/a"))
        )

        ActionChains(driver).move_to_element(Net_Promoter_Score_hover).perform()
        time.sleep(2) # Allow submenu to appear

        Scores_click = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//span[@class='rmText' and text()='NPS®, KSI & SQI Scores-F&I Manager']"))
        )
        Scores_click.click()

        time.sleep(5)
        # New website activated
        # Switch to the newest tab
        driver.switch_to.window(driver.window_handles[-1])

        # Adjust Quarter Start and Quarter End Filters
        # 1. Purchase loop
        previous_quarter = get_quarter_options()[0]  # First item is previous quarter
        current_quarter = get_quarter_options()[1]   # Second item is current quarter

        # Adjust Quarter Start and Quarter End Filters
        previous_quarter = get_quarter_options()[0]  # First item is previous quarter

        # Get the dropdown and create the Select object
        dropdown = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_Main_ReportViewer1_ctl08_ctl07_ddValue']"))
        )
        select = Select(dropdown)
        found = False
        for option in select.options:
            clean_option = option.text.replace('\xa0', ' ').replace(',', '').strip()
            clean_target = previous_quarter.replace(',', '').strip()
            if clean_target in clean_option:
                select.select_by_visible_text(option.text)
                print(f"Selected: {option.text}")
                found = True
                time.sleep(1)  # Wait for any postback or page update
                break
        if not found:
            print(f"Could not find option for: {previous_quarter}")

        time.sleep(2)

       # Select current quarter in the second dropdown
        current_quarter = get_quarter_options()[1]
        dropdown_curr = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='ctl00_Main_ReportViewer1_ctl08_ctl09_ddValue']"))
        )
        select_curr = Select(dropdown_curr)
        found_curr = False
        for option in select_curr.options:
            clean_option = option.text.replace('\xa0', ' ').replace(',', '').strip()
            clean_target = current_quarter.replace(',', '').strip()
            if clean_target in clean_option:
                select_curr.select_by_visible_text(option.text)
                print(f"Selected current quarter: {option.text}")
                found_curr = True
                time.sleep(1)
                break
        if not found_curr:
            print(f"Could not find option for current quarter: {current_quarter}")

        # Click the expand button for "AutoNation Subaru Carlsbad"
        expand_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//img[contains(@src, 'TogglePlus.gif') and @alt='Expand AutoNation Subaru Scottsdale']"
            ))
        )
        expand_button.click()
        time.sleep(2)

        # 1. Find all month elements (e.g., "September 2025")
        month_elements = driver.find_elements(
            By.XPATH,
            "//div[contains(@class, 'canGrowTextBoxInTablix') and contains(@id, '_aria') and contains(., '2025')]"
        )

        # Get the last 3 months (assuming they are in order)
        last_3_months = month_elements[-3:]

        for month_elem in last_3_months:
            month_text = month_elem.text.strip()
            # Find the parent row (tr) of the month cell
            parent_row = month_elem.find_element(By.XPATH, "./ancestor::tr[1]")

            # Extract Returns value (first <td> with the specific class in the row)
            returns_td = parent_row.find_element(
                By.XPATH,
                ".//td[contains(@class, 'Abf036ccf0e5c4342b0991681657255b3347c')]//div[contains(@class, 'canGrowTextBoxInTablix')]"
            )
            returns_value = returns_td.text.strip()

            # Extract NPS score (first <td> with the specific class in the row)
            nps_td = parent_row.find_element(
                By.XPATH,
                ".//td[contains(@class, 'Abf036ccf0e5c4342b0991681657255b3351c')]//div[contains(@class, 'canGrowTextBoxInTablix')]"
            )
            nps_value = nps_td.text.strip()

            print(f"Month: {month_text}, Returns: {returns_value}, NPS: {nps_value}")

    finally:
        if 'driver' in locals():
            def safe_del(self):
                try:
                    self.quit()
                except Exception:
                    pass  # Silently ignore all errors
            uc.Chrome.__del__ = safe_del

    return

if __name__ == "__main__":
    subaru_update()