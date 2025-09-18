from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc
import time
from datetime import datetime
import shutil
from selenium.webdriver.common.action_chains import ActionChains


# Setup Chrome options
chrome_options = uc.ChromeOptions()
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--incognito")
chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36")

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
password = "Auto!Nation38"

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
        EC.element_to_be_clickable((By.XPATH, "<span class='rmText'>NPS®, KSI &amp; SQI Scores</span>"))
    )
    Scores_click.click()

    time.sleep(5)
    time.sleep(5)

finally:
    # Always quit the driver to avoid WinError 6
    driver.quit()