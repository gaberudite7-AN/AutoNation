from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import undetected_chromedriver as uc
import time
from datetime import datetime
import shutil

# Setup Chrome options
chrome_options = uc.ChromeOptions()
chrome_options.add_argument("--headless") # Run without opening window in GUI
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36")

# Path to ChromeDriver
chrome_driver_path = r"C:\Development\Chrome_Driver\chromedriver-win64\chromedriver.exe"
url = "https://na-ftp.urbanscience.com/ThinClient/WTM/public/index.html#/login"


# Get today's date in the format YYYYMMDD
today_str = datetime.today().strftime('%Y%m%d')

# Construct the expected filename
filename = f"AutoNation_SalesFile_NationalSales_{today_str}.txt"

# Launch browser
driver = uc.Chrome(
    driver_executable_path=chrome_driver_path,
    options=chrome_options,
    use_subprocess=True
)

# Open the login page
driver.set_page_load_timeout(20)
driver.get(url)

# Your credentials
username = "Stuartm"
password = "f$w8Q)$z%pt)"

try:
    # Wait for username field and enter username
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "inputUsername"))
    ).send_keys(username)

    # Wait for password field and enter password
    WebDriverWait(driver, 10).until(
        EC.visibility_of_element_located((By.ID, "password"))
    ).send_keys(password)

    # Click the login button
    WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "signIn"))
    ).click()
    print("Current URL:", driver.current_url)
    time.sleep(5)

    # Wait for the file list to load and find the checkbox by its label's "for" attribute
    checkbox = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, f"input[id*='{filename}']"))
    )

    # Click the checkbox
    checkbox.click()
    print(f"Clicked checkbox for file: {filename}")
    time.sleep(5)

    # Click download
    download_button = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'toolbar-button') and .//span[text()='Download']]"))
    )

    # Click the download button
    download_button.click()
    print(f"Download button clicked. Downloaded file {filename} to downloads")
    time.sleep(15)

    # Move file to desination
    # Get today's date
    today = datetime.today()

    # Format as M.D (e.g., 7.3)
    date_str = f"{today.month}.{today.day}"

    # Construct the new filename
    filename_final = f"AutoNation_SalesFile_NationalSales_{date_str}.txt"


    File_Destination = rf'W:\Corporate\Inventory\Urban Science\Historics\{filename_final}'
    shutil.move(rf"C:\Users\BesadaG\Downloads\{filename}", File_Destination)
    print(f"Successfully sent file from downloads to {filename_final}")

except Exception as e:
    print(f"No file found for today ({filename}) or an error occurred: {e}")

finally:
    # Always quit the driver to avoid WinError 6
    driver.quit()