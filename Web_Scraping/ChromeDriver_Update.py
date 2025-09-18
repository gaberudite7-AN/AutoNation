import os
import zipfile
import shutil
import subprocess
import sys
import re
import tempfile
from urllib.request import urlopen
from urllib.error import URLError
from io import BytesIO

def get_chrome_version():
    try:
        output = subprocess.check_output(
            r'reg query "HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon" /v version',
            shell=True,
            stderr=subprocess.DEVNULL,
            stdin=subprocess.DEVNULL
        ).decode()
        version = re.search(r"version\s+REG_SZ\s+([^\s]+)", output)
        if version:
            return version.group(1)
    except Exception as e:
        print("Failed to get Chrome version:", e)
    return None

def download_chromedriver(version, dest_dir):
    base_url = f"https://edgedl.me.gvt1.com/edgedl/chrome/chrome-for-testing/{version}/win64/chromedriver-win64.zip"
    try:
        print(f"Downloading ChromeDriver version {version}...")
        response = urlopen(base_url)
        with zipfile.ZipFile(BytesIO(response.read())) as zip_ref:
            zip_ref.extractall(dest_dir)
        print("Download and extraction complete.")
        return True
    except URLError as e:
        print(f"Failed to download ChromeDriver: {e}")
    return False

def replace_chromedriver(new_driver_path, target_path):
    try:
        if os.path.exists(target_path):
            os.remove(target_path)
        shutil.copy(new_driver_path, target_path)
        print(f"ChromeDriver updated at: {target_path}")
    except Exception as e:
        print(f"Failed to replace ChromeDriver: {e}")

def verify_chromedriver_version(driver_path):
    try:
        output = subprocess.check_output([driver_path, "--version"]).decode()
        print("Installed ChromeDriver version:", output.strip())
    except Exception as e:
        print("Failed to verify ChromeDriver version:", e)

# Main logic
chrome_version = get_chrome_version()
if not chrome_version:
    sys.exit("Could not determine installed Chrome version.")

major_version = ".".join(chrome_version.split(".")[:4])
print("Detected Chrome version:", major_version)

# Define paths
target_driver_path = r"C:\Development\Chrome_Driver\chromedriver-win64\chromedriver.exe"
with tempfile.TemporaryDirectory() as tmpdir:
    if download_chromedriver(major_version, tmpdir):
        new_driver = os.path.join(tmpdir, "chromedriver-win64", "chromedriver.exe")
        if os.path.exists(new_driver):
            replace_chromedriver(new_driver, target_driver_path)
            verify_chromedriver_version(target_driver_path)
        else:
            print("Downloaded ChromeDriver executable not found.")
    else:
        print("ChromeDriver update failed.")