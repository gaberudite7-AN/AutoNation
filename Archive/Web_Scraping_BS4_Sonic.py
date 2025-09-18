from bs4 import BeautifulSoup
from urllib3 import PoolManager
from urllib3.util.ssl_ import create_urllib3_context
import requests
import warnings
import pandas as pd
import urllib3
import re
import html

path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Sonic.html'

# Load your raw HTML or JS-like string
with open(path, "r", encoding="utf-8") as f:
    raw_data = f.read()

# Unescape entities
clean_html = html.unescape(raw_data)

# Parse HTML with BeautifulSoup
soup = BeautifulSoup(clean_html, "html.parser")

results = []

# Loop through each store name
for org_tag in soup.find_all("span", class_="org"):
    store_name = org_tag.get_text(strip=True)

    # Find the next instances of locality and region in the document
    locality_tag = org_tag.find_next("span", class_="locality")
    region_tag = org_tag.find_next("span", class_="region")

    city = locality_tag.get_text(strip=True) if locality_tag else ""
    state = region_tag.get_text(strip=True) if region_tag else ""

    results.append({
        "Store Name": store_name,
        "City": city,
        "State": state
    })

# Convert to DataFrame
df = pd.DataFrame(results)

# Save to CSV
df.to_csv(r'C:\Users\BesadaG\OneDrive - AutoNation\Mikael\Urban_Science\Sonic.csv', index=False)
