from bs4 import BeautifulSoup
from urllib3 import PoolManager
from urllib3.util.ssl_ import create_urllib3_context
import requests
import warnings
import pandas as pd
import urllib3
import re
import html

path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Penske.html'

# Load your raw HTML or JS-like string
with open(path, "r", encoding="utf-8") as f:
    raw_data = f.read()

# Regex to extract HTML chunks inside locations.push([...])
pattern = re.compile(r"locations\.push\(\[\s*'(<div class=\"map-popup\">.*?</div>)'", re.DOTALL)
html_blocks = pattern.findall(raw_data)

results = []

for block in html_blocks:
    # Unescape any HTML entities
    clean_html = html.unescape(block)

    # Parse with BeautifulSoup
    soup = BeautifulSoup(clean_html, "html.parser")

    name_tag = soup.find("h3")
    address_tag = soup.find("p", class_="address")

    if name_tag and address_tag:
        name = name_tag.get_text(strip=True)

        # Replace <br/> with commas for easier parsing
        for br in address_tag.find_all("br"):
            br.replace_with(", ")

        address_text = address_tag.get_text(separator=" ", strip=True)
        lines = [line.strip() for line in address_text.split(",") if line.strip()]

        # Assemble city, state ZIP correctly
        if len(lines) >= 3:
            street = ", ".join(lines[:-2])  # all lines except last two
            city = lines[-2]
            state_zip = lines[-1]
            full_address = f"{street}, {city}, {state_zip}"
        else:
            full_address = ", ".join(lines)

        results.append({
            "Name": name,
            "Full Address": full_address
        })

# Convert to DataFrame
df = pd.DataFrame(results)

# Display output
df.to_csv(r'C:\Users\BesadaG\OneDrive - AutoNation\Mikael\Urban_Science\Penske.csv')