from bs4 import BeautifulSoup
from urllib3 import PoolManager
from urllib3.util.ssl_ import create_urllib3_context
import requests
import warnings
import pandas as pd
import urllib3
import re
import html

path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Carmax.html'

# Load your raw HTML or JS-like string
with open(path, "r", encoding="utf-8") as f:
    raw_data = f.read()

# Unescape entities
clean_html = html.unescape(raw_data)

# Regex to match all {"id":"...","name":"...","lat":...,"city":"...","state":"..."} objects
pattern = re.compile(r'{"id":"\d+","name":"(.*?)".*?"city":"(.*?)","state":"(.*?)"}', re.DOTALL)

# Extract matches
matches = pattern.findall(raw_data)

# Convert to DataFrame
df = pd.DataFrame(matches, columns=["Store Name", "City", "State"])

# Save to CSV
df.to_csv(r'C:\Users\BesadaG\OneDrive - AutoNation\Mikael\Urban_Science\Carmax.csv', index=False)
