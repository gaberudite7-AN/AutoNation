from bs4 import BeautifulSoup
from urllib3 import PoolManager
from urllib3.util.ssl_ import create_urllib3_context
import requests
import warnings
import pandas as pd
import urllib3

path = r'C:\Development\.venv\Scripts\Python_Scripts\Web_Scraping\Asbury.html'

# Read local HTML file
with open(path, 'r', encoding='utf-8') as f:
    html = f.read()

# Parse with BeautifulSoup
soup = BeautifulSoup(html, 'html.parser')

# Find all dealership name and address pairs
dealers = []
# Find all h3 tags with the right class (name)
name_tags = soup.find_all('h3', class_='text-sm text-gray-700 font-semibold')

for name_tag in name_tags:
    name = name_tag.text.strip()

    # Find the next sibling <p> with the address
    p_tag = name_tag.find_next_sibling('p', class_='mt-1 text-sm text-gray-500')

    if p_tag:
        # Replace <br> with newline
        for br in p_tag.find_all('br'):
            br.replace_with('\n')
        address = p_tag.text.strip()
    else:
        address = ''

    dealers.append({'Name': name, 'Address': address})

# Convert to DataFrame
df = pd.DataFrame(dealers)

# Display or export
df.to_csv(r'C:\Users\BesadaG\OneDrive - AutoNation\Mikael\Urban_Science\test1.csv')