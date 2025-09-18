from bs4 import BeautifulSoup
from urllib3 import PoolManager
from urllib3.util.ssl_ import create_urllib3_context
import requests
import warnings
import pandas as pd
import urllib3

# Suppress InsecureRequestWarning
warnings.filterwarnings("ignore", message="Unverified HTTPS request")

# Get page HTML (disable SSL verification for testing)
url = 'https://www.asburyauto.com/locations'
response = requests.get(url, verify=False)

# Parse the HTML
soup = BeautifulSoup(response.text, 'html.parser')



# Extract names
names = [tag.text for tag in soup.find_all('h3', class_='text-sm text-gray-700 font-semibold')]
print(names)

# Extract addresses and clean <br> tags
addresses = []
for p in soup.find_all('p', class_='mt-1 text-sm text-gray-500'):
    for br in p.find_all('br'):
        br.replace_with('\n')
    addresses.append(p.text.strip())

print(addresses)

# Create DataFrames
df_names = pd.DataFrame(names, columns=['Name'])
df_addresses = pd.DataFrame(addresses, columns=['Address'])

# Preview
print(df_names)
print(df_addresses)