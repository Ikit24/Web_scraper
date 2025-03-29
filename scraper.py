import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import os

headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36'
}
input_ticker = input("Please paste your desired ticker here: ").upper()

url = f'https://groww.in/us-stocks/{input_ticker}'

all_data = []
print("Starting to scrape...")

try:
    page = requests.get(url, headers=headers)
    soup = BeautifulSoup(page.text, 'html.parser')
   
    company_options = [
        soup.find('h1', {'class': 'usph14Head displaySmall'}),
        soup.find('h1', {'class': 'displaySmall lh28 fontSize28'}),
        soup.find('h1')  # Fallback to any h1
    ]
    company = next((elem.text.strip() for elem in company_options if elem), "Unknown Company")
   
    price_options = [
        soup.find('span', {'class': 'uht141Pri contentPrimary displayBase'}),
        soup.find('span', {'class': 'fontSize32 fontWeightBold'}),
        soup.find('div', {'class': 'lh36'})
    ]
    price = next((elem.text.strip() for elem in price_options if elem), "N/A")
   
    change_options = [
        soup.find('div', {'class': 'uht141Day bodyBaseHeavy contentNegative'}),
        soup.find('div', {'class': 'uht141Day bodyBaseHeavy contentPositive'}),
        soup.find('span', {'class': 'bodyNormal'})
    ]
    change = next((elem.text.strip() for elem in change_options if elem), "N/A")
   
    volume = "N/A"
    tables = soup.find_all('table')
    for table in tables:
        if table.find('th', string=lambda t: t and 'Volume' in t):
            volume_cell = table.find('td', string=lambda t: t and any(c.isdigit() for c in t))
            if volume_cell:
                volume = volume_cell.text.strip()
                break

    market_cap = [
        soup.find('td', string='Market Cap').find_next_sibling('td'),
        soup.find('td', {'class': 'ustf141Value contentPrimary bodyLargeHeavy right-align'})
    ]
    market_cap = next((elem.text.strip() for elem in market_cap if elem), "N/A")

    PE_ratio = [
        soup.find('td', string='P/E Ratio').find_next_sibling('td'),
        soup.find('td', {'class': 'ustf141Value contentPrimary bodyLargeHeavy right-align'})
    ]
    PE_ratio = next((elem.text.strip() for elem in PE_ratio if elem), "N/A")

    EPS_ratio = [
        soup.find('td', string='EPS(TTM)').find_next_sibling('td'),
        soup.find('td', {'class': 'ustf141Value contentPrimary bodyLargeHeavy right-align'})
    ]
    EPS_ratio = next((elem.text.strip() for elem in EPS_ratio if elem), "N/A")

    Div_yield = [
        soup.find('td', string='Div. Yield').find_next_sibling('td'),
        soup.find('td', {'class': 'ustf141Value contentPrimary bodyLargeHeavy right-align'})
    ]
    Div_yield = next((elem.text.strip() for elem in Div_yield if elem), "N/A")
   
    stock_data = [company, price, change, market_cap, volume, PE_ratio, EPS_ratio, Div_yield]
    all_data.append(stock_data)
   
except Exception as e:
    print(f"An error occurred: {e}")

if all_data:
    column_names = ["Company", "Price", "Change", "Market Cap", "Volume", "P_E", "EPS(TTM)", "Div. Yield"]
    df = pd.DataFrame(all_data, columns=column_names)
    if os.path.exists('stocks.xlsx'):
        os.remove('stocks.xlsx')
        print("Existing Excel file deleted")
    df.to_excel('stocks.xlsx', index=False)
    print("Scraping finished!")
else:
    print("No data scraped")