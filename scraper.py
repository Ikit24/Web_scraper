import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import os

headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36'
}

urls = [
    'https://groww.in/us-stocks/ko',
    'https://groww.in/us-stocks/fdp',
    'https://groww.in/us-stocks/aapl',
    'https://groww.in/us-stocks/bti',
    'https://groww.in/us-stocks/amzn',
    'https://groww.in/us-stocks/pm',
    'https://groww.in/us-stocks/nvda'
]

all_data = []
print("Starting to scrape...")
for url in urls:
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
        
        stock_data = [company, price, change, volume]
        all_data.append(stock_data)
        
    except Exception as e:
        pass
    
    # Wait to avoid rate limiting
    time.sleep(10)

if all_data:
    column_names = ["Company", "Price", "Change", "Volume"]
    df = pd.DataFrame(all_data, columns=column_names)

    if os.path.exists('stocks.xlsx'):
        os.remove('stocks.xlsx')
        print("Existing Excel file deleted")
        print("Scraping finished!")

    df.to_excel('stocks.xlsx', index=False)
else:
    pass