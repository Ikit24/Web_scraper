import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import os

def extract_basic_info(ticker):
    headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36'
}

    url = f'https://groww.in/us-stocks/{ticker}'

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

        PB_ratio = [
            soup.find('td', string='P/B Ratio').find_next_sibling('td'),
            soup.find('td', {'class': 'ustf141Value contentPrimary bodyLargeHeavy right-align'})
        ]
        PB_ratio = next((elem.text.strip() for elem in PB_ratio if elem), "N/A")

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

        book_value = [
            soup.find('td', string='Book Value').find_next_sibling('td'),
            soup.find('td', {'class': 'ustf141Value contentPrimary bodyLargeHeavy right-align'})
        ]
        book_value = next((elem.text.strip() for elem in book_value if elem), "N/A")
    
        basic_data = [company, price, change, market_cap, volume, PE_ratio, PB_ratio, EPS_ratio, Div_yield, book_value]
        return [basic_data]
    
    except Exception as e:
        print(f"An error occurred: {e}")

def extract_detailed_info(ticker):
    headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36'
}
    url = f'https://groww.in/us-stocks/{ticker}/company-financial'

    try:
        page = requests.get(url, headers=headers)
        page.raise_for_status()
        soup = BeautifulSoup(page.text, 'html.parser')

        def get_latest_value(label):
            label_element = soup.find('div', string=label)
            if not label_element:
                return "N/A"
            
            value_divs = label_element.find_next_siblings('div', class_='tr12Col')
            
            if not value_divs:
                return "N/A"
            
            latest_value_div = value_divs[-1]
            value = latest_value_div.find('div', class_='cf223RowHead')
            return value.text.strip() if value else "N/A"

        revenues = get_latest_value('Revenues')
        cost_of_revenue = get_latest_value('Cost of Revenue')
        gnrl_admin_expenses = get_latest_value('General & Administrative Expenses')
        operating_expenses = get_latest_value('Operating Expenses')
        interest_expenses = get_latest_value('Interest Expense')
        depreciation_and_amortization = get_latest_value('Depreciation, Amortization & Accretion')
        ebitda = get_latest_value('Earnings & Depreciation Amortization (EBITDA)')
        gross_profit = get_latest_value('Gross Profit')
        net_income = get_latest_value('Net Income')
        weighted_avrg_shares = get_latest_value('Weighted Average Shares')
        oprting_income = get_latest_value('Operating Income')

        return [
            [revenues, cost_of_revenue, gnrl_admin_expenses, operating_expenses,
            interest_expenses, depreciation_and_amortization, ebitda, gross_profit,
            net_income, weighted_avrg_shares, oprting_income]
            ]

    except Exception as e:
        print(f"An error occurred: {e}")
        return ["N/A", "N/A"]
        
def export_to_excel(basic_data, detailed_data, filename='stocks.xlsx'):
    if basic_data is not None and detailed_data is not None:
        column_names1 = [
            "Company", "Price", "Change", "Market Cap", "Volume", "P_E ratio", "P_B ratio",
            "EPS(TTM)", "Div. Yield", "Book Value"
            ]
        column_names2 = [
            "Revenues", "Cost of Revenues",
            "General & Administrative Expenses in USD millions",
            "Operating Expenses in USD millions", "Interest Expense in USD millions",
            "Depreciation, Amortization & Accretion in USD millions",
            "EBITDA", "Gross Profit", "Net Income", "Weighted Average Shares", "Operating Income"
            ]
        
        df1 = pd.DataFrame(basic_data, columns=column_names1)
        df2 = pd.DataFrame(detailed_data, columns=column_names2)

        if os.path.exists(filename):
            os.remove(filename)
            print("Existing Excel file deleted")

        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name='Sheet1', index=False, startrow=0)
            
            startrow_df2 = len(df1) + 3
            df2.to_excel(writer, sheet_name='Sheet1', index=False, startrow=startrow_df2)
        
        print("Scraping finished and data exported!")
    else:
        print("No data scraped")

def main():
    ticker = input("Please paste your desired ticker here: ").upper()
    basic_data = extract_basic_info(ticker)
    detailed_data = extract_detailed_info(ticker)
    export_to_excel(basic_data, detailed_data)

if __name__ == '__main__':
    main()
