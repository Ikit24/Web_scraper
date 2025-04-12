import requests
from bs4 import BeautifulSoup
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time
import logging

def extract_basic_info(ticker):
    headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36'
}

    url = f'https://groww.in/us-stocks/{ticker}'

    print('Starting to scrape...')

    try:
        page = requests.get(url, headers=headers)
        soup = BeautifulSoup(page.text, 'html.parser')
    
        company_options = [
            soup.find('h1', {'class': 'usph14Head displaySmall'}),
            soup.find('h1', {'class': 'displaySmall lh28 fontSize28'}),
            soup.find('h1')  # Fallback to any h1
        ]
        company = next((elem.text.strip() for elem in company_options if elem), 'Unknown Company')
    
        price_options = [
            soup.find('span', {'class': 'uht141Pri contentPrimary displayBase'}),
            soup.find('span', {'class': 'fontSize32 fontWeightBold'}),
            soup.find('div', {'class': 'lh36'})
        ]
        price = next((elem.text.strip() for elem in price_options if elem), 'N/A')
    
        change_options = [
            soup.find('div', {'class': 'uht141Day bodyBaseHeavy contentNegative'}),
            soup.find('div', {'class': 'uht141Day bodyBaseHeavy contentPositive'}),
            soup.find('span', {'class': 'bodyNormal'})
        ]
        change = next((elem.text.strip() for elem in change_options if elem), 'N/A')
    
        volume = 'N/A'
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
        market_cap = next((elem.text.strip() for elem in market_cap if elem), 'N/A')

        PE_ratio = [
            soup.find('td', string='P/E Ratio').find_next_sibling('td'),
            soup.find('td', {'class': 'ustf141Value contentPrimary bodyLargeHeavy right-align'})
        ]
        PE_ratio = next((elem.text.strip() for elem in PE_ratio if elem), 'N/A')

        PB_ratio = [
            soup.find('td', string='P/B Ratio').find_next_sibling('td'),
            soup.find('td', {'class': 'ustf141Value contentPrimary bodyLargeHeavy right-align'})
        ]
        PB_ratio = next((elem.text.strip() for elem in PB_ratio if elem), 'N/A')

        EPS_ratio = [
            soup.find('td', string='EPS(TTM)').find_next_sibling('td'),
            soup.find('td', {'class': 'ustf141Value contentPrimary bodyLargeHeavy right-align'})
        ]
        EPS_ratio = next((elem.text.strip() for elem in EPS_ratio if elem), 'N/A')

        Div_yield = [
            soup.find('td', string='Div. Yield').find_next_sibling('td'),
            soup.find('td', {'class': 'ustf141Value contentPrimary bodyLargeHeavy right-align'})
        ]
        Div_yield = next((elem.text.strip() for elem in Div_yield if elem), 'N/A')

        book_value = [
            soup.find('td', string='Book Value').find_next_sibling('td'),
            soup.find('td', {'class': 'ustf141Value contentPrimary bodyLargeHeavy right-align'})
        ]
        book_value = next((elem.text.strip() for elem in book_value if elem), 'N/A')
    
        basic_data = [company, price, change, market_cap, volume, PE_ratio, PB_ratio, EPS_ratio, Div_yield, book_value]
        return [basic_data]
    
    except Exception as e:
        print(f'An error occurred: {e}')

def extract_detailed_info(ticker):
    headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36'
}
    url = f'https://groww.in/us-stocks/{ticker}/company-financial'
    driver = None

    try:
        firefox_options = Options()
        firefox_options.add_argument("--headless")
        firefox_options.set_preference("general.useragent.override", headers['user-agent'])

        driver = webdriver.Firefox(options=firefox_options)
        driver.get(url)

        try:
            quarterly_results_tab = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Quarterly Results')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView();", quarterly_results_tab)
            quarterly_results_tab.click()

            time.sleep(5)

            try:
                WebDriverWait(driver, 30).until(
                    lambda d: "Quarterly Results" in d.page_source and len(d.find_elements(By.XPATH, "//div[contains(text(), 'Revenues')]")) > 0
                    )
            except Exception as e:
                print(f"Warning: Wait condition failed: {e}")
        
        except Exception as e:
            print(f"Error clicking Quarterly Results tab: {e}")

        soup = BeautifulSoup(driver.page_source, 'html.parser')

        quarterly_results_terms = ["Revenues", "Cost of Revenue", "General & Administrative Expenses",
                                   "Operating Expenses", "Interest Expense", "Depreciation, Amortization & Accretion",
                                   "Earnings & Depreciation Amortization (EBITDA)", "Gross Profit",
                                   "Net Income", "Weighted Average Shares", "Operating Income"]
        
        for term in quarterly_results_terms:
            elements = soup.find_all(string=lambda s: s and term in s)
            if elements:
                for elem in elements[:2]:
                    parent= elem.parent

        def get_latest_value(label):
            try:
                label_element = None

                all_divs = soup.find_all('div')
                for div in all_divs:
                    if div.string and label.lower() in div.string.lower():
                        label_element = div
                        break

                if not label_element:
                    label_elements = soup.find_all('div', class_='valign-wrapper')
                    for element in label_elements:
                        if element.find('div', string=lambda s: s and label.lower() in s.lower().strip()):
                            label_element = element
                            break

                if not label_element:
                    elements = soup.find_all(string=lambda s: s and label.lower() in s.lower().strip())
                    if elements:
                        label_element = elements[0].parent

                if not label_element:
                    print(f"Could not find element for {label}")
                    return 'N/A'
                
                row = label_element.find_parent('div', class_=lambda c: c and 'Row' in c)

                if not row:
                    current = label_element
                    for _ in range(5):
                        if current.parent:
                            current = current.parent
                            if current.name == 'div' and len(current.find_all('div')) > 2:
                                row = current
                                break
                
                if not row:
                    print(f"Could not find container row for {label}")
                    return 'N/A'
                
                value_divs = row.find_all('div', class_=lambda c:c and 'Col' in c)

                if not value_divs or len(value_divs) < 2:
                    value_divs = row.find_all('div', revursive=False)
                
                if not value_divs or len(value_divs) < 2:
                    print(f"No value columns found for {label}")
                    return 'N/A'
                
                latest_value_div = value_divs[-1]

                value = latest_value_div.find('div', class_=lambda c: c and 'RowHead' in c)

                if not value:
                    value = latest_value_div.get_text(strip=True)
                    return value if value else 'N/A'
                
                return value.text.strip() if value and hasattr(value, 'text') else 'N/A'
            except Exception as e:
                logging.error(f"Error extracting {label}: {e}")
                return 'N/A'
            
        quarterly_balance_data = {
            'Revenues': get_latest_value('Revenues'),
            'Cost of Revenue': get_latest_value('Cost of Revenue'),
            'General & Administrative Expenses': get_latest_value('General & Administrative Expenses'),
            'Operating Expenses': get_latest_value('Operating Expenses'),
            'Interest Expense': get_latest_value('Interest Expense'),
            'Depreciation, Amortization & Accretion': get_latest_value('Depreciation, Amortization & Accretion'),
            'Earnings & Depreciation Amortization (EBITDA)': get_latest_value('Earnings & Depreciation Amortization (EBITDA)'),
            'Gross Profit': get_latest_value('Gross Profit'),
            'Net Income': get_latest_value('Net Income'),
            'Weighted Average Shares': get_latest_value('Weighted Average Shares'),
            'Operating Income': get_latest_value('Operating Income')
        }

        print("\nExtracted quarterly results:")
        for key, value in quarterly_balance_data.items():
            print(f"{key}: {value}")

        if all(value== 'N/A' for value in quarterly_balance_data.values()):
            logging.warning(f"No meaningful data extracted for {ticker}")

        return [[
            quarterly_balance_data['Revenues'],
            quarterly_balance_data['Cost of Revenue'],
            quarterly_balance_data['General & Administrative Expenses'],
            quarterly_balance_data['Operating Expenses'],
            quarterly_balance_data['Interest Expense'],
            quarterly_balance_data['Depreciation, Amortization & Accretion'],
            quarterly_balance_data['Earnings & Depreciation Amortization (EBITDA)'],
            quarterly_balance_data['Gross Profit'],
            quarterly_balance_data['Net Income'],
            quarterly_balance_data['Weighted Average Shares'],
            quarterly_balance_data['Operating Income'],
        ]]
    except Exception as e:
        logging.error(f'Error extracting quarterly results for {ticker}: {e}')
        return [['N/A'] * 10]
    finally:
        if driver:
            driver.quit()         


def extract_balance_sheet(ticker):
    headers = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    url = f'https://groww.in/us-stocks/{ticker}/company-financial'
    driver = None
    
    try:
        firefox_options = Options()
        firefox_options.add_argument("--headless")
        firefox_options.set_preference("general.useragent.override", headers['user-agent'])
        
        driver = webdriver.Firefox(options=firefox_options)
        driver.get(url)
        
        try:
            balance_sheet_tab = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Balance Sheet')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView();", balance_sheet_tab)
            balance_sheet_tab.click()
                       
            time.sleep(5)
            
            try:
                WebDriverWait(driver, 30).until(
                    lambda d: "Balance Sheet" in d.page_source and len(d.find_elements(By.XPATH, "//div[contains(text(), 'Shareholders Equity')]")) > 0
                )
            except Exception as e:
                print(f"Warning: Wait condition failed: {e}")
            
        except Exception as e:
            print(f"Error clicking Balance Sheet tab: {e}")
        
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        
        def get_latest_value(label):
            try:
                label_element = None
                
                all_divs = soup.find_all('div')
                for div in all_divs:
                    if div.string and label.lower() in div.string.lower():
                        label_element = div
                        break
                
                if not label_element:
                    label_elements = soup.find_all('div', class_='valign-wrapper')
                    for element in label_elements:
                        if element.find('div', string=lambda s: s and label.lower() in s.lower().strip()):
                            label_element = element
                            break
                
                if not label_element:
                    elements = soup.find_all(string=lambda s: s and label.lower() in s.lower().strip())
                    if elements:
                        label_element = elements[0].parent
                
                if not label_element:
                    print(f"Could not find element for {label}")
                    return 'N/A'
                
                row = label_element.find_parent('div', class_=lambda c: c and 'Row' in c)
                
                if not row:
                    current = label_element
                    for _ in range(3):
                        if current.parent:
                            current = current.parent
                            if current.name == 'div' and len(current.find_all('div')) > 2:
                                row = current
                                break
                
                if not row:
                    print(f"Could not find container row for {label}")
                    return 'N/A'
                
                # Try different ways to extract value
                
                value_divs = row.find_all('div', class_=lambda c: c and 'Col' in c)
                
                if not value_divs or len(value_divs) < 2:
                    value_divs = row.find_all('div', recursive=False)
                
                if not value_divs or len(value_divs) < 2:
                    print(f"No value columns found for {label}")
                    return 'N/A'
                
                latest_value_div = value_divs[-1]
                
                value = latest_value_div.find('div', class_=lambda c: c and 'RowHead' in c)
                
                if not value:
                    value = latest_value_div.get_text(strip=True)
                    return value if value else 'N/A'
                    
                return value.text.strip() if value and hasattr(value, 'text') else 'N/A'
            except Exception as e:
                logging.error(f"Error extracting {label}: {e}")
                return 'N/A'
            
        balance_sheet_data = {
            'Shareholders Equity': get_latest_value('Shareholders Equity'),
            'Total Assets': get_latest_value('Total Assets'),
            'Current Assets': get_latest_value('Current Assets'),
            'Non-Current Assets': get_latest_value('Assets Non-Current'),
            'Current Liabilities': get_latest_value('Current Liabilities'),
            'Non-Current Liabilities': get_latest_value('Liabilities Non-Current'),
            'Tax Liabilities': get_latest_value('Tax Liabilities'),
            'Tax Assets': get_latest_value('Tax Assets'),
            'Cash and Equivalents': get_latest_value('Cash and Equivalents'),
            'Total Liabilities': get_latest_value('Total Liabilities')
        }
        
        print("\nExtracted balance sheet data:")
        for key, value in balance_sheet_data.items():
            print(f"{key}: {value}")
        
        if all(value == 'N/A' for value in balance_sheet_data.values()):
            logging.warning(f"No meaningful data extracted for {ticker}")
        
        return [[
            balance_sheet_data['Shareholders Equity'], 
            balance_sheet_data['Total Assets'], 
            balance_sheet_data['Current Assets'], 
            balance_sheet_data['Non-Current Assets'],
            balance_sheet_data['Current Liabilities'], 
            balance_sheet_data['Non-Current Liabilities'], 
            balance_sheet_data['Tax Liabilities'], 
            balance_sheet_data['Tax Assets'],
            balance_sheet_data['Cash and Equivalents'], 
            balance_sheet_data['Total Liabilities']
        ]]
    except Exception as e:
        logging.error(f'Error extracting balance sheet for {ticker}: {e}')
        return [['N/A'] * 10]
    finally:
        if driver:
            driver.quit()

def extract_cash_flow(ticker):
    headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.105 Safari/537.36'
}
    url = f'https://groww.in/us-stocks/{ticker}/company-financial'
    driver = None

    try:
        firefox_options = Options()
        firefox_options.add_argument("--headless")
        firefox_options.set_preference("general.useragent.override", headers['user-agent'])

        driver = webdriver.Firefox(options=firefox_options)
        driver.get(url)

        try:
            cash_flow_tab = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Cash Flow')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView();", cash_flow_tab)
            cash_flow_tab.click()

            time.sleep(5)

            try:
                WebDriverWait(driver, 30).until(
                    lambda d: "Cash Flow" in d.page_source and len(d.find_elements(By.XPATH, "//div[contains(text(), 'Investing Cash Flow')]")) > 0
                )
            except Exception as e:
                print(f"Warning: Wait condition failed: {e}")
        
        except Exception as e:
            print(f"Error clicking Quarterly Results tab: {e}")

        soup = BeautifulSoup(driver.page_source, 'html.parser')

        cash_flow_terms = ["Investing Cash Flow", "Operations Cash Flow", "Financing Cash Flow",
                           "Net Cash Flow", "Free Cash Flow", "Capital Expenditure", "Cash and Equivalents",
                           "Payments & Cash Distribution", "Basic Common Share", "Working Capital"]
        
        for term in cash_flow_terms:
            elements = soup.find_all(string=lambda s: s and term in s)
            if elements:
                for elem in elements[:2]:
                    parent = elem.parent

        def get_latest_value(label):
            try:
                label_element = None

                all_divs = soup.find_all('div')
                for div in all_divs:
                    if div.string and label.lower() in div.string.lower():
                        label_element = div
                        break

                if not label_element:
                    label_elements = soup.find_all('div', class_='valign-wrapper')
                    for element in label_elements:
                        if element.find('div', string=lambda s: s and label.lower() in s.lower().strip()):
                            label_element = element
                            break

                if not label_element:
                    elements = soup.find_all(string=lambda s: s and label.lower() in s.lower().strip())
                    if elements:
                        label_element = elements[0].parent

                if not label_element:
                    print(f"Could not find element for {label}")
                    return 'N/A'
                
                row = label_element.find_parent('div', class_=lambda c: c and 'Row' in c)

                if not row:
                    current = label_element
                    for _ in range(5):
                        if current.parent:
                            current = current.parent
                            if current.name == 'div' and len(current.find_all('div')) > 2:
                                row = current
                                break
                
                if not row:
                    print(f"Could not find container row for {label}")
                    return 'N/A'
                
                value_divs = row.find_all('div', class_=lambda c:c and 'Col' in c)

                if not value_divs or len(value_divs) < 2:
                    value_divs = row.find_all('div', revursive=False)
                
                if not value_divs or len(value_divs) < 2:
                    print(f"No value columns found for {label}")
                    return 'N/A'
                
                latest_value_div = value_divs[-1]

                value = latest_value_div.find('div', class_=lambda c: c and 'RowHead' in c)

                if not value:
                    value = latest_value_div.get_text(strip=True)
                    return value if value else 'N/A'
                
                return value.text.strip() if value and hasattr(value, 'text') else 'N/A'
            except Exception as e:
                logging.error(f"Error extracting {label}: {e}")
                return 'N/A'
            
        cash_flow_data = {
            'Investing Cash Flow': get_latest_value('Investing Cash Flow'),
            'Operations Cash Flow': get_latest_value('Operations Cash Flow'),
            'Financing Cash Flow': get_latest_value('Financing Cash Flow'),
            'Net Cash Flow': get_latest_value('Net Cash Flow'),
            'Free Cash Flow': get_latest_value('Free Cash Flow'),
            'Capital Expenditure': get_latest_value('Capital Expenditure'),
            'Cash and Equivalents': get_latest_value('Cash and Equivalents'),
            'Payments & Cash Distribution': get_latest_value('Payments & Cash Distribution'),
            'Basic Common Share': get_latest_value('Basic Common Share'),
            'Working Capital': get_latest_value('Working Capital'),
        }
            
        print("\nExtracted quarterly results:")
        for key, value in cash_flow_data.items():
            print(f"{key}: {value}")

        if all(value== 'N/A' for value in cash_flow_data.values()):
            logging.warning(f"No meaningful data extracted for {ticker}")

        return [[
            cash_flow_data['Investing Cash Flow'],
            cash_flow_data['Operations Cash Flow'],
            cash_flow_data['Financing Cash Flow'],
            cash_flow_data['Net Cash Flow'],
            cash_flow_data['Free Cash Flow'],
            cash_flow_data['Capital Expenditure'],
            cash_flow_data['Cash and Equivalents'],
            cash_flow_data['Payments & Cash Distribution'],
            cash_flow_data['Basic Common Share'],
            cash_flow_data['Working Capital']
        ]]
    except Exception as e:
        logging.error(f'Error extracting quarterly results for {ticker}: {e}')
        return [['N/A'] * 10]
    finally:
        if driver:
            driver.quit() 

def export_to_excel(basic_data, detailed_data, balance_sheet, cash_flow, filename='stocks.xlsx'):
    if all(data_source is not None for data_source in(basic_data, detailed_data, balance_sheet, cash_flow)):
        basic_names = [
            'Company', 'Price', 'Change', 'Market Cap', 'Volume', 'P_E ratio', 'P_B ratio',
            'EPS(TTM)', 'Div. Yield', 'Book Value'
            ]
        
        detailed_info_names = [
            'Revenues', 'Cost of Revenues',
            'General & Administrative Expenses in USD millions',
            'Operating Expenses in USD millions', 'Interest Expense in USD millions',
            'Depreciation, Amortization & Accretion in USD millions',
            'EBITDA', 'Gross Profit', 'Net Income', 'Weighted Average Shares', 'Operating Income'
            ]
        
        balance_sheet_names = [
            'Shareholders Equity', 'Total Assets', 'Current Assets', 'Assets Non-Current',
            'Current Liabilities', 'Liabilities Non-Current', 'Tax Liabilities',
            'Tax Assets', 'Cash and Equivalents (USD)', 'Total Liabilities'
            ]
        
        cash_flow_names = [
            "Investing Cash Flow", "Operations Cash Flow", "Financing Cash Flow",
            "Net Cash Flow", "Free Cash Flow", "Capital Expenditure", "Cash and Equivalents",
            "Payments & Cash Distribution", "Basic Common Share", "Working Capital"
        ]
        
        df1 = pd.DataFrame(basic_data, columns=basic_names)
        df2 = pd.DataFrame(detailed_data, columns=detailed_info_names)
        df3 = pd.DataFrame(balance_sheet, columns=balance_sheet_names)
        df4 = pd.DataFrame(cash_flow, columns=cash_flow_names)

        if os.path.exists(filename):
            os.remove(filename)
            print('\nExisting Excel file deleted')

        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            df1.to_excel(writer, sheet_name='Sheet1', index=False, startrow=0)
            
            startrow_df2 = len(df1) + 2
            df2.to_excel(writer, sheet_name='Sheet1', index=False, startrow=startrow_df2)

            startrow_df3 = len(df2) + 5
            df3.to_excel(writer, sheet_name='Sheet1', index=False, startrow=startrow_df3)

            startrow_df3 = len(df3) + 8
            df4.to_excel(writer, sheet_name='Sheet1', index=False, startrow=startrow_df3)
        
        print('\nScraping finished and data exported!')
    else:
        print('No data scraped')

def main():
    ticker = input('Please paste your desired ticker here: ').upper()
    basic_data = extract_basic_info(ticker)
    detailed_data = extract_detailed_info(ticker)
    balance_sheet = extract_balance_sheet(ticker)
    cash_flow = extract_cash_flow(ticker)
    export_to_excel(basic_data, detailed_data, balance_sheet, cash_flow)

if __name__ == '__main__':
    main()