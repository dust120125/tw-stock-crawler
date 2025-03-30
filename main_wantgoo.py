# --- START OF FILE main_f22_wantgoo_excel_final.py ---

import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import json
import time
import os
import random
import pandas as pd
import subprocess # Needed for devnull

# Selenium imports
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

# Stealth import
from selenium_stealth import stealth

# ---

print("Fetching company lists...")
# (Company list fetching code remains the same...)
# 獲取第一個公司ID資訊表
company_info_url1 = "https://openapi.twse.com.tw/v1/opendata/t187ap03_L"
try:
    response1 = requests.get(company_info_url1, timeout=30)
    response1.raise_for_status() # Raise an exception for bad status codes
    company_info_data1 = response1.json()
except requests.exceptions.RequestException as e:
    print(f"Error fetching company list 1: {e}")
    company_info_data1 = []

# 提取第一個公司ID資訊表的公司ID和公司名稱
company_id_name_map1 = {entry["公司代號"]: entry["公司名稱"]
                        for entry in company_info_data1 if "公司代號" in entry and "公司名稱" in entry}

# 獲取第二個公司ID資訊表
company_info_url2 = "https://www.tpex.org.tw/openapi/v1/mopsfin_t187ap03_O"
try:
    response2 = requests.get(company_info_url2, timeout=30)
    response2.raise_for_status() # Raise an exception for bad status codes
    company_info_data2 = response2.json()
except requests.exceptions.RequestException as e:
    print(f"Error fetching company list 2: {e}")
    company_info_data2 = []

# 提取第二個公司ID資訊表的公司ID和公司名稱
company_id_name_map2 = {entry["SecuritiesCompanyCode"]
    : entry["CompanyName"] for entry in company_info_data2 if "SecuritiesCompanyCode" in entry and "CompanyName" in entry}

# 合併兩個公司ID和公司名稱的映射
company_id_name_map = {**company_id_name_map1, **company_id_name_map2}

total_companies = len(company_id_name_map)
print(f"總計 {total_companies} 家公司")

# 初始化一個空的列表來存儲所有公司的資訊
all_companies_data = []

# --- Selenium WebDriver Setup ---
print("Setting up WebDriver...")
options = webdriver.ChromeOptions()
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.102 Safari/537.36")
options.add_argument('--headless')
options.add_argument("--log-level=3") # Keep this, it helps with browser-level logs
options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"]) # Exclude both automation and logging switches
options.add_experimental_option('useAutomationExtension', False)
options.add_argument('--disable-logging') # Explicitly disable logging
options.add_argument('--silent') # Try to make webdriver-manager quieter
options.add_argument('--enable-unsafe-swiftshader')


driver = None
processed_count = 0

try:
    # --- Suppress ChromeDriver Console Output ---
    # Get the driver path using webdriver-manager
    driver_path = ChromeDriverManager().install()
    # Define the service with suppressed output
    service = ChromeService(
        executable_path=driver_path,
        log_output=subprocess.DEVNULL # Redirect log output to null device
        # Older versions might use: log_path=os.devnull
        )
    print("Initializing Chrome Driver with suppressed service logs...")
    # Pass the custom service object to the driver
    driver = webdriver.Chrome(service=service, options=options)
    # -----------------------------------------

    print("Applying stealth...")
    stealth(driver,
            languages=["en-US", "en"],
            vendor="Google Inc.",
            platform="Win32",
            webgl_vendor="Intel Inc.",
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=True,
            )

    driver.set_page_load_timeout(45)
    print("WebDriver setup complete with stealth.")
    # -----------------------------

    # (Scraping loop remains the same as the previous version)
    # 迭代所有公司ID
    for idx, (company_id, company_name) in enumerate(company_id_name_map.items(), start=1):
        # Limit processing for testing if needed
        # if processed_count >= 5:
        #     print("\nReached processing limit for testing.")
        #     break

        print(f"\nProcessing {idx}/{total_companies}: ID {company_id}, Name: {company_name}")
        url = f"https://www.wantgoo.com/stock/{company_id}/financial-statements/monthly-revenue"
        html_content = ""

        try:
            print(f"  Navigating to {url}...")
            driver.get(url)

            wait_timeout = 25
            wait = WebDriverWait(driver, wait_timeout)
            target_selector = "tbody[monthlyrevenue] tr[monthlyrevenue-item]"
            print(f"  Waiting up to {wait_timeout}s for table row visibility ({target_selector})...")
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, target_selector)))

            print("  Table row should be visible. Fetching page source...")
            time.sleep(0.5)
            html_content = driver.page_source

            if not html_content:
                 print("  Error: html_content is empty after wait. Skipping parsing.")
                 continue

            print("  Parsing HTML...")
            soup = BeautifulSoup(html_content, 'html.parser')

            table = soup.find('table', class_='table-sticky')
            if not table:
                print(f"  Error: Could not find the revenue table for {company_id} on page. Skipping.")
                continue

            tbody = table.find('tbody', attrs={'monthlyrevenue': ''})
            if not tbody:
                print(f"  Error: Could not find the table body (tbody[monthlyrevenue]) for {company_id}. Skipping.")
                continue

            rows = tbody.find_all('tr', attrs={'monthlyrevenue-item': ''})
            if not rows:
                print(f"  Warning: No data rows found in the table for {company_id}. Skipping.")
                continue

            company_revenue_data = []
            print(f"  Extracting data from {len(rows)} rows...")
            for row in rows:
                cols = row.find_all('td')
                if len(cols) == 8:
                    try:
                        row_data = {
                            '年度/月份': cols[0].text.strip(),
                            '當月營收': cols[1].text.strip(), # Keep as string for now
                        }
                        company_revenue_data.append(row_data)
                    except Exception as e:
                         print(f"  Error extracting data from a row for {company_id}: {e}. Skipping row.")
                else:
                    print(f"  Warning: Found row with {len(cols)} columns for {company_id}, expected 8. Skipping row.")

            if not company_revenue_data:
                 print(f"  Error: Extracted 0 valid data rows for {company_id} after parsing. Skipping.")
                 continue

            company_info = {
                '公司名稱': company_name,
                '公司代號': company_id,
                '營收資訊': company_revenue_data
            }
            all_companies_data.append(company_info)
            processed_count += 1
            print(f"  Successfully processed {company_id}. Found {len(company_revenue_data)} rows.")

        except TimeoutException:
            print(f"  Error: Timeout ({wait_timeout}s) waiting for table row visibility for {company_id} at {url}. Skipping.")
        except WebDriverException as e:
             # Filter out common "disconnected" or "target crashed" errors which might be less critical
             err_str = str(e).lower()
             if "disconnected" not in err_str and "target crashed" not in err_str:
                 print(f"  Error: WebDriver error for {company_id} at {url}: {e}. Skipping.")
             else:
                 print(f"  Warning: WebDriver connection issue for {company_id} ({err_str}). Attempting to continue.")
        except Exception as e:
            print(f"  Error: An unexpected error occurred processing {company_id}: {e}. Skipping.")

        time.sleep(random.uniform(1.0, 2.5))


except Exception as e:
    print(f"An critical error occurred during WebDriver setup or main loop: {e}")
finally:
    if driver:
        print("\nClosing WebDriver...")
        driver.quit()

# --- Data Restructuring and Excel Export ---
if all_companies_data:
    print(f"\nRestructuring data for {len(all_companies_data)} companies...")
    excel_data = {}
    all_dates = set()

    for company_data in all_companies_data:
        company_id = company_data['公司代號']
        revenue_dict = {}
        for monthly_info in company_data['營收資訊']:
            date_str = monthly_info['年度/月份']
            revenue_str = monthly_info['當月營收']
            revenue_dict[date_str] = revenue_str
            all_dates.add(date_str)
        excel_data[company_id] = revenue_dict

    sorted_dates = sorted(list(all_dates), reverse=True)

    df = pd.DataFrame.from_dict(excel_data, orient='index')
    df = df.reindex(columns=sorted_dates)
    df = df.reset_index()
    df = df.rename(columns={'index': '公司代號'})

    print("Converting revenue data to numeric...")
    revenue_columns = df.columns[1:]
    for col in revenue_columns:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(',', '', regex=False)
            df[col] = pd.to_numeric(df[col], errors='coerce').astype('Int64')

    excel_output_file = '公司五年營收.xlsx'
    sheet_name = '營收資料'
    print(f"Saving data to Excel file with formatting: {excel_output_file}...")

    try:
        with pd.ExcelWriter(excel_output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            currency_format = '#,##0'

            for col_idx in range(2, worksheet.max_column + 1):
                for row_idx in range(2, worksheet.max_row + 1):
                    worksheet.cell(row=row_idx, column=col_idx).number_format = currency_format

            print("Adjusting column widths...")
            for column_cells in worksheet.columns:
                # Get the column letter
                column_letter = column_cells[0].column_letter
                try:
                    # Calculate max length needed for the column
                    max_length = 0
                    # Check header length first
                    header_value = str(column_cells[0].value)
                    max_length = len(header_value)

                    # Check data cell lengths
                    for cell in column_cells[1:]: # Skip header row
                        if cell.value is not None: # Handle empty cells
                            try:
                                # Attempt formatted length for numbers, fallback to string length
                                if isinstance(cell.value, (int, float)) and cell.number_format != 'General':
                                     # Use a simple way to estimate formatted length (commas add width)
                                     s_val = str(cell.value)
                                     num_commas = (len(s_val) - 1) // 3 if cell.value >= 1000 else 0
                                     cell_len = len(s_val) + num_commas
                                else:
                                    cell_len = len(str(cell.value))
                            except Exception: # Fallback if formatting fails unexpectedly
                                cell_len = len(str(cell.value))

                            if cell_len > max_length:
                                max_length = cell_len

                    # Apply width with padding, limiting max width
                    adjusted_width = min(max_length + 2, 50) # Increased max slightly
                    worksheet.column_dimensions[column_letter].width = adjusted_width

                except Exception as e_width:
                    # Silently ignore width adjustment errors for specific columns if they occur
                    # print(f"  Info: Could not adjust width for column {column_letter}: {e_width}") # Keep commented out
                    pass

        print(f"Successfully saved formatted data to {excel_output_file}")

    except Exception as e:
        print(f"Error saving or formatting Excel file: {e}")

else:
    print("\nNo data was successfully processed. Skipping Excel export.")

print("Script finished.")
# --- END OF FILE main_f22_wantgoo_excel_final.py ---