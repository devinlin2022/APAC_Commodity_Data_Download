import os
import time
import base64
import pandas as pd
from retrying import retry
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pygsheets

print("🚀 Script starting...")

try:
    # 直接使用 pygsheets.authorize 创建一个全局客户端对象
    gc = pygsheets.authorize(service_file='service_account_key.json')
    print("✅ Google Sheets authenticated successfully.")
except FileNotFoundError:
    print(f"❌ FATAL: Service account file 'service_account_key.json' not found.")
    raise
except Exception as e:
    print(f"❌ FATAL: Error authenticating. Check your GCP_SA_KEY secret's JSON format. Error: {e}")
    raise


def save_pdf(driver, path):
    """保存当前页面为 PDF，用于调试。"""
    print(f"📄 Saving debug PDF to {path}...")
    settings = {
        "landscape": False, "displayHeaderFooter": False,
        "printBackground": True, "preferCSSPageSize": True
    }
    result = driver.execute_cdp_cmd("Page.printToPDF", settings)
    pdf_data = base64.b64decode(result['data'])
    with open(path, 'wb') as f:
        f.write(pdf_data)
    print("✅ PDF saved.")

@retry(stop_max_attempt_number=3, wait_fixed=5000)
def fetch_data(link):
    """使用 Selenium 抓取数据"""
    print(f"➡️ Fetching data from: {link}")
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=options)
    
    username = os.getenv("ICIS_USERNAME")
    password = os.getenv("ICIS_PASSWORD")

    try:
        driver.implicitly_wait(10)
        driver.get(link)

        wait = WebDriverWait(driver, 60)
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#login-button')))

        driver.execute_script(f'document.querySelector("#username-input").value = "{username}"')
        driver.execute_script(f'document.querySelector("#password-input").value = "{password}"')
        driver.execute_script(f'document.querySelector("#login-button").click()')
        print("🔐 Login submitted.")
        
        try:
            wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#continue-login-button')))
            time.sleep(2)
            driver.execute_script(f'document.querySelector("#continue-login-button").click()')
            print("🖱️ 'Continue Login' button clicked.")
        except Exception:
            print("ℹ️ 'Continue Login' button not found, proceeding.")

        price_selector = '#content > div > div > div > div > div.Zoomstyle__BodyContainer-LbgNq.fhHJpQ > div.Zoomstyle__Section-hqZqfX.jKLgrv > div.Largestyle__DisplayWrapperLarge-iWzxqM.hISDst > div.Largestyle__DisplayItem-vzpFY.fbUftf > div > div:nth-child(2) > div > div > div.PriceDeltastyle__DeltaContainer-jdFEoE.dtfcmD > div.Textstyles__Heading1Blue-gtxuIB.dzShK'
        date_selector = '#content > div > div > div > div > div.Zoomstyle__BodyContainer-LbgNq.fhHJpQ > div.Zoomstyle__Section-hqZqfX.jKLgrv > div.Largestyle__DisplayWrapperLarge-iWzxqM.hISDst > div.Mainstyle__Group-ciNpsy.fYvNPb > div > div > div:nth-child(2) > div'

        print("⏳ Waiting for data to load...")
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, price_selector)))

        element_script = f"return document.querySelector('{price_selector}').textContent;"
        element = driver.execute_script(element_script)
        date_element = driver.find_element(By.CSS_SELECTOR, date_selector)
        date = date_element.text
        print(f"📊 Data fetched successfully: Date='{date}', Price='{element}'")
        return element, date

    except Exception as e:
        print(f"❌ An error occurred during fetch_data. Saving debug PDF.")
        save_pdf(driver, "webpage_error.pdf")
        raise e
    finally:
        driver.quit()

# --- vvv 这里是唯一的、必要的修改 vvv ---
def upload_to_google_sheet(client, data, sheet_key, worksheet_name, row):
    """
    !!! 已做微小修改，不再依赖全局 'creds' !!!
    它现在接收一个已认证的客户端 'client' 作为第一个参数。
    """
    # gc = pygsheets.client.Client(creds)  <-- 已删除此行
    wb_key = client.open_by_key(sheet_key) # <-- 直接使用传入的 client
    try:
        sheet = wb_key.worksheet_by_title(worksheet_name)
    except pygsheets.WorksheetNotFound:
        print(f"Worksheet {worksheet_name} not found in sheet with key {sheet_key}.")
        return

    # ... 函数的其余部分与您的原始版本完全相同 ...
    try:
        if row[4] == "ICIS_APAC":
            new_row = data.values.tolist()[0]
        elif row[4] == "ICIS_Common":
            original_row = data.values.tolist()[0]
            new_row = [original_row[0], '', original_row[1]]
        else:
            print(f"Unexpected value in column E: {row[4]}")
            return

        all_values = sheet.get_all_values()
        last_non_empty_row = 0

        for i, existing_row in enumerate(all_values):
            if any(cell.strip() for cell in existing_row):
                last_non_empty_row = i + 1

        empty_row_index = last_non_empty_row + 1

        if empty_row_index > sheet.rows:
            sheet.add_rows(empty_row_index - sheet.rows + 1000)

        sheet.update_row(empty_row_index, new_row)
        print(f"✅ Row successfully added to {worksheet_name}.")
    except Exception as e:
        print(f"❌ Failed to add row to {worksheet_name}: {e}")
# --- ^^^ 这里是唯一的、必要的修改 ^^^ ---

def main():
    """主函数"""
    print("\n--- Starting Main Process ---")
    # 注意：这里不再需要 gc = pygsheets.authorize(...)，因为它已经在顶层完成了
    sh = gc.open_by_key('1clmwUEhzplke2naZlCrCwAh2jJ017vbZd9pNVSKh_EI')
    wks = sh.worksheet_by_title('Python_Commodity')

    data = wks.get_all_values()

    for i, row in enumerate(data[1:], start=2):
        print(f"\n--- Processing Master Sheet Row {i} ---")
        if len(row) < 5:
            print(f"⏭️ Skipping row {i}: not enough columns.")
            continue

        sheet_key, worksheet_name, commodity_name, link, category = row[:5]

        if not sheet_key or not worksheet_name or not link:
            print(f"⏭️ Skipping row {i}: missing critical information.")
            continue

        try:
            price, date = fetch_data(link)
            price_data = pd.DataFrame([[date, price]], columns=['Date', 'Commodity'])
            # 将全局客户端 gc 传递给函数
            upload_to_google_sheet(gc, price_data, sheet_key, worksheet_name, row)
        except Exception as e:
            print(f"☠️ FATAL ERROR processing row {i}. Moving to next row. Error: {e}")

if __name__ == "__main__":
    main()
    print("\n🎉 Script finished.")
