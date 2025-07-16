import os
import time
import base64
import pandas as pd
from retrying import retry
import pygsheets
import gspread # 新增

# --- 认证模块 (已采纳你的建议) ---
try:
    # 你的建议是使用 service_account_key.json，我们在这里保持一致
    SERVICE_ACCOUNT_FILE = 'service_account_key.json'
    
    # 初始化 pygsheets 客户端 (用于你现有的 upload 函数)
    gc_pygsheets = pygsheets.authorize(service_file=SERVICE_ACCOUNT_FILE)
    
    # 初始化 gspread 客户端 (可用于未来的新功能)
    gc_gspread = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
    
    # 兼容旧代码，创建一个 legacy creds 对象给 upload_to_google_sheet 函数使用
    creds = gc_pygsheets.client.creds
    
    print("✅ Google Sheets authenticated successfully for both pygsheets and gspread.")

except Exception as e:
    print(f"❌ Error authenticating with Google Sheets service account: {e}")
    # 认证失败是致命错误，直接退出
    raise SystemExit(e)

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ... (从这里开始，save_pdf, fetch_data, upload_to_google_sheet, main 函数与之前版本完全相同)
# ... (为简洁起见，此处省略，你无需修改这些函数)

def save_pdf(driver, path):
    """保存当前页面为 PDF，用于调试。"""
    settings = {
        "landscape": False,
        "displayHeaderFooter": False,
        "printBackground": True,
        "preferCSSPageSize": True
    }
    result = driver.execute_cdp_cmd("Page.printToPDF", settings)
    pdf_data = base64.b64decode(result['data'])
    with open(path, 'wb') as f:
        f.write(pdf_data)
    print(f"📄 调试 PDF 已保存至: {path}")

@retry(stop_max_attempt_number=3, wait_fixed=5000)
def fetch_data(link):
    """使用 Selenium 抓取数据。"""
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument("window-size=1920,1080")

    icis_username = os.environ.get('ICIS_USERNAME')
    icis_password = os.environ.get('ICIS_PASSWORD')
    if not icis_username or not icis_password:
        raise ValueError("ICIS_USERNAME 或 ICIS_PASSWORD 环境变量未设置！")

    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(10)
    print(f"➡️ 正在访问: {link}")
    driver.get(link)

    wait = WebDriverWait(driver, 60)
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#login-button')))
    
    driver.execute_script(f'document.querySelector("#username-input").value = "{icis_username}"')
    driver.execute_script(f'document.querySelector("#password-input").value = "{icis_password}"')
    driver.execute_script(f'document.querySelector("#login-button").click()')
    print("🔐 已输入用户名密码并点击登录。")
    
    try:
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#continue-login-button')))
        time.sleep(2)
        driver.execute_script(f'document.querySelector("#continue-login-button").click()')
        print("🖱️ 已点击 'Continue Login' 按钮。")
    except Exception:
        print("ℹ️ 未找到 'Continue Login' 按钮，继续执行。")
        pass

    price_selector = '#content > div > div > div > div > div.Zoomstyle__BodyContainer-LbgNq.fhHJpQ > div.Zoomstyle__Section-hqZqfX.jKLgrv > div.Largestyle__DisplayWrapperLarge-iWzxqM.hISDst > div.Largestyle__DisplayItem-vzpFY.fbUftf > div > div:nth-child(2) > div > div > div.PriceDeltastyle__DeltaContainer-jdFEoE.dtfcmD > div.Textstyles__Heading1Blue-gtxuIB.dzShK'
    date_selector = '#content > div > div > div > div > div.Zoomstyle__BodyContainer-LbgNq.fhHJpQ > div.Zoomstyle__Section-hqZqfX.jKLgrv > div.Largestyle__DisplayWrapperLarge-iWzxqM.hISDst > div.Mainstyle__Group-ciNpsy.fYvNPb > div > div > div:nth-child(2) > div'

    try:
        print("⏳ 正在等待数据加载...")
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, price_selector)))
        print("✅ 数据已加载。")
    except Exception as e:
        print(f"❌ 等待数据加载超时或失败: {e}")
        save_pdf(driver, "webpage_error.pdf")
        driver.quit()
        raise

    element_script = f"return document.querySelector('{price_selector}').textContent;"
    element = driver.execute_script(element_script)
    date_element = driver.find_element(By.CSS_SELECTOR, date_selector)
    date = date_element.text
    
    print(f"📊 抓取成功: Price='{element}', Date='{date}'")
    driver.quit()
    return element, date

def upload_to_google_sheet(data, sheet_key, worksheet_name, row):
    """
    将数据上传到指定的 Google Sheet。
    此函数与您的原始脚本完全相同，以满足您的要求。
    它依赖于在脚本顶层创建的全局 'creds' 对象。
    """
    gc = pygsheets.client.Client(creds)
    wb_key = gc.open_by_key(sheet_key)
    try:
        sheet = wb_key.worksheet_by_title(worksheet_name)
    except pygsheets.WorksheetNotFound:
        print(f"Worksheet {worksheet_name} not found in sheet with key {sheet_key}.")
        return

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
        print(f"✅ 成功将数据行添加到工作表: {worksheet_name}。")
    except Exception as e:
        print(f"❌ 添加到 {worksheet_name} 失败: {e}")

def main():
    """主函数，用于协调整个流程。"""
    # 使用 pygsheets 客户端读取主控表
    sh = gc_pygsheets.open_by_key('1clmwUEhzplke2naZlCrCwAh2jJ017vbZd9pNVSKh_EI')
    wks = sh.worksheet_by_title('Python_Commodity')
    
    master_data = wks.get_all_records()
    print(f"ℹ️ 从主控表找到 {len(master_data)} 条记录。")

    for i, record in enumerate(master_data):
        print(f"\n--- 正在处理第 {i+1}/{len(master_data)} 条记录 ---")
        sheet_key = record.get('sheet_key')
        worksheet_name = record.get('worksheet_name')
        commodity_name = record.get('commodity_name')
        link = record.get('link')
        category = record.get('category')
        
        original_row = [sheet_key, worksheet_name, commodity_name, link, category]

        if not all([sheet_key, worksheet_name, link, category]):
            print(f"⏭️ 跳过不完整的记录: {record}")
            continue
        
        print(f" commodity: {commodity_name}, category: {category}")

        try:
            price, date = fetch_data(link)
            price_data = pd.DataFrame([[date, price]], columns=['Date', 'Commodity'])
            upload_to_google_sheet(price_data, sheet_key, worksheet_name, original_row)
        except Exception as e:
            print(f"☠️ 处理链接 {link} 时发生严重错误: {e}")
            continue

if __name__ == "__main__":
    main()
