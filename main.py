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
from google.oauth2 import service_account

# --- GitHub Actions 环境下的认证 (已修正) ---

# 1. 明确定义所需的 API 权限范围
#    - spreadsheets: 读写 Google Sheets
#    - drive: 访问 Google Drive (pygsheets 需要此权限来按名称查找电子表格)
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# 2. 从 service_account.json 文件和定义的 SCOPES 创建凭证
try:
    SERVICE_ACCOUNT_FILE = 'service_account.json'
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=SCOPES  # 使用我们刚刚定义的 SCOPES
    )
except FileNotFoundError:
    print("错误：'service_account.json' 未找到。请确保 GitHub Actions 工作流正确生成了此文件。")
    creds = None
except Exception as e:
    print(f"加载凭证时发生错误: {e}")
    creds = None

# --- GitHub Actions 环境下的认证 ---
# 从 GitHub Secrets 生成的 service_account.json 文件进行授权
# pygsheets.DEFAULT_SCOPES 包含了读写 aheets 和 drive 的权限

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
    print(f"调试 PDF 已保存至: {path}")

# 使用 retry 装饰器，如果出现异常，会自动重试2次，每次间隔5秒
@retry(stop_max_attempt_number=3, wait_fixed=5000)
def fetch_data(link):
    """使用 Selenium 抓取数据。"""
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument("window-size=1920,1080") # 设置窗口大小以避免元素不可见

    # 从环境变量中获取用户名和密码
    icis_username = os.environ.get('ICIS_USERNAME')
    icis_password = os.environ.get('ICIS_PASSWORD')
    if not icis_username or not icis_password:
        raise ValueError("ICIS_USERNAME 或 ICIS_PASSWORD 环境变量未设置！")

    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(10)
    print(f"正在访问: {link}")
    driver.get(link)

    wait = WebDriverWait(driver, 60)
    wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#login-button')))
    
    # 使用环境变量填充登录信息
    driver.execute_script(f'document.querySelector("#username-input").value = "{icis_username}"')
    driver.execute_script(f'document.querySelector("#password-input").value = "{icis_password}"')
    driver.execute_script(f'document.querySelector("#login-button").click()')
    print("已输入用户名密码并点击登录。")
    
    try:
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#continue-login-button')))
        time.sleep(2) # 等待页面可能的变化
        driver.execute_script(f'document.querySelector("#continue-login-button").click()')
        print("已点击 'Continue Login' 按钮。")
    except Exception:
        print("未找到 'Continue Login' 按钮，继续执行。")
        pass

    # 等待关键数据元素出现
    price_selector = '#content > div > div > div > div > div.Zoomstyle__BodyContainer-LbgNq.fhHJpQ > div.Zoomstyle__Section-hqZqfX.jKLgrv > div.Largestyle__DisplayWrapperLarge-iWzxqM.hISDst > div.Largestyle__DisplayItem-vzpFY.fbUftf > div > div:nth-child(2) > div > div > div.PriceDeltastyle__DeltaContainer-jdFEoE.dtfcmD > div.Textstyles__Heading1Blue-gtxuIB.dzShK'
    date_selector = '#content > div > div > div > div > div.Zoomstyle__BodyContainer-LbgNq.fhHJpQ > div.Zoomstyle__Section-hqZqfX.jKLgrv > div.Largestyle__DisplayWrapperLarge-iWzxqM.hISDst > div.Mainstyle__Group-ciNpsy.fYvNPb > div > div > div:nth-child(2) > div'

    try:
        print("正在等待数据加载...")
        wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, price_selector)))
        print("数据已加载。")
    except Exception as e:
        print(f"等待数据加载超时或失败: {e}")
        save_pdf(driver, "webpage_error.pdf") # 保存页面快照以便调试
        driver.quit()
        raise # 重新抛出异常，触发 retry

    element_script = f"return document.querySelector('{price_selector}').textContent;"
    element = driver.execute_script(element_script)
    date_element = driver.find_element(By.CSS_SELECTOR, date_selector)
    date = date_element.text
    
    print(f"抓取成功: Price='{element}', Date='{date}'")
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
        print(f"工作表 {worksheet_name} 在 Key 为 {sheet_key} 的文件中未找到。")
        return

    try:
        if row[4] == "ICIS_APAC":
            new_row = data.values.tolist()[0]
        elif row[4] == "ICIS_Common":
            original_row = data.values.tolist()[0]
            new_row = [original_row[0], '', original_row[1]]
        else:
            print(f"在 E 列发现意外的值: {row[4]}")
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
        print(f"成功将数据行添加到工作表: {worksheet_name}。")
    except Exception as e:
        print(f"添加到 {worksheet_name} 失败: {e}")

def main():
    """主函数，用于协调整个流程。"""
    if not creds:
        print("因凭证加载失败，程序终止。")
        return

    # 使用带有凭证的客户端对象
    gc = pygsheets.authorize(custom_credentials=creds)
    sh = gc.open_by_key('1clmwUEhzplke2naZlCrCwAh2jJ017vbZd9pNVSKh_EI')
    wks = sh.worksheet_by_title('Python_Commodity')
    
    master_data = wks.get_all_records() # 使用 get_all_records 更方便处理
    print(f"从主控表找到 {len(master_data)} 条记录。")

    for i, record in enumerate(master_data):
        print(f"\n--- 处理第 {i+1} 条记录 ---")
        # 从记录中获取数据
        sheet_key = record.get('sheet_key')
        worksheet_name = record.get('worksheet_name')
        commodity_name = record.get('commodity_name')
        link = record.get('link')
        category = record.get('category')
        
        # 将原始数据行转换为列表，以兼容您现有的 upload 函数
        original_row = [sheet_key, worksheet_name, commodity_name, link, category]

        if not all([sheet_key, worksheet_name, link, category]):
            print(f"跳过不完整的记录: {record}")
            continue
        
        print(f"商品: {commodity_name}, 类别: {category}")

        try:
            price, date = fetch_data(link)
            price_data = pd.DataFrame([[date, price]], columns=['Date', 'Commodity'])
            upload_to_google_sheet(price_data, sheet_key, worksheet_name, original_row)
        except Exception as e:
            print(f"处理链接 {link} 时发生严重错误: {e}")
            # 如果需要，可以在此处添加邮件通知等失败处理逻辑
            continue

if __name__ == "__main__":
    main()
