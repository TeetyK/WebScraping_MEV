# import ipdb; ipdb.set_trace()
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.common.action_chains import ActionChains
import time
import random
import pandas as pd
import numpy as np
import os
import glob
import re
import json
import requests
import io
import pdfplumber
import warnings
import sys
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.simplefilter(action='ignore',category=DeprecationWarning)

def cleaning_data(df:pd.DataFrame,df2:pd.DataFrame,df3:pd.DataFrame,df4:pd.DataFrame,df5:pd.DataFrame)->pd.DataFrame:
    df.columns = df.iloc[3,:]
    df = df.iloc[4:,:]
    df = df.rename(columns={np.nan:"Q"})
    df = df[df.notna().sum(axis=1) > 1]
    df = df[['Expenditure on gross domestic product (11)']].tail(1)
    df = df.rename(columns={df.columns[0]:"Norminal GDP by Expenditure"})
    df2.columns = df2.iloc[3,:]
    df2 = df2.iloc[4:,:]
    df2 = df2.rename(columns={np.nan:"Q"})
    df2 = df2[df2.notna().sum(axis=1) > 1]
    df2 = df2[['Expenditure on gross domestic product (CVM) (14)']].tail(1)
    df2 = df2.rename(columns={df2.columns[0]:"GDP by Expenditure Value NSA THG PQ"})
    df3.columns = df3.iloc[3,:]
    df3 = df3.iloc[4:,:]
    df3 = df3.rename(columns={np.nan:"Q"})
    df3 = df3[df3.notna().sum(axis=1) > 1]
    df3 = df3[['Gross National Income (27)']].tail(1)
    df3 = df3.rename(columns={df3.columns[0]:"Norminal Gross National Product"})
    df4.columns = df4.iloc[3,:]
    df4 = df4.iloc[4:,:]
    df4 = df4.rename(columns={np.nan:"Q"})
    df4 = df4[df4.notna().sum(axis=1) > 1]
    df4 = df4[['Agriculture (1)']].tail(1)
    df4 = df4.rename(columns={df4.columns[0]:"Argricultural production"})
    df5.columns = df5.iloc[3,:]
    df5 = df5.iloc[4:,:]
    df5 = df5.rename(columns={np.nan:"Q"})
    df5 = df5[df5.notna().sum(axis=1) > 1]
    df5 = df5[['Gross Domestic Product (25)','Manufacturing (6)','Accommodation  and food service activities (13)','Real Estate Activities (16)','Services (9)']].tail(1)
    df5 = df5.rename(columns={
        df5.columns[0]:"GDP by Expenditure Value NSA THG PQSA",
        df5.columns[1]:"Manufacturing (GDP)",
        df5.columns[2]:"Accommodation and Food Service Activities (GDP)",
        df5.columns[3]:"Real Estate Activities (GDP)",
        df5.columns[4]:"Services (GDP)"
        })
    df = pd.concat([df.reset_index(drop=True),df2.reset_index(drop=True),df3.reset_index(drop=True),df4.reset_index(drop=True),df5.reset_index(drop=True)],axis=1)
    df.head(1)
    return df.copy()
def cpi_base():
    options = uc.ChromeOptions()
    options.add_argument('--headless') 
    driver = uc.Chrome(options=options)
    try:
        print("กำลังเข้าเว็บ...")
        url = "https://www.thaibma.or.th/EN/CPI/CPIIndex.aspx" 
        driver.get(url)
        print("รอข้อมูลตาราง gen ขึ้นมา...")
        wait = WebDriverWait(driver, 20)
        
        wait.until(EC.presence_of_element_located((By.CLASS_NAME, "k-grid-content")))

        rows = driver.find_elements(By.CSS_SELECTOR, ".k-grid-content tr")
        
        print(f"เจอตารางทั้งหมด {len(rows)} บรรทัด")

        if rows:
            last_row = rows[-1]
            cells = last_row.find_elements(By.TAG_NAME, "td")
            month = cells[0].text
            cpi_value = cells[1].text
            
            print("\n" + "="*30)
            print(f"เดือนล่าสุด: {month}")
            print(f"ค่า CPI    : {cpi_value}")
            print("="*30)
            
        else:
            print("NOT FOUND")

    except Exception as e:
        print(f"Error: {e}")
    finally:
        driver.quit()
        driver.quit = lambda: None

def GDP():
    options = uc.ChromeOptions()
    options.binary_location = r'F:\chrome-win64\chrome-win64\chrome.exe'
    options.add_argument('--headless')
    download_folder = os.path.join(os.getcwd(), "nesdc_data")
    if not os.path.exists(download_folder):
        os.makedirs(download_folder)
    prefs = {
        "download.default_directory": download_folder,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)

    driver = uc.Chrome(options=options)

    try:
        print("กำลังเข้าเว็บ NESDC...")
        driver.get("https://www.nesdc.go.th/?p=85846") 
        time.sleep(random.uniform(0.1, 0.3))
        print("กำลังคลิกปุ่มดาวน์โหลด...")
        # link = driver.find_element(By.PARTIAL_LINK_TEXT, "All Tables QGDP")
        link = driver.find_element(By.CSS_SELECTOR, "a[href*='ddl=85847']")
        link.click()
        
        print("รอไฟล์โหลด...")
        timeout = 60
        elapsed = 0
        
        while elapsed < timeout:
            files = glob.glob(os.path.join(download_folder, "*.xls*"))
            if files:
                break
            time.sleep(random.uniform(0.1, 0.3))
            elapsed += 1
            print(f".", end="", flush=True)

        if not files:
            raise Exception("หมดเวลา หาไฟล์ไม่เจอ!")

        latest_file = max(files, key=os.path.getctime)
        print(f"\nดาวน์โหลดสำเร็จ: {latest_file}")

        print("กำลังอ่านข้อมูล...")
        gdp_by_end_data = pd.read_excel(latest_file,sheet_name="Table 1")
        gdp_data = pd.read_excel(latest_file,sheet_name="Table 2")
        gnp_data = pd.read_excel(latest_file,sheet_name="Table 3")
        arg_data = pd.read_excel(latest_file,sheet_name="Table 6")
        other_data = pd.read_excel(latest_file,sheet_name="Table 6")
        
        print("--- ตัวอย่างข้อมูล ---")
        data = cleaning_data(gdp_by_end_data,gdp_data,gnp_data,arg_data,other_data)
        print(data)
    except Exception as e:
        print(f"\nError : {e}")
    finally:
        driver.quit()
        driver.quit = lambda: None
def main():
    # options = Options()
    options = uc.ChromeOptions()
    options.binary_location = r'F:\chrome-win64\chrome-win64\chrome.exe'
    options.add_argument('--headless')
    all_data = []
    try:
        with uc.Chrome(options=options,headless=True) as driver:
            driver.get("http://books.toscrape.com/")
            time.sleep(random.uniform(0.1, 0.3))
            products = driver.find_elements(By.CLASS_NAME, "product_pod")
            print(f"เจอหนังสือทั้งหมด {len(products)} เล่ม\n")
            for product in products:
                h3_tag = product.find_element(By.TAG_NAME, "h3")
                a_tag = h3_tag.find_element(By.TAG_NAME, "a")
                title = a_tag.get_attribute("title")
                price = product.find_element(By.CLASS_NAME, "price_color").text

                stock = product.find_element(By.CLASS_NAME, "instock").text.strip()
                star_elem = product.find_element(By.CLASS_NAME, "star-rating")
                star_class = star_elem.get_attribute("class") 
                rating = star_class.split(" ")[-1] 
                print(f"📖 เจอ: {title[:20]}... | ราคา: {price}")
                all_data.append({
                    "Title": title,
                    "Price": price,
                    "Stock": stock,
                    "Rating": rating
                })
            print(driver.title)
        print("Hello from webscraping-mev!")
    finally:
        driver.quit()
        driver.quit = lambda: None
    if all_data:
        df = pd.DataFrame(all_data)
        print("\n" + "="*30)
        print("ตัวอย่างข้อมูลที่ได้:")
        print(df.head()) 
        
        filename = "google_results.csv"
        df.to_csv(filename, index=False, encoding="utf-8-sig")
        print(f"\nบันทึกไฟล์สำเร็จ! ชื่อไฟล์: {filename}")
    else:
        print("ไม่พบข้อมูล (อาจจะโดน Google บล็อก หรือเปลี่ยนโครงสร้างเว็บ)")
def cpi_core():
    caps = DesiredCapabilities.CHROME
    caps['goog:loggingPrefs'] = {'performance': 'ALL'}
    options = uc.ChromeOptions()
    options.binary_location = r'F:\chrome-win64\chrome-win64\chrome.exe'
    options.add_argument('--headless')

    driver = uc.Chrome(options=options , desired_capabilities=caps)

    try:
        url = "https://index.tpso.go.th/cpi/index-analysis-report/1"
        print(f"กำลังเข้าเว็บ {url}")
        driver.get(url) 
        time.sleep(random.uniform(10, 18))
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(random.uniform(10, 18))
        html_source = driver.page_source
        pdf_pattern = r'(https?://[^\s"\']+\.pdf)'
        found_links = re.findall(pdf_pattern, html_source)
        unique_links = list(set(found_links))
        if unique_links:
            print(f"{len(unique_links)}")
            for link in unique_links:
                print(link)
                if "tpso.go.th" in link:
                    print("Like this.")
        else:
            print("NOT FOUND")
        pdf_url = None
        logs = driver.get_log('performance')
        for entry in logs:
            try:
                message_obj = json.loads(entry.get('message'))
                message = message_obj.get('message')
                method = message.get('method')

                if method == 'Network.responseReceived':
                    response = message.get('params', {}).get('response', {})
                    mime_type = response.get('mimeType', '')
                    found_url = response.get('url', '')

                    if 'application/pdf' in mime_type or found_url.endswith('.pdf'):
                        if "blob:" not in found_url:
                            print(f"เจอ URL แล้ว! -> {found_url}")
                            pdf_url = found_url
                            break
            except Exception as e:
                continue

        if pdf_url:
            print("-" * 30)
            print(f"✅ สำเร็จ! URL จริงของ PDF คือ:\n{pdf_url}")
            print("-" * 30)
        else:
            print("❌ หาไม่เจอ:")
        response = requests.get(pdf_url)
        if response.status_code == 200:
            with pdfplumber.open(io.BytesIO(response.content)) as pdf:
                print("\n" + "="*20 + " เนื้อหาแรกของ PDF " + "="*20)
                page1 = pdf.pages[-2]
                text = page1.extract_text()
                print(text)
                print("\n" + "="*20 + "ข้อมูลตาราง" + "="*20)
                tables = page1.extract_tables()
                if tables:
                    for i , table in enumerate(tables):
                        print(f"ตาราง {i+1}")
                        df = pd.DataFrame(table[1:], columns=table[0])
                        print(df)
                        month_cpi_core = df.columns[2]
                        print(month_cpi_core)
                        df = df.iloc[2:,:]
                        col = ['รายการ','สัดส่วน/น้ำหนัก','ดัชนี ธค 68','ดัชนี ธค 67','Change ธค M/M','Change ธค Y/Y','Change ธค A/A',
                                'ดัชนี พย 68','Change พย M/M','Change พย Y/Y','Change พย A/A',
                                ]
                        df.columns = col
                        df = df[df['รายการ'].isin(['ดัชนีรำคำผู้บริโภคพนื้ ฐำน *'])][['ดัชนี ธค 68','Change ธค Y/Y']]
                        df.columns = ['Core CPI','(Inflation)']
                        df.to_excel(f"table_{i+1}.xlsx", index=False)
                        print(df)
                else:
                    print("NOT FOUND")

    except Exception as e:
        print(f"\nError : {e}")
    finally:
        driver.quit()
        driver.quit = lambda: None
def set_index():
    caps = DesiredCapabilities.CHROME
    caps['goog:loggingPrefs'] = {'performance': 'ALL'}
    options = uc.ChromeOptions()
    options.page_load_strategy = 'eager'
    options.binary_location = r'F:\chrome-win64\chrome-win64\chrome.exe'
    options.add_argument('--headless=new')
    options.add_argument('--disable-popup-blocking')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--start-maximized')
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    driver = uc.Chrome(options=options , desired_capabilities=caps)
    driver.set_window_size(1920, 1080)
    try:
        url = "https://th.investing.com/indices/thailand-set-historical-data"
        print(f"กำลังเข้าเว็บ {url}")
        try:
            driver.get(url)
        except:
            driver.execute_script("window.stop();") 
        wait = WebDriverWait(driver,6)
        actions = ActionChains(driver)
        try:
            close_btn = driver.find_element(By.CSS_SELECTOR, "i.popupCloseIcon, div.e-dialog__close, svg[data-test='close-icon']")
            close_btn.click()
            print("ปิด Popup แล้ว")
        except:
            pass
        # TEST 3
        print("\n1. อ่านค่าตารางปัจจุบัน...")
        driver.execute_script("window.scrollBy(0, 300);")
        old_date_text = ""
        try:
            first_date_elem = wait.until(EC.visibility_of_element_located((
                By.CSS_SELECTOR, "table tbody tr:first-child td:first-child"
            )))
            old_date_text = first_date_elem.text.strip()
            print(f"   -> ค่าเดิม (รายวัน): '{old_date_text}'")
        except:
            pass
        print("\n2. กำลังกดเปลี่ยนเป็น 'รายเดือน'...")
        dropdown_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(@class, 'selection-arrow')]")))
        actions.move_to_element(dropdown_btn).click().perform()
        time.sleep(1)
        monthly_option = wait.until(EC.element_to_be_clickable((
            By.XPATH, 
            "//div[contains(@class, 'menu-row') and .//span[contains(text(), 'รายเดือน')]]"
        )))
        actions.move_to_element(monthly_option).click().perform()
        print("   -> จิ้มเลือกรายเดือนแล้ว!")
        print("\n3. กำลังรอข้อมูลรายเดือน (ต้องขึ้นต้นด้วย '01')...")
        is_monthly_loaded = False
        max_wait_sec = 20 
        for i in range(max_wait_sec):
            try:
                current_elem = driver.find_element(By.CSS_SELECTOR, "table tbody tr:first-child td:first-child")
                current_text = current_elem.text.strip()
                if current_text != old_date_text and current_text.startswith("01"):
                    print(f"\n✅ ใช่เลย! ข้อมูลเปลี่ยนเป็นรายเดือนแล้ว: '{current_text}'")
                    is_monthly_loaded = True
                    break
                else:
                    sys.stdout.write(f".") 
                    sys.stdout.flush()
                    time.sleep(1)
            except:
                time.sleep(1)
        if not is_monthly_loaded:
            print("\n⚠️ หมดเวลา! ตารางยังไม่ยอมเปลี่ยนเป็นวันที่ 01 (ลองดูดข้อมูลเผื่อฟลุ๊ค)")
        # dropdown_btn = wait.until(EC.element_to_be_clickable((
        #     By.XPATH, 
        #     "//div[contains(@class, 'selection-arrow')]"
        # )))
        # driver.execute_script("arguments[0].style.border='3px solid red'", dropdown_btn)
        # driver.execute_script("arguments[0].click();", dropdown_btn)
        # time.sleep(random.uniform(5, 6))
        # monthly_xpath = "//div[contains(@class, 'menu-row') and .//span[contains(text(), 'รายเดือน')]]"
        # monthly_option = wait.until(EC.element_to_be_clickable((By.XPATH, monthly_xpath)))
        # driver.execute_script("arguments[0].style.border='3px solid blue'", monthly_option)
        # driver.execute_script("arguments[0].click();", monthly_option)
        time.sleep(random.uniform(2, 3))
        driver.execute_script("window.scrollTo(0, 500);")
        time.sleep(random.uniform(2, 3))
        driver.execute_script("window.scrollTo(500, 0);")
        dfs = pd.read_html(driver.page_source)
        target_df = None
        for df in dfs:
            if 'วันเดือนปี' in df.columns or 'Date' in df.columns:
                target_df = df
                break
        if target_df is not None:
            print(target_df.head())
            target_df.to_excel('set_index_historical_data.xlsx',index=False)
        else:
            print("NOT FOUND")


    except Exception as e:
        print(f"\nError : {e}")
    finally:
        driver.quit()
        driver.quit = lambda: None
    pass
if __name__ == "__main__":
    # main()
    # GDP()
    # cpi_base()
    # cpi_core()
    set_index()