import time
import pandas as pd
import compare

from selenium import webdriver
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common import exceptions
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from bs4 import BeautifulSoup
from datetime import datetime

time_string = datetime.now().strftime('%Y-%m-%d')
timeout = 20
data = []

def get_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')                 # 瀏覽器不提供可視化頁面
    options.add_argument('-no-sandbox')               # 以最高權限運行
    options.add_argument('--start-maximized')        # 縮放縮放（全屏窗口）設置元素比較準確
    options.add_argument('--disable-gpu')            # 谷歌文檔說明需要加上這個屬性來規避bug
    options.add_argument('--window-size=1920,1080')  # 設置瀏覽器按鈕（窗口大小）
    options.add_argument('--incognito')               # 啟動無痕

    driver = webdriver.Chrome(options=options)
    url = 'https://bank.sinopac.com/mma8/bank/html/rate/bank_ExchangeRate.html'
        
    # driver.implicitly_wait(10)
    # driver.get(url)
    # driver.delete_all_cookies() #清cookie 

    # with open('cookies.yml', 'r', encoding='utf-8') as f:
    #     cookies = yaml.load(f, Loader=yaml.FullLoader)
    #     for c in cookies:
    #         driver.add_cookie(c)

    # print("Current cookies:", driver.get_cookies())
    driver.get(url)
    

    return driver    

def get_rate():
    driver = get_driver()
    rows = BeautifulSoup(WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, '//*[@id="tab1"]/table/tbody'))).get_attribute('innerHTML'), 'html.parser').find_all('tr')   
     
    for row in rows:
        cols = row.find_all('td')
        if len(cols) > 0:
            Currency = cols[0].text
            R_buy = cols[1].text
            R_shell = cols[2].text
            N_buy = cols[3].text
            N_shell = cols[4].text
            # print(N_buy, N_shell)
            # print(Currency, R_buy, R_shell, N_buy, N_shell)
        
            record = {
              '幣別': Currency,
              '匯款-銀行買入-Bank Buy': R_buy,
              '匯款-銀行賣出-Bank Sell': R_shell,
              '現鈔-銀行買入-Bank Buy': N_buy,
              '現鈔-銀行賣出-Bank Sell': N_shell,
             }
        
            data.append(record)
            
    df = pd.DataFrame(data)
    df.to_excel(f'{time_string}_永豐牌告匯率.xlsx', index=False, header=True)
    # print(df)

if __name__ == "__main__":
    get_rate()