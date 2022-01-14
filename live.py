from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, WebDriverException
import pdb
import pandas as pd
import datetime

chrome_options = Options()
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--disable-popup-blocking")

browser = webdriver.Chrome(chrome_options=chrome_options, executable_path="C:/chromedriver.exe")

date = input("날짜를 입력해주세요 ex: 20220109 :")

browser.get("https://display.cjonstyle.com/p/tv/tvSchedule?broadType=live#bdDt%3A{}".format(date))

try:
    myElem = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".state_bar")))
    
    try:
      browser.execute_script("window.scrollTo(0, 0);")
      WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".pass")))
    except:
      pass
    
    state_bars = browser.find_elements(By.CLASS_NAME, "state_bar")
    time_arr = []
    item_name_arr = []
    item_code_arr = []
    
    for index, state_bar in enumerate(state_bars):
        live_time = ""
        try:
            text_list = state_bar.text.split("\n")
            if index == 0:
                live_time = text_list[1]
            else:
                live_time = text_list[0]
        except:
            pass
        
        if len(live_time) > 0:
            time_arr.append(live_time)
    
    for i in range(3, 49):
        try:
            item = browser.find_elements(By.CSS_SELECTOR, "#scheduleItem > ul:nth-child({}) > li:nth-child(1) > div > strong > a".format(i))
            item_name_arr.append(item[0].text)
            item_code_arr.append(int(item[0].get_attribute('href').split("/")[-1].split("?")[0]))
        except:
            pass

    days = ["월", "화", "수", "목", "금", "토", "일"]
    b = days[datetime.date(int(date[0:4]), int(date[4:6]), int(date[6:])).weekday()]
    date_arr = ["{}/{}({})".format(date[4:6], date[6:], b) for i in range(len(time_arr))]
    
    df = pd.DataFrame({
        '날짜': date_arr,
        '시간대': time_arr,
        '상품명': item_name_arr,
        '상품코드': item_code_arr
    })
    writer = pd.ExcelWriter("live_{}.xlsx".format(date), engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    print("엑셀파일이 생성되었습니다")
  
except TimeoutException:
    print("로딩 타임아웃!!!")
  
browser.quit()
  
