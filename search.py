from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, WebDriverException
import pdb
import time
import pandas as pd

df_sheet_index = pd.read_excel("./keywords.xlsx")

chrome_options = Options()
chrome_options.add_argument("--disable-extensions")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--disable-popup-blocking")
chrome_options.add_argument("--window-size=1920,1080")
chrome_options.add_argument("--lang=ko-KR")
chrome_options.add_argument('--user-agent="Mozilla/5.0 (Linux; Android 8.0; SM-S10 Lite) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.121 Mobile Safari/537.36"')

keywords = df_sheet_index['키워드목록']
if len(keywords) > 0:
  browser = webdriver.Chrome(chrome_options=chrome_options)
  for index, keyword in enumerate(keywords):
    browser.get("https://display.cjonstyle.com/p/search/searchAllList?k={}&searchType=ALL".format(keyword))
    # time.sleep(3)
    try:
        myElem = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.CLASS_NAME, 'search_result_info')))
        has_searched_result = False
        try:
          browser.find_element_by_class_name("dsc_notice")
        except WebDriverException:
          has_searched_result = True
        
        df_sheet_index['검색결과'][index] = "있음" if has_searched_result else "없음"
        print("검색결과 {}".format("있음!" if has_searched_result else "없음ㅠ"))
    except TimeoutException:
        print("로딩 타임아웃!!!")
  
  df_sheet_index.to_excel('keywords_result.xlsx', sheet_name="Sheet1", index=False)
  browser.quit()
else:
  print("키워드 리스트를 입력해주세요.")

# browser = webdriver.Chrome(chrome_options=chrome_options)
# browser.get("https://display.cjonstyle.com/p/search/searchAllList?k=%ED%99%94%EC%9E%A5%ED%92%88&searchType=ALL") # 검색결과 있는경우
# # browser.get("https://display.cjonstyle.com/p/search/searchAllList?k=%EA%B0%9C%EA%BF%80%EB%B0%A5&searchType=ALL") # 검색결과 없는경우