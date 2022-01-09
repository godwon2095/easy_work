from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, WebDriverException
import pdb
import pandas as pd

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

elements_count = int(input("키워드 개수를 입력해주세요: "))
print("\n")
keywords = list(input() for _ in range(elements_count))
search_result_arr = []

if len(keywords) > 0:
  browser = webdriver.Chrome(chrome_options=chrome_options, executable_path="C:/chromedriver.exe")
  for index, keyword in enumerate(keywords):
    browser.get("https://display.cjonstyle.com/p/search/searchAllList?k={}&searchType=ALL".format(keyword))
    # time.sleep(3)
    try:
        myElem = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, ".shop_area, .dsc_notice")))
        has_searched_result = False
        # pdb.set_trace()
        try:
          browser.find_element_by_class_name("dsc_notice")
        except WebDriverException:
          has_searched_result = True
          
        have_text = "있음"
        no_text = "없음"
        
        if has_searched_result:
          search_result_arr.append(have_text)
        else:
          search_result_arr.append(no_text)
      
        print("검색결과 {}".format("있음!" if has_searched_result else "없음ㅠ"))
    except TimeoutException:
        print("로딩 타임아웃!!!")


  df = pd.DataFrame({
      '키워드': keywords,
      '검색결과': search_result_arr,
  })
  writer = pd.ExcelWriter("keywords_{}.xlsx".format(elements_count), engine='xlsxwriter')
  df.to_excel(writer, sheet_name='Sheet1', index=False)
  writer.save()

  df_sheet_index = pd.read_excel("./keywords_{}.xlsx".format(elements_count))
  styled = (df_sheet_index.style.applymap(lambda v: 'color: %s' % 'red' if v=='없음' else ''))
  styled.to_excel("./keywords_{}.xlsx".format(elements_count), sheet_name="Sheet1", index=False)
  
  print("엑셀파일이 생성되었습니다")
        
  browser.quit()
else:
  print("키워드 리스트를 입력해주세요.")

# browser = webdriver.Chrome(chrome_options=chrome_options)
# browser.get("https://display.cjonstyle.com/p/search/searchAllList?k=%ED%99%94%EC%9E%A5%ED%92%88&searchType=ALL") # 검색결과 있는경우
# # browser.get("https://display.cjonstyle.com/p/search/searchAllList?k=%EA%B0%9C%EA%BF%80%EB%B0%A5&searchType=ALL") # 검색결과 없는경우
