# import requests
# import bs4
import base64
import pytesseract
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from PIL import Image


service = Service('./drivers/chromedriver.exe')
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')
driver = webdriver.Chrome(service=service, options=options)

pytesseract.pytesseract.tesseract_cmd = r"D:\정윤지\Tesseract-OCR\tesseract.exe"

os.environ['TESSDATA_PREFIX'] = r"D:\정윤지\Tesseract-OCR\tessdata"

# 공고 리스트 페이지 이동
url = 'https://www.jobkorea.co.kr/recruit/joblist?menucode=local&localorder=1'
driver.get(url)
time.sleep(3)

# XPath 기준 공고 요소 추출
elements = driver.find_elements(By.XPATH, '/html/body/div[5]/div[1]/div/div[2]/div[5]/div/div[5]/table/tbody/tr/td[2]/div/strong/a')
print(f"[공고 수]: {len(elements)}개")

# 공고 URL 리스트 만들기
links = [elem.get_attribute('href') for elem in elements]

# 저장 디렉토리 생성
os.makedirs("screenshots", exist_ok=True)

# 검색 키워드
keywords = ["자격증"]

# 하나씩 방문하면서 OCR 검사 + 저장
for i, link in enumerate(links):
    try:
        driver.get(link)
        time.sleep(3)

        elements = driver.find_elements(By.CLASS_NAME, "tbRow")
        if not elements:
            print(f"[오류 발생 요소 없음]: {link}")
            break

        element = elements[0]
        location = element.location
        size = element.size
        print(f"위치: {location}, 크기: {size}")

        screenshot_path = f"screenshots/full_{i}.png"
        driver.save_screenshot(screenshot_path)

        img = Image.open(screenshot_path)
        left = location['x']
        top = location['y']
        right = left + size['width']
        bottom = top + size['height']
        img = img.crop((left, top, right, bottom))

        temp_path = f"screenshots/temp_{i}.png"
        img.save(temp_path)

        try:
            text = pytesseract.image_to_string(img, lang='kor')
            if text:
                # utf-8 디코딩 에러 방지
                clean_text = text.encode('utf-8', errors='ignore').decode('utf-8')
                print(clean_text)
            else:
                clean_text = ""

        except Exception as e:
            print(f"OCR 오류: {e}")
            clean_text = ""

        if "자격증" in clean_text:
            save_path = f"screenshots/공고_{i+1}.png"
            os.rename(temp_path, save_path)
            print(f"[통과]: {save_path}")
        else:
            os.remove(temp_path)
            print(f"[불통과]: {link}")

        driver.switch_to.default_content()

    except Exception as e:
        print(f"[오류 발생 - 링크 접속 실패]: {link}\n{e}")
        break  # break 대신 계속 진행

driver.quit()