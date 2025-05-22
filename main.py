import requests
import bs4
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

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# 공고 리스트 페이지 이동
url = 'https://www.saramin.co.kr/zf_user/jobs/list/industry?ind_key=12117&panel_type=&search_optional_item=n&search_done=y&panel_count=y&preview=y'
driver.get(url)
time.sleep(3)

# XPath 기준 공고 요소 추출
elements = driver.find_elements(By.XPATH, '/html/body/div[3]/div[1]/div/div[5]/div/div[3]/section/div/div/div[1]/div[2]/div[1]/a')
print(f"[공고 수]: {len(elements)}개")

# 공고 URL 리스트 만들기
links = [elem.get_attribute('href') for elem in elements]

# 저장 디렉토리 생성
os.makedirs("screenshots", exist_ok=True)

# 첫번째 공고 바디 저장
element = driver.find_element(By.CSS_SELECTOR, ".user_content")

# 하나씩 방문하면서 OCR 검사 + 저장
for i, link in enumerate(links):
    driver.get(link)
    time.sleep(3)

    # CDP 기반 전체 화면 캡처
    rect = driver.execute_script("""
        const elem = arguments[0];
        const rect = elem.getBoundingClientRect();
        return {
            x: rect.x,
            y: rect.y,
            width: rect.width,
            height: rect.height
        };
        """, element)
    # page_rect = driver.execute_cdp_cmd('Page.getLayoutMetrics', {})
    # screenshot_config = {
    #     'captureBeyondViewport': True,
    #     'fromSurface': True,
    #     'clip': {
    #         'width': page_rect['cssContentSize']['width'],
    #         'height': page_rect['cssContentSize']['height'],
    #         'x': 0,
    #         'y': 0,
    #         'scale': 1
    #     }
    # }
    base64_png = driver.execute_cdp_cmd('Page.captureScreenshot', rect)

    # 임시 파일 저장
    temp_path = f"screenshots/temp_{i}.png"
    with open(temp_path, "wb") as f:
        f.write(base64.urlsafe_b64decode(base64_png['data']))

    # OCR로 텍스트 추출
    text = pytesseract.image_to_string(Image.open(temp_path), lang='kor')

    keywords = ["자격", "자격증", "소지자", "취득"]

    if any(keyword in text for keyword in keywords):
        save_path = f"screenshots/공고_{i+1}.png"
        os.rename(temp_path, save_path)
        print(f"[통과]: {save_path}")
    else:
        os.remove(temp_path)
        print(f"[불통과]: {link}")

driver.quit()

