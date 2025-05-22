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

# 검색 키워드
keywords = ["자격", "자격증", "소지자", "취득"]

# 하나씩 방문하면서 OCR 검사 + 저장
for i, link in enumerate(links):
    try:
        driver.get(link)
        time.sleep(3)  # 페이지 로딩 대기

        # iframe이 있다면 switch_to.frame()으로 해당 iframe으로 전환
        try:
            iframe = driver.find_element(By.TAG_NAME, 'iframe')
            driver.switch_to.frame(iframe)  # iframe으로 전환

            # .user_content 요소 찾기
            user_content_elements = driver.find_elements(By.CLASS_NAME, "user_content")
            if user_content_elements:
                # 첫 번째 .user_content 요소 찾기
                user_content_element = user_content_elements[0]

                # 요소의 위치와 크기 가져오기
                location = user_content_element.location
                size = user_content_element.size
                print(f"위치: {location}, 크기: {size}")

                # 전체 화면 캡처
                screenshot_path = f"screenshots/full_{i}.png"
                driver.save_screenshot(screenshot_path)

                # PIL로 캡처된 이미지를 열고, 요소 영역만 잘라내기
                img = Image.open(screenshot_path)
                left = location['x']
                top = location['y']
                right = left + size['width']
                bottom = top + size['height']

                # 영역 자르기
                img = img.crop((left, top, right, bottom))
                temp_path = f"screenshots/temp_{i}.png"
                img.save(temp_path)

                # OCR 텍스트 추출
                text = pytesseract.image_to_string(img, lang='kor')

                # 키워드 필터링
                if any(keyword in text for keyword in keywords):
                    save_path = f"screenshots/공고_{i+1}.png"
                    os.rename(temp_path, save_path)
                    print(f"[통과]: {save_path}")
                else:
                    os.remove(temp_path)
                    print(f"[불통과]: {link}")

            else:
                print(f"[오류 발생 - .user_content 요소 없음]: {link}")

            # 작업이 끝난 후 iframe 밖으로 나오기
            driver.switch_to.default_content()

        except Exception as e:
            print(f"[오류 발생 - iframe 처리 중 오류]: {link} - {e}")

    except Exception as e:
        print(f"[오류 발생 - 링크 접속 실패]: {link}\n{e}")

driver.quit()