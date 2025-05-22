import requests
import bs4
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time

service = Service('./drivers/chromedriver.exe')
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')
driver = webdriver.Chrome(service=service, options=options)

# url, 경로 설정
url = 'https://www.naver.com/'
path = 'TagName_Body-headlessOptions.png'

<<<<<<< HEAD
=======
#실행
driver.get(url)
time.sleep(1)
el = driver.find_element(By.TAG_NAME,'body')
el.screenshot(path)

# google 껐다 키기
# service = Service('./drivers/chromedriver.exe')
# driver = webdriver.Chrome(service = service)

# driver.get("http://www.google.com")
# print(driver.title)
# driver.quit()

# 크롤링 기본 options 설정
# options = webdriver.ChromeOptions()
# options.add_argument("--disable-blink-features=AutomationControlled")
# options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/117 Safari/537.36")

# driver = webdriver.Chrome(options=options)

>>>>>>> 6cee9a1 (full screen 캡처 도전 -> 성공)
