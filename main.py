import requests
import bs4
import base64
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

# 실행
driver.get(url)
time.sleep(1)

# We need the dimensions of the content
page_rect = driver.execute_cdp_cmd('Page.getLayoutMetrics', {})

# parameters needed for ful page screenshot
# note we are setting the width and height of the viewport to screenshot, same as the site's content size
screenshot_config = {'captureBeyondViewport': True,
                             'fromSurface': True,
                             'clip': {'width': page_rect['cssContentSize']['width'],
                                      'height': page_rect['cssContentSize']['height'], #contentSize -> cssContentSize
                                      'x': 0,
                                      'y': 0,
                                      'scale': 1},
                             }
# Dictionary with 1 key: data
base_64_png = driver.execute_cdp_cmd('Page.captureScreenshot', screenshot_config)
# driver.execute_script("document.body.style.zoom='80%'")

# Write img to file
with open("chrome-devtools-protocol.png", "wb") as fh:
    fh.write(base64.urlsafe_b64decode(base_64_png['data']))

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
