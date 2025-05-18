import requests
import bs4
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time

options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/117 Safari/537.36")

driver = webdriver.Chrome(options=options)

# 링크 예시로 하나만
hrefs = ["https://www.jobkorea.co.kr/Recruit/GI_Read/46954687?rPageCode=SL&logpath=21&sn=6&sc=612"]  # <- 여기에 실제 링크들 넣어줘

for href in hrefs:
    driver.get(href)
    time.sleep(3)

    print(f"\n🔎 {href}")

    ## ✅ 산업(업종) 항목
    try:
        info_section = driver.find_element(By.CSS_SELECTOR, "div.tbCoInfo > dl.tbList")
        dt_elements = info_section.find_elements(By.TAG_NAME, "dt")
        dd_elements = info_section.find_elements(By.TAG_NAME, "dd")

        for dt, dd in zip(dt_elements, dd_elements):
            if "산업" in dt.text:
                try:
                    text_tag = dd.find_element(By.TAG_NAME, "text")
                    print(f"업종: {text_tag.text}")
                except:
                    print("⚠️ text 태그가 dd 안에 없음")
    except Exception as e:
        print("⚠️ 산업(업종) 섹션 오류:", e)

    ## ✅ 자격증 항목
    try:
        dl_section = driver.find_element(By.CSS_SELECTOR, "dl.tbAdd.tbPref")
        dt_elements = dl_section.find_elements(By.TAG_NAME, "dt")
        dd_elements = dl_section.find_elements(By.TAG_NAME, "dd")

        for dt, dd in zip(dt_elements, dd_elements):
            if "자격증" in dt.text:
                span = dd.find_element(By.CSS_SELECTOR, "span.pref")
                print(f"자격증: {span.text}")
    except Exception as e:
        print("⚠️ 자격증 섹션 오류:", e)

driver.quit()
