import requests
import bs4
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time

driver = webdriver.Chrome()

# 연습 링크
driver.get("https://www.jobkorea.co.kr/Recruit/GI_Read/46954687?rPageCode=SL&logpath=21&sn=6&sc=612")  # 원하는 공고 링크

time.sleep(2)

# dlPref 영역 전체 가져오기
# /html/body/div[5]/section/section/div[1]/article/div[2]/div[1]/dl/dd/dl/dd/span//////
pref_section = driver.find_element(By.CSS_SELECTOR, "dl.tbList")
dt_elements = pref_section.find_elements(By.TAG_NAME, "dt")
dd_elements = pref_section.find_elements(By.TAG_NAME, "dd")

# dt와 dd를 순서대로 쌍으로 매칭
for dt, dd in zip(dt_elements, dd_elements):
    if dt.text.strip() == "자격증":
        try:
            cert_text = dd.find_element(By.CSS_SELECTOR, "span.pref").text.strip()
            print(f"✅ 자격증 : {cert_text}")
        except:
            print("❌ 자격증 span.pref 없음")
        break
else:
    print("⚠️ '자격증' 항목 없음")

driver.quit()