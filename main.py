import requests
from bs4 import BeautifulSoup
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
url = "https://www.jobkorea.co.kr/Recruit/Home/_GI_List/"

headers = {
    "User-Agent": "Mozilla/5.0",
    "Referer": "https://www.jobkorea.co.kr/recruit/joblist?menucode=local",
    "Content-Type": "application/json",
    "X-Requested-With": "XMLHttpRequest"
}

payload = {
    "condition": {
        "dutyCtgr": 0,
        "duty": "1000308",               # 프롬프트 엔지니어 등 해당 직무 코드
        "dutyArr": ["1000308"],
        "dutyCtgrSelect": ["10038"],
        "dutySelect": ["1000308"],
        "isAllDutySearch": False
    },
    "TotalCount": 55,
    "Page": 1,
    "PageSize": 40
}

response = requests.post(url, headers=headers, json=payload)

if response.status_code == 200:
    soup = BeautifulSoup(response.text, "html.parser")
    jobs = soup.select("table .tplTit > .titBx")  # 리스트 형태에 따라 수정 필요
    for job in jobs:
        title = job.select_one(".titBx > strong").text.strip() if job.select_one(".titBx > strong") else "제목 없음"
        print("🧾 채용 공고:", title)
else:
    print("❌ 요청 실패:", response.status_code)