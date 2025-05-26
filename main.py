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


# 공고 리스트 페이지 이동
url = "https://www.jobkorea.co.kr/Recruit/Home/_GI_List/"

headers = {
    "User-Agent": "Mozilla/5.0",
    "Referer": "https://www.jobkorea.co.kr/recruit/joblist?menucode=local",
    "Content-Type": "application/json",
    "X-Requested-With": "XMLHttpRequest"
}

# 프롬프트 엔지니어 직무 코드
payload = {
    "condition": {
        "dutyCtgr": 0,
        "duty": "1000308",
        "dutyArr": ["1000308"],
        "dutyCtgrSelect": ["10038"],
        "dutySelect": ["1000308"],
        "isAllDutySearch": False
    },
    "TotalCount": 55,
    "Page": 1,
    "PageSize": 50
}

response = requests.post(url, headers=headers, json=payload)

if response.status_code == 200:
    soup = BeautifulSoup(response.text, "html.parser")
    jobs = soup.select("table .tplTit > .titBx")

    for job in jobs:
        a_tag = job.select_one("a")
        title = a_tag.text.strip() if a_tag else "제목 없음"
        href = a_tag["href"] if a_tag else None

        print("🧾 채용 공고:", title)

        if href:
            detail_url = url + href
            detail_res = requests.get(detail_url, headers=headers)

            if detail_res.status_code == 200:
                detail_soup = BeautifulSoup(detail_res.text, "html.parser")

                # '자격'이라는 텍스트가 있는 dt를 찾고, 그 다음 dd를 추출
                dt_elements = detail_soup.select(".artReadJobSum .tbList dt")
                for dt in dt_elements:
                    if '자격' in dt.text:
                        next_dd = dt.find_next_sibling("dd")
                        if next_dd:
                            print("📝 자격 요건:", next_dd.text.strip())
            else:
                print("❌ 상세 페이지 요청 실패:", detail_res.status_code)
else:
    print("❌ 요청 실패:", response.status_code)