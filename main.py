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
import re



# 잡코리아
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
        "duty": "1000262",
        "dutyArr": ["1000262"],
        "dutyCtgrSelect": ["10032"],
        "dutySelect": ["10010002620414"],
        "isAllDutySearch": False
    },
    "TotalCount": 34,
    "Page": 1,
    "PageSize": 50
}

base_url = "https://www.jobkorea.co.kr"
session = requests.Session()
session.headers.update(headers)

response = session.post(url, json=payload)

if response.status_code == 200:
    soup = BeautifulSoup(response.text, "html.parser")
    jobs = soup.select("table .tplTit > .titBx")

    for job in jobs:
        a_tag = job.select_one("a")
        title = a_tag.text.strip() if a_tag else "제목 없음"
        href = a_tag["href"] if a_tag else None

        print("🧾 채용 공고:", title)

        if href:
            # 공고 ID 추출 (ex. /Recruit/GI_Read/49693541 → 49693541)
            match = re.search(r'/Recruit/GI_Read/(\d+)', href)
            if match:
                gno = match.group(1)
                detail_url = f"https://www.jobkorea.co.kr/Recruit/GI_Read/{gno}"

                try:
                    detail_res = session.get(detail_url, timeout=10)
                    time.sleep(1.5)

                    if detail_res.status_code == 200:
                        detail_soup = BeautifulSoup(detail_res.text, "html.parser")
                        dt_elements = detail_soup.select(".artReadJobSum .tbList dt")
                        for dt in dt_elements:
                            if '자격' in dt.text:
                                dd = dt.find_next_sibling("dd")
                                if dd:
                                    print("   📝 자격 요건:", dd.text.strip())
                    else:
                        print("❌ 상세 페이지 응답 실패:", detail_res.status_code)

                except Exception as e:
                    print("❌ 상세 페이지 요청 오류:", e)
            else:
                print("⚠️ 링크에서 공고 ID 추출 실패:", href)

else:
    print("❌ 리스트 페이지 요청 실패:", response.status_code)