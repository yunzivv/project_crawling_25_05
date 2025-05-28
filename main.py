# 크롤링 라이브러리
import requests
from bs4 import BeautifulSoup
import time
import re

import pandas as pd

# 캡쳐, OCR
import os

# 날짜 출력
from datetime import date
import random


# 추적 방지를 위한 헤더 설정
headers = {
    "User-Agent": "Mozilla/5.0",
    "Referer": "https://www.jobkorea.co.kr/recruit/joblist?menucode=duty",
    "Content-Type": "application/json",
    "X-Requested-With": "XMLHttpRequest",
    "Accept-Encoding": "gzip, deflate, br"
}

# 기획, 전략 > 마케팅 기획 2787
dutyCtgr = "10026" # 직무 카테코리
duty = "1000187" # 직무

payload = {
    "condition": {
        "dutyCtgr": 0,
        "duty": duty,
        "dutyArr": [duty],
        "dutyCtgrSelect": [dutyCtgr],
        "dutySelect": [duty],
        "isAllDutySearch": False
    },
    "TotalCount": 2787,
    "Page": 1,
    "PageSize": 500
}

# 세션 생성 -> headers 추가 -> POST 방식으로 요청 보내기
# 잡코리아 공고 기본 url
url = "https://www.jobkorea.co.kr/Recruit/Home/_GI_List/"

# 쿠키
session = requests.Session()
session.get("https://www.jobkorea.co.kr/")
session.headers.update(headers)
response = session.post(url, json=payload)

# 검색된 자격증 저장 리스트
certificates = []

# 요청 성공 시 html 문서 파싱
if response.status_code == 200:
    
    soup = BeautifulSoup(response.text, "lxml")
    jobs = soup.select(".devTplTabBx table .tplTit > .titBx")

    for job in jobs:

        # 공고에서 링크 추출
        a_tag = job.select_one("a")
        href = a_tag["href"] if a_tag else None 
        if href:

            # 공고 ID 추출
            match = re.search(r'/Recruit/GI_Read/(\d+)', href)

            # 각 공고 상세 페이지 요청
            if match:
                gno = match.group(1)
                detail_url = f"https://www.jobkorea.co.kr/Recruit/GI_Read/{gno}"
                try:
                    detail_res = session.get(detail_url, timeout=10)
                    time.sleep(random.uniform(1, 3))
                    # 요청 성공 시 html 문서 파싱, 해당 요소 찾기
                    if detail_res.status_code == 200:
                        detail_soup = BeautifulSoup(detail_res.text, "lxml")

                        # 팝업 dt
                        popup_pref = detail_soup.select_one("#popupPref")

                        if popup_pref:
                            dt_elements = popup_pref.select(".tbAdd dt")

                        else:
                            dt_elements = detail_soup.select(".artReadJobSum .tbList dt")

                        # 우대 자격증 추출 / 저장
                        for dt in dt_elements:
                            if '자격' in dt.text:
                                dd = dt.find_next_sibling("dd")
                                if dd:
                                    
                                    cert_text = dd.text.strip().rstrip(',')
                                    cert_list = [cert.strip() for cert in cert_text.split(',') if cert.strip()] 

                                    for cert in cert_list:
                                        print(gno + "번 공고 우대 자격증 : " + cert)
                                        certificates.append({
                                            "직무 카테고리": dutyCtgr,
                                            "직무 코드": duty,
                                            "공고번호": gno,
                                            "자격증": cert,
                                            "수집일": date.today()
                                        })
                    else:
                        print("[오류]"+ gno + "번 상세 페이지 응답 실패:", detail_res.status_code)

                except Exception as e:
                    print("[오류]"+ gno + "번 상세 페이지 요청 오류:", e)
            else:
                if "www.gamejob.co.kr" in href:
                    continue
                print("[오류] 링크에서 공고 ID 추출 실패:", href)

else:
    print("[오류] 리스트 페이지 요청 실패:", response.status_code)

# pandas 사용 -> 엑셀 파일로 저장
df = pd.DataFrame(certificates)
file_path = "jobkorea_requirements.xlsx"

if os.path.exists(file_path):
    # 기존 파일 읽기
    existing_df = pd.read_excel(file_path)
    # 기존 + 신규 데이터 결합
    combined_df = pd.concat([existing_df, df], ignore_index=True)
else:
    combined_df = df

# 저장 (덮어쓰기)
combined_df.to_excel(file_path, index=False)
print("✔ 종료")
