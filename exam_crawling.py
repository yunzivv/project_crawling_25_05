import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook
import time

# ocr-env_exam\Scripts\activate
# python exam_crawling.py

# # 게시판 링크 가져오기
# # 웹 페이지 요청
# url = 'https://www.comcbt.com/xe/anne'
# response = requests.get(url)
# soup = BeautifulSoup(response.text, 'html.parser')

# # CSS 선택자에 맞는 링크들 추출
# links = soup.select('#gnb > ul > li:nth-child(2) > ul > li:nth-child(6) > ul li > a')

# # 딕셔너리 형태로 파싱
# data = []
# for link in links:
#     title = link.get_text(strip=True)
#     href = link.get('href')
#     full_url = requests.compat.urljoin(url, href)  # 상대 경로 보완
#     data.append({'자격등급': '산업기사', 
#                  '자격증명': title, 
#                  'href': full_url, 
#                  '종류': '필기', 
#                  'regDate': time.strftime('%Y.%m.%d - %H:%M:%S')})

# file_path = 'exam.xlsx'

# # 데이터프레임 생성
# df_new = pd.DataFrame(data)

# # 기존 파일의 첫 번째 시트의 마지막 행 다음부터 이어쓰기
# book = load_workbook(file_path)
# sheetname = book.sheetnames[0]
# start_row = book[sheetname].max_row

# # 기존 파일에 이어쓰기
# with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
#     df_new.to_excel(writer, sheet_name=sheetname, startrow=start_row, index=False, header=False)


# 게시판에서 hwp 다운로드

import os
import re
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://www.comcbt.com"
BOARD_URL = "https://www.comcbt.com/xe/df"
SAVE_DIR = "hwp_files"
os.makedirs(SAVE_DIR, exist_ok=True)

def get_post_links():
    res = requests.get(BOARD_URL)
    soup = BeautifulSoup(res.text, 'html.parser')
    posts = []
    links = soup.select('td.title a[href]')
    print(f"게시글 링크 개수: {len(links)}")

    for a in soup.select('td.title a[href]'):
        
        title = a.get_text(strip=True)
        print(f"발견된 제목: {title}")

        # (복원중) 포함 시 건너뛰기
        if '(복원중)' in title:
            print(f"⏭️ 건너뜀 (복원중): {title}")
            continue

        # 제목에서 연도 추출 (4자리 숫자)
        year_match = re.search(r'(20\d{2})', title)
        if year_match:
            year = int(year_match.group(1))
            if year < 2020:
                print(f"⏭️ 건너뜀 (연도미달): {title}")
                continue
        else:
            print(f"⏭️ 건너뜀 (연도없음): {title}")
            continue

        href = a['href']
        full_url = requests.compat.urljoin(BASE_URL, href)
        posts.append((title, full_url))
    return posts

def download_hwp_from_post(title, post_url):
    print(f"➡️ 게시글 접속: {title} - {post_url}")
    res = requests.get(post_url)
    soup = BeautifulSoup(res.text, 'html.parser')

    links = soup.find_all('a')
    print(f"  전체 링크 개수: {len(links)}")

    hwp_links = []
    for link in links:
        link_text = link.get_text(strip=True)
        if link_text.endswith('.hwp') and '(교사용)' in link_text:
            hwp_links.append((link_text, link.get('href')))

    print(f"  조건에 맞는 링크 개수: {len(hwp_links)}")

    if not hwp_links:
        print("  ⚠️ 조건에 맞는 '(교사용).hwp' 링크 없음")
        return

    # 첫 번째 링크만 다운로드
    link_text, href = hwp_links[0]
    file_url = href  # href는 이미 절대 URL임
    file_name = file_url.split('file_srl=')[-1] + '.hwp'  # 혹은 link_text 그대로 써도 됨
    save_path = os.path.join(SAVE_DIR, file_name)

    print(f"📥 다운로드 중: {link_text} - 파일명: {file_name}")
    file_content = requests.get(file_url).content
    with open(save_path, 'wb') as f:
        f.write(file_content)

posts = get_post_links()
for title, post_url in posts:
    download_hwp_from_post(title, post_url)
