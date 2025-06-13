import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook

# ocr-env_exam\Scripts\activate
# python exam_crawling.py

# 웹 페이지 요청
url = 'https://www.comcbt.com/xe/anne'
response = requests.get(url)
soup = BeautifulSoup(response.text, 'html.parser')

# CSS 선택자에 맞는 링크들 추출
links = soup.select('#gnb > ul > li:nth-child(2) > ul > li:nth-child(4) > ul li > a')

# 딕셔너리 형태로 파싱
data = []
for link in links:
    title = link.get_text(strip=True)
    href = link.get('href')
    full_url = requests.compat.urljoin(url, href)  # 상대 경로 보완
    data.append({'자격증명': title, '필기 기출 링크': full_url, '종류': '필기'})

file_path = 'exam.xlsx'

# 데이터프레임 생성
df_new = pd.DataFrame(data)

# 기존 파일의 첫 번째 시트의 마지막 행 다음부터 이어쓰기
book = load_workbook(file_path)
sheetname = book.sheetnames[0]
start_row = book[sheetname].max_row

# 기존 파일에 이어쓰기 (book 직접 건드리지 않음)
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
    df_new.to_excel(writer, sheet_name=sheetname, startrow=start_row, index=False, header=False)