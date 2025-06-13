import requests
from bs4 import BeautifulSoup

# ocr-env_exam\Scripts\activate
# python exam_crawling.py

# 웹 페이지 요청
url = 'https://www.comcbt.com/xe/anne'
response = requests.get(url)
soup = BeautifulSoup(response.text, 'html.parser')

# CSS 선택자에 맞는 링크들 추출
links = soup.select('#gnb > ul > li:nth-child(2) > ul > li:nth-child(4) > ul li > a')

# 딕셔너리로 저장
link_dict = {link.get_text(strip=True): link.get('href') for link in links}


# 결과 출력
for text, href in link_dict.items():
    print(f"자격증: {text}, 필기 기출 링크: {href}")