import requests
from bs4 import BeautifulSoup
import pandas as pd

# ocr-env_certList
# certList_crawling.py

url = "https://www.pqi.or.kr/inf/qul/infQulNanLis.do"
headers = {
    "User-Agent": "Mozilla/5.0",
    "Content-Type": "application/x-www-form-urlencoded"
}

data_rows = []  # 저장할 결과 리스트

# 1 ~ 21페이지까지 반복
for page in range(1, 22):
    data = {
        "pageIndex": str(page),
        "natId": "",                # 필요 시 채우기
        "orderId": "qulNmAsc"
    }

    response = requests.post(url, data=data, headers=headers)
    soup = BeautifulSoup(response.text, "html.parser")

    rows = soup.select("table.board_list tr")

    for row in rows:
        tds = row.select("td")
        row_data = {}
        for idx, td in enumerate(tds, start=1):
            text = td.get_text(strip=True)

            # 5의 배수 번째 td는 href도 함께 추출
            if idx % 5 == 0:
                a_tag = td.find("a")
                href = a_tag['href'] if a_tag and a_tag.has_attr('href') else None
                row_data[f"col_{idx}_text"] = text
                row_data[f"col_{idx}_href"] = href
            else:
                row_data[f"col_{idx}_text"] = text

        if row_data:
            data_rows.append(row_data)

# DataFrame 생성
df = pd.DataFrame(data_rows)

# 엑셀로 저장
df.to_excel("national_certList.xlsx", index=False)
print("✅ national_certList.xlsx 파일 저장 완료!")
