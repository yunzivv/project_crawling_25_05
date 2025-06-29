import requests
import xmltodict
import pandas as pd
import os


# ocr-env_api
# openAPIToDB.py
# ▶️ API 기본 정보
url = "http://openapi.q-net.or.kr/api/service/rest/InquiryQualInfo/getList"
params = {
    'serviceKey': 'oq+sPntu1thBj9w0wQqBXe6iKqSV+H2oHC/Tq2g0AYbAHmNQBT3e0GjkSjydagFYwkUrreBF47Ylrvozi99hxw==',
    'seriesCd': '04'
}

# ▶️ API 요청
response = requests.get(url, params=params)
data_dict = xmltodict.parse(response.content)

# ▶️ 오류 응답 체크
header = data_dict['response']['header']
if header['resultCode'] != '00':
    print(f"❌ API 오류: {header['resultMsg']}")
    exit()

# ▶️ 데이터 추출
items = data_dict['response']['body'].get('items', {}).get('item', [])

# ▶️ 단일 항목 처리
if isinstance(items, dict):
    items = [items]

# ✅ new_records는 항상 정의됨
new_records = []
for item in items:
    new_records.append({
        'career': item.get('career'),
        'implNm': item.get('implNm'),
        'instiNm': item.get('instiNm'),
        'jmNm': item.get('jmNm'),
        'job': item.get('job'),
        'mdobligFldNm': item.get('mdobligFldNm'),
        'seriesNm': item.get('seriesNm')
    })

# ▶️ 기존 파일과 합치기
file_path = 'qnet_certifications.xlsx'
if os.path.exists(file_path):
    existing_df = pd.read_excel(file_path)
    combined_df = pd.concat([existing_df, pd.DataFrame(new_records)], ignore_index=True)
    print("📌 기존 파일에 데이터 추가")
else:
    combined_df = pd.DataFrame(new_records)
    print("📄 새 파일 생성")

# ▶️ 저장
combined_df.to_excel(file_path, index=False)
print(f"✅ 저장 완료: {file_path}")