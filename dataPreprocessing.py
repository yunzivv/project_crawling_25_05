import pandas as pd

# ocr-env_preprocess\Scripts\activate
# python dataPreprocessing.py


# 데이터 불러오기
df_target = pd.read_excel("jobkorea_requirements.xlsx")
df_cert = pd.read_excel("certList.xlsx")     

# 2. certName 공백 제거
df_target['certName_clean'] = df_target['certName'].str.replace(r'\s+', '', regex=True)

# 3. 매핑 함수 정의
def find_matching_id(cert_name):
    for _, row in df_cert.iterrows():
        cleaned_name = str(row['name']).replace(" ", "")
        if cleaned_name in cert_name:
            return row['id']
    return None

# 4. 매핑 수행
df_target['certId'] = df_target['certName_clean'].apply(find_matching_id)

# 5. 임시 컬럼 제거
df_target.drop(columns=['certName_clean'], inplace=True)

# 결과 저장
print("매핑 완료")
df_target.to_excel("매핑.xlsx", index=False)
