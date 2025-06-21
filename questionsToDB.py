# ocr-env_db\Scripts\activate
# python questionsToDB.py

import pandas as pd
import numpy as np
from sqlalchemy import create_engine, text

# 엑셀 파일 읽기
df = pd.read_excel('questions.xlsx')
print("📄 엑셀 컬럼:", df.columns.tolist())

# 문자열 전처리: 공백 제거 및 NaN 처리
df['certName'] = df['certName'].astype(str).str.strip().replace({'': pd.NA, 'nan': pd.NA})
df['subjectName'] = df['subjectName'].astype(str).str.strip().replace({'': pd.NA, 'nan': pd.NA})

# DB 연결
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

# 매핑 준비
with engine.connect() as conn:
    cert_rows = conn.execute(text("SELECT id, name FROM certificate")).mappings().fetchall()
    cert_map = {row['name']: row['id'] for row in cert_rows}

    subject_rows = conn.execute(text("SELECT id, certId, name FROM certSubject")).mappings().fetchall()
    subject_map = {(row['certId'], row['name']): row['id'] for row in subject_rows}

# certId, subjectId 매핑
df['certId'] = df['certName'].map(cert_map)
df['subjectId'] = df.apply(lambda r: subject_map.get((r['certId'], r['subjectName'])), axis=1)

# 매핑 실패 경고 및 필터링
missing_cert = df[df['certId'].isna()]
if not missing_cert.empty:
    print("❌ 매핑되지 않은 certName:")
    print(missing_cert[['certName']].drop_duplicates())

missing_subject = df[df['subjectId'].isna()]
if not missing_subject.empty:
    print("❌ 매핑되지 않은 subjectName:")
    print(missing_subject[['certName', 'subjectName']].drop_duplicates())

# 유효한 데이터만 필터링
df = df[~df['certId'].isna() & ~df['subjectId'].isna()]

# 필요한 컬럼만 추출
df_filtered = df[['id', 'certId', 'examId', 'subjectId', 'questNum', 'body', 'hasImage', 'imgUrl']]
df_filtered = df_filtered.dropna(subset=['id', 'certId', 'examId', 'subjectId', 'questNum', 'body', 'hasImage'])

# INSERT
with engine.begin() as conn:
    for _, row in df_filtered.iterrows():
        stmt = text("""
            INSERT INTO questions (
                id, certId, examId, subjectId, questNum, body, hasImage, imgUrl, regDate, updateDate
            ) VALUES (
                :id, :certId, :examId, :subjectId, :questNum, :body, :hasImage, :imgUrl, NOW(), NOW()
            )
        """)
        conn.execute(stmt, {
            "id": int(row["id"]),
            "certId": int(row["certId"]),
            "examId": int(row["examId"]),
            "subjectId": int(row["subjectId"]),
            "questNum": int(row["questNum"]),
            "body": str(row["body"]),
            "hasImage": bool(row["hasImage"]),
            "imgUrl": None if pd.isna(row["imgUrl"]) else str(row["imgUrl"])
        })

print("✅ questions 테이블에 저장 완료!")
