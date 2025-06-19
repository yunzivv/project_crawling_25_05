import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime


# ocr-env_db\Scripts\activate
# python subjectsToDB.py

# 엑셀 로드
df = pd.read_excel("subjects.xlsx")

# DB 연결
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

# certificate 테이블에서 certName → id 매핑 가져오기
with engine.connect() as conn:
    cert_rows = conn.execute(text("SELECT id, name FROM certificate")).mappings().fetchall()
    cert_map = {row['name']: row['id'] for row in cert_rows}

# certId 매핑
df['certId'] = df['certName'].map(cert_map)

# 누락된 certName 확인
missing = df[df['certId'].isna()]
if not missing.empty:
    print("❌ 매핑되지 않은 certName:")
    print(missing['certName'].drop_duplicates())
    df = df[~df['certId'].isna()]  # 매핑된 것만 남김

# 컬럼 정제 및 DB 저장
df_filtered = df[['certId', 'subjectNum', 'name']].dropna()

with engine.connect() as conn:
    for _, row in df_filtered.iterrows():
        stmt = text("""
            INSERT INTO certSubject (certId, subjectNum, name, regDate, updateDate)
            VALUES (:certId, :subjectNum, :name, NOW(), NOW())
        """)
        conn.execute(stmt, {
            "certId": int(row["certId"]),
            "subjectNum": int(str(row["subjectNum"]).replace("과목", "")),
            "name": row["name"]
        })
    conn.commit()

print("✅ DB 저장 완료!")