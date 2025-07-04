import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime

# ocr-env_db\Scripts\activate
# python jobCatToDB.py

# 엑셀 읽기
df = pd.read_excel('jobkorea_jobCat.xlsx')

print(df.columns)  # 컬럼 구조 확인

# DB 연결
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

# 필요한 컬럼만 추출
df_filtered = df[['id', 'name']].dropna()

# 데이터 저장
with engine.connect() as conn:
    for _, row in df_filtered.iterrows():
        stmt = text("""
            INSERT INTO jobCat (id, name, regDate, updateDate)
            VALUES (:id, :name, NOW(), NOW())
        """)
        conn.execute(stmt, {
            "id": int(row["id"]),
            "name": row["name"]
        })
    conn.commit()

print("데이터베이스에 저장 완료!")
