import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime

# ocr-env_db
# 엑셀 읽기
df = pd.read_excel('jobkorea_jobCode DB 데이터.xlsx')

print(df.columns)  # 컬럼 구조 확인

# DB 연결
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

# 필요한 컬럼만 추출
df_filtered = df[['jobCatId', 'jobCatName', 'code', 'name']].dropna()

# 데이터 저장
with engine.connect() as conn:
    for _, row in df_filtered.iterrows():
        stmt = text("""
            INSERT INTO jobCode (jobCatId, jobCatName, code, name, regDate, updateDate)
            VALUES (:jobCatId, :jobCatName, :code, :name, NOW(), NOW())
        """)
        conn.execute(stmt, {
            "jobCatId": int(row["jobCatId"]),
            "jobCatName": row["jobCatName"],
            "code": int(row["code"]),
            "name": row["name"]
        })
    conn.commit()

print("데이터베이스에 저장 완료!")
