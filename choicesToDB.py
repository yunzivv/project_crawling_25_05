import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime

# ocr-env_db\Scripts\activate
# python choicesToDB.py

# 엑셀 읽기
df = pd.read_excel('choicesToDBTest.xlsx')

print(df.columns)  # 컬럼 구조 확인

# DB 연결
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

# 필요한 컬럼만 추출
df_filtered = df[['id', 'questId', 'label', 'body', 'isCorrect']].dropna()

# 데이터 저장
with engine.connect() as conn:
    for _, row in df_filtered.iterrows():
        stmt = text("""
            INSERT INTO choices (id, questId, label, body, isCorrect, regDate, updateDate)
            VALUES (:id, :questId, :label, :body, :isCorrect, NOW(), NOW())
        """)
        conn.execute(stmt, {
            "id": int(row["id"]),
            "questId": int(row["questId"]),
            "label": int(row["label"]),
            "body": str(row["body"]),
            "isCorrect": bool(row["isCorrect"])
        })
    conn.commit()

print("데이터베이스에 저장 완료!")
