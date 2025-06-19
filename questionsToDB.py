import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime

# ocr-env_db\Scripts\activate
# python questionsToDB.py

# 엑셀 파일 읽기
df = pd.read_excel('questionsToDBTest.xlsx')

print("📄 엑셀 컬럼:", df.columns.tolist())

# DB 연결
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

# 필요한 컬럼만 추출 (테이블 컬럼과 일치)
columns_needed = ['id', 'certId', 'examId', 'subjectId', 'questNum', 'body', 'hasImage', 'imgUrl']
df_filtered = df[columns_needed].dropna(subset=['id', 'examId', 'questNum', 'body'])

# INSERT 실행
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
            "certId": int(row["certId"]) if not pd.isna(row["certId"]) else None,
            "examId": int(row["examId"]),
            "subjectId": int(row["subjectId"]) if not pd.isna(row["subjectId"]) else None,
            "questNum": int(row["questNum"]),
            "body": str(row["body"]),
            "hasImage": bool(row["hasImage"]),
            "imgUrl": str(row["imgUrl"]) if not pd.isna(row["imgUrl"]) else None
        })

print("✅ questions 테이블에 저장 완료!")