import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime

# ocr-env_db\Scripts\activate
# python questionsToDB.py

# 엑셀 파일 읽기
df = pd.read_excel('questions.xlsx')

# INSERT용 컬럼 확인
print("📄 엑셀 컬럼:", df.columns.tolist())

db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

# certificate 테이블에서 certName → id 매핑 가져오기
with engine.connect() as conn:
    cert_rows = conn.execute(text("SELECT id, name FROM certificate")).mappings().fetchall()
    cert_map = {row['name']: row['id'] for row in cert_rows}

    subject_rows = conn.execute(text("SELECT id, name FROM certSubject")).mappings().fetchall()
    subject_map = {row['name']: row['id'] for row in subject_rows}

    
# certId 매핑
df['certId'] = df['certName'].map(cert_map)
df['subjectId'] = df['subjectName'].map(subject_map)

# 누락된 certName 확인
missing = df[df['certId'].isna()]
if not missing.empty:
    print("❌ 매핑되지 않은 certName:")
    print(missing['certName'].drop_duplicates())
    df = df[~df['certId'].isna()]  # 매핑된 것만 남김

missing = df[df['subjectId'].isna()]
if not missing.empty:
    print("❌ 매핑되지 않은 subjectName:")
    print(missing['subjectName'].drop_duplicates())
    df = df[~df['subjectId'].isna()]  # 매핑된 것만 남김


# 컬럼 정제 및 DB 저장
df_filtered = df[['id', 'certId', 'examId', 'subjectId', 'questNum', 'body', 'hasImage', 'imgUrl']].dropna()

with engine.begin()  as conn:
    for _, row in df_filtered.iterrows():
        stmt = text("""
            INSERT INTO questions (id, certId, examId, subjectId, questNum, 
                    body, hasImage, imgUrl, regDate, updateDate)
            VALUES (:id, :certId, :examId, :subjectId, :questNum, 
                    :body, :hasImage, :imgUrl, NOW(), NOW())
        """)
        conn.execute(stmt, {
            "id": int(row["id"]),
            "certId": int(row["certId"]),
            "examId": int(row["examId"]),
            "subjectId": int(row["subjectId"]),
            "questNum": int(row["questNum"]),
            "body": str(row["body"]),
            "hasImage": bool(row["hasImage"]),
            "imgUrl": str(row["imgUrl"])
        })
    conn.commit()

print("✅ DB 저장 완료!")