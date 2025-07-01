import pandas as pd
from sqlalchemy import create_engine, text

# ocr-env_db\Scripts\activate
# python certMentionToDB.py

# 1. 엑셀 파일 읽기
df = pd.read_excel('매핑.xlsx')

# 2. certId가 null인 행 제외
df_filtered = df[df['certId'].notnull()].copy()

# 3. 정수형 컬럼 타입 변환 (오류 방지용)
df_filtered['certId'] = df_filtered['certId'].astype(int)
df_filtered['jobCatId'] = df_filtered['jobCatId'].astype(int)
df_filtered['jobCodeId'] = df_filtered['jobCodeId'].astype(int)
df_filtered['gno'] = df_filtered['gno'].astype(int)

# 4. 날짜 컬럼을 datetime 타입으로 변환 (형식 안 맞을 경우 오류 대비)
df_filtered['regDate'] = pd.to_datetime(df_filtered['regDate'], errors='coerce')
df_filtered['updateDate'] = pd.to_datetime(df_filtered['updateDate'], errors='coerce')

# 5. DB 연결
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

# 6. 데이터 INSERT
with engine.connect() as conn:
    for _, row in df_filtered.iterrows():
        stmt = text("""
            INSERT INTO certMention (
                id, jobCatId, jobCodeId, certId, gno, source, regDate, updateDate
            ) VALUES (
                :id, :jobCatId, :jobCodeId, :certId, :gno, :source, :regDate, :updateDate
            )
        """)
        conn.execute(stmt, {
            "id": row["id"],
            "jobCatId": row["jobCatId"],
            "jobCodeId": row["jobCodeId"],
            "certId": row["certId"],
            "gno": row["gno"],
            "source": row.get("source", "jobkorea"),
            "regDate": row["regDate"],
            "updateDate": row["updateDate"]
        })
    conn.commit()

print("✅ certMention 테이블에 데이터 저장 완료")