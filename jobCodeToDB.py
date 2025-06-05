import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime

# 엑셀 파일 읽기
df = pd.read_excel('national_cert.xlsx') 
print(df.columns)

# DB 연결
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

# 필수 컬럼만 추출 (id, name, certGrade, isNational, agency, parentId)
df_filtered = df[['id', 'name', 'certGrade', 'isNational', 'agency', 'parentId']].dropna(subset=['id', 'name'])

# 현재 시간 (regDate, updateDate용)
now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

with engine.connect() as conn:
    for _, row in df_filtered.iterrows():
        stmt = text("""
            INSERT INTO certificate (id, name, certGrade, isNational, agency, parentId, regDate, updateDate)
            VALUES (:id, :name, :certGrade, :isNational, :agency, :parentId, NOW(), NOW())
        """)
        conn.execute(stmt, {
            "id": int(row["id"]),
            "name": row["name"],
            "certGrade": int(row["certGrade"]) if not pd.isna(row["certGrade"]) else None,
            "isNational": int(row["isNational"]) if not pd.isna(row["isNational"]) else None,
            "agency": row["agency"] if not pd.isna(row["agency"]) else None,
            "parentId": int(row["parentId"]) if not pd.isna(row["parentId"]) else None,
        })
    conn.commit()

print("데이터베이스에 저장 완료!")