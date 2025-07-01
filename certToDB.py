import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime

# ocr-env_db\Scripts\activate
# python certToDB.py

# 엑셀 파일 읽기
df = pd.read_excel('certList.xlsx') 
print(df.columns)

# DB 연결
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

# 필수 컬럼만 추출 (id, name, certGrade, isNational, agency, parentId)
df_filtered = df[['id', 'name', 'certGrade', 'isNational', 'agency', 'parentId', 'href']].dropna(subset=['id', 'name'])

# ✅ ID가 884보다 큰 데이터만 필터링
# df_filtered = df_filtered[df_filtered['id'] > 884]

# 현재 시간 (regDate, updateDate용)
now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

with engine.connect() as conn:
    for _, row in df_filtered.iterrows():
        row_id = int(row["id"])

        # ✅ 중복 여부 확인
        check_stmt = text("SELECT COUNT(*) FROM certificate WHERE id = :id")
        result = conn.execute(check_stmt, {"id": row_id}).scalar()

        if result > 0:
            print(f"⚠️ 중복된 ID {row_id} 건너뜀")
            continue 

        # ✅ 데이터 삽입
        stmt = text("""
            INSERT INTO certificate (id, name, certGrade, isNational, agency, parentId, href, regDate, updateDate)
            VALUES (:id, :name, :certGrade, :isNational, :agency, :href, :parentId, NOW(), NOW())
        """)
        conn.execute(stmt, {
            "id": row_id,
            "name": row["name"],
            "certGrade": int(row["certGrade"]) if not pd.isna(row["certGrade"]) else None,
            "isNational": int(row["isNational"]) if not pd.isna(row["isNational"]) else None,
            "agency": row["agency"] if not pd.isna(row["agency"]) else None,
            "parentId": int(row["parentId"]) if not pd.isna(row["parentId"]) else None,
            "href": row["href"] if not pd.isna(row["href"]) else None
        })

    conn.commit()


print("데이터베이스에 저장 완료")