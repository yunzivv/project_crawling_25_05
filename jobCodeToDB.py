import pandas as pd
from sqlalchemy import create_engine, text

# 엑셀 파일 읽기
df = pd.read_excel('jobkorea_jobCode.xlsx') 
print(df.columns)

# DB 연결
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

df_filtered = df[['jobCatId', 'id', 'name']].dropna(subset=['jobCatId', 'id', 'name'])

# 데이터 저장
with engine.connect() as conn:
    for _, row in df_filtered.iterrows():
        stmt = text("""
            INSERT INTO jobCode (jobCatId, id, name)
            VALUES (:jobCatId, :id, :name)
        """)
        conn.execute(stmt, {
            "jobCatId": row["jobCatId"],
            "id": row["id"],
            "name": row["name"]
        })
    conn.commit()

print("DB 저장 완료")