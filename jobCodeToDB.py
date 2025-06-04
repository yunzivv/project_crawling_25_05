import pandas as pd
from sqlalchemy import create_engine, text

# 엑셀 파일 읽기
df = pd.read_excel('jobkorea_jobCode.xlsx') 
print(df.columns)

# DB 연결
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

df_filtered = df[['id', 'name']].dropna(subset=['id', 'name'])

# 데이터 저장
with engine.connect() as conn:
    for _, row in df_filtered.iterrows():
        stmt = text("""
            INSERT INTO jobCat (id, name)
            VALUES (:id, :name)
        """)
        conn.execute(stmt, {
            "id": row["id"],
            "name": row["name"]
        })
    conn.commit()

print("데이터베이스에 저장 완료!")