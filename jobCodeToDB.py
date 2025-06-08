# import pandas as pd
# from sqlalchemy import create_engine, text

# # ocr-env_db
# # 1. 엑셀 파일 읽기
# df = pd.read_excel('매핑완료.xlsx')

# # 2. certId가 null인 행 제외
# df_filtered = df[df['certId'].notnull()].copy()

# # 3. 정수형 컬럼 타입 변환 (오류 방지용)
# df_filtered['certId'] = df_filtered['certId'].astype(int)
# df_filtered['jobCatId'] = df_filtered['jobCatId'].astype(int)
# df_filtered['jobCodeId'] = df_filtered['jobCodeId'].astype(int)
# df_filtered['gno'] = df_filtered['gno'].astype(int)

# # 4. 날짜 컬럼을 datetime 타입으로 변환 (형식 안 맞을 경우 오류 대비)
# df_filtered['regDate'] = pd.to_datetime(df_filtered['regDate'], errors='coerce')
# df_filtered['updateDate'] = pd.to_datetime(df_filtered['updateDate'], errors='coerce')

# # 5. DB 연결
# db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
# engine = create_engine(db_url)

# # 6. 데이터 INSERT
# with engine.connect() as conn:
#     for _, row in df_filtered.iterrows():
#         stmt = text("""
#             INSERT INTO certMention (
#                 jobCatId, jobCodeId, certId, gno, source, regDate, updateDate
#             ) VALUES (
#                 :jobCatId, :jobCodeId, :certId, :gno, :source, :regDate, :updateDate
#             )
#         """)
#         conn.execute(stmt, {
#             "jobCatId": row["jobCatId"],
#             "jobCodeId": row["jobCodeId"],
#             "certId": row["certId"],
#             "gno": row["gno"],
#             "source": row.get("source", "jobkorea"),  # 없으면 기본값
#             "regDate": row["regDate"],
#             "updateDate": row["updateDate"]
#         })
#     conn.commit()

# print("✅ certMention 테이블에 데이터 저장 완료")

import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime

# ocr-env_db
# 엑셀 파일 읽기
df = pd.read_excel('national_cert.xlsx') 
print(df.columns)

# DB 연결
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

# 필수 컬럼만 추출 (id, name, certGrade, isNational, agency, parentId)
df_filtered = df[['id', 'name', 'certGrade', 'isNational', 'agency', 'parentId']].dropna(subset=['id', 'name'])

# ✅ ID가 884보다 큰 데이터만 필터링
df_filtered = df_filtered[df_filtered['id'] > 884]

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
            continue  # 이미 있는 ID면 insert 건너뜀

        # ✅ 데이터 삽입
        stmt = text("""
            INSERT INTO certificate (id, name, certGrade, isNational, agency, parentId, regDate, updateDate)
            VALUES (:id, :name, :certGrade, :isNational, :agency, :parentId, NOW(), NOW())
        """)
        conn.execute(stmt, {
            "id": row_id,
            "name": row["name"],
            "certGrade": int(row["certGrade"]) if not pd.isna(row["certGrade"]) else None,
            "isNational": int(row["isNational"]) if not pd.isna(row["isNational"]) else None,
            "agency": row["agency"] if not pd.isna(row["agency"]) else None,
            "parentId": int(row["parentId"]) if not pd.isna(row["parentId"]) else None,
        })

    conn.commit()


print("데이터베이스에 저장 완료")
