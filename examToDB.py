import os
import re
import pandas as pd
from sqlalchemy import create_engine, text

# ocr-env_db\Scripts\activate
# python examToDB.py

# 경로: docx 파일들이 있는 폴더
folder_path = "기출문제Docx"

# DB 연결 설정
db_url = "mysql+pymysql://root@localhost:3306/project_25_05"
engine = create_engine(db_url)

# certificate 테이블에서 name → id 매핑 가져오기
with engine.connect() as conn:
    cert_rows = conn.execute(text("SELECT id, name FROM certificate")).mappings().fetchall()
    cert_map = {row['name']: row['id'] for row in cert_rows}

# docx 파일 목록 가져오기
filenames = [f for f in os.listdir(folder_path) if f.endswith(".docx")]

exam_data = []

for filename in filenames:
    # 예: 가스기사20200606.docx → 한글이름, 날짜 분리
    match = re.match(r"([가-힣]+)(\d{8})\.docx", filename)
    if not match:
        print(f"⚠️ 스킵됨: 형식 불일치 - {filename}")
        continue

    cert_name, date_str = match.groups()
    exam_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:]}"  # YYYY-MM-DD

    cert_id = cert_map.get(cert_name)
    if not cert_id:
        print(f"❌ 자격증명 매핑 실패: {cert_name}")
        continue

    exam_data.append({
        "certId": cert_id,
        "category": "필기",
        "examDate": exam_date
    })

# DB 저장
with engine.connect() as conn:
    for row in exam_data:
        stmt = text("""
            INSERT INTO exam (certId, category, examDate, regDate, updateDate)
            VALUES (:certId, :category, :examDate, NOW(), NOW())
        """)
        conn.execute(stmt, row)
    conn.commit()

print("✅ exam 테이블 저장 완료!")
