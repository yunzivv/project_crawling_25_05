import re, os
import pymysql
from pyhwp import HWPDocument

# ocr-env_examToDB\Scripts\activate

# DB 연결
conn = pymysql.connect(
    host='localhost', user='your_user',
    password='your_pw', db='exam_db', charset='utf8'
)
cursor = conn.cursor()

def get_or_create_certificate(cert_name):
    # 1) 조회
    cursor.execute(
        "SELECT id FROM certificate WHERE `name`=%s",
        (cert_name,)
    )
    row = cursor.fetchone()
    if row:
        return row[0]
    # 2) 없으면 삽입
    cursor.execute(
        """
        INSERT INTO certificate
          (`name`, certGrade, isNational, agency,
           parentId, href, regDate, updateDate)
        VALUES
          (%s, NULL, NULL, NULL, NULL, NULL, NOW(), NOW())
        """,
        (cert_name,)
    )
    conn.commit()
    return cursor.lastrowid

# 파일명에서 자격증명·날짜 추출
def parse_filename(fname):
    m = re.match(r'(.+?)(\d{4})(\d{2})(\d{2})\.hwp$', fname)
    cert, y, mth, d = m.groups()
    return cert, f"{y}-{mth}-{d}"

# subjects 테이블 UPSERT 함수
subject_pattern = re.compile(r'^\s*\d+과목\s*[:\-]\s*(.+)$')
def extract_subjects(hwp_path):
    doc = HWPDocument(hwp_path)
    subs = set()
    for para in doc.bodytext.paragraph_list:
        text = ''.join(run.text for run in para.run_list).strip()
        m = subject_pattern.match(text)
        if m:
            subs.add(m.group(1).strip())
    return subs

def upsert_subject(cert_id, subject_name):
    cursor.execute(
        "SELECT id FROM subjects WHERE cert_id=%s AND name=%s",
        (cert_id, subject_name)
    )
    row = cursor.fetchone()
    if row:
        return row[0]
    cursor.execute(
        "INSERT INTO subjects (cert_id, name) VALUES (%s, %s)",
        (cert_id, subject_name)
    )
    conn.commit()
    return cursor.lastrowid

# --- 실행 예시 ---
hwp_file = '가스기사20200606.hwp'
cert_name, session_date = parse_filename(hwp_file)

# 1) 자격증 id 확보
cert_id = get_or_create_certificate(cert_name)
print("cert_id:", cert_id)

# 2) 시험회차(exam_sessions) INSERT
cursor.execute(
    "INSERT INTO exam_sessions (cert_id, session_date) VALUES (%s, %s)",
    (cert_id, session_date)
)
conn.commit()
session_id = cursor.lastrowid
print("session_id:", session_id)

# 3) 과목(subjects) 추출·저장
subjects = extract_subjects(hwp_file)
subject_ids = {}
for name in subjects:
    sid = upsert_subject(cert_id, name)
    subject_ids[name] = sid
print("subject_ids:", subject_ids)
