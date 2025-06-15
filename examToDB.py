import fitz  # PyMuPDF
import re, os, datetime
from pathlib import Path
from collections import defaultdict

# ocr-env_examToDB\Scripts\activate

# 파일 경로 설정
PDF_PATH = "가스기사20200606.pdf"
IMG_DIR = "images"

# 정규표현식 정의
subject_re = re.compile(r'^\s*(\d+)과목\s*[:\-]\s*(.+)')
question_re = re.compile(r'^(\d{1,3})\.\s*(.*)')
choice_re = re.compile(r'([①②③④❶❷❸❹])\s*([^①②③④❶❷❸❹]+)')

# 보기를 숫자로 매핑
label_map = {
    "①": 1, "②": 2, "③": 3, "④": 4,
    "❶": 1, "❷": 2, "❸": 3, "❹": 4
}

# 정답지 (페이지 마지막에 수작업으로 추출했음)
answer_keys = [
    2, 3, 3, 3, 1, 1, 1, 1, 2, 4,
    2, 2, 1, 4, 1, 4, 3, 4, 3, 1,
    2, 3, 2, 2, 4, 1, 1, 4, 1, 4,
    1, 2, 3, 2, 4, 2, 3, 2, 3, 3,
    2, 2, 1, 3, 1, 2, 2, 1, 4, 3,
    1, 4, 1, 4, 2, 2, 1, 4, 4, 3,
    2, 2, 1, 3, 4, 1, 1, 2, 2, 2,
    1, 4, 3, 1, 1, 3, 3, 2, 3, 1,
    2, 2, 2, 3, 3, 1, 4, 3, 4, 4,
    1, 2, 2, 3, 3, 1, 4, 3, 1, 1
]

# 인증 정보
cert_name = "가스기사"
exam_date = "2020-06-06"

# 결과 저장용
questions = []
current_subject = "알 수 없음"
current_question = None

# PDF 열기
doc = fitz.open(PDF_PATH)

# 1. 텍스트 추출 및 파싱
for page in doc:
    lines = page.get_text("text").splitlines()
    for line in lines:
        line = line.strip()
        if not line:
            continue

        # 과목 감지
        subj_match = subject_re.match(line)
        if subj_match:
            current_subject = subj_match.group(2).strip()
            continue

        # 문제 감지
        q_match = question_re.match(line)
        if q_match:
            # 기존 문제 저장
            if current_question:
                questions.append(current_question)

            current_question = {
                "number": int(q_match.group(1)),
                "body": q_match.group(2).strip(),
                "subject": current_subject,
                "choices": []
            }
            continue

        # 보기 감지 (여러 개일 수 있음)
        if current_question:
            for c_match in choice_re.finditer(line):
                label_str = c_match.group(1)
                label_num = label_map.get(label_str)
                body = c_match.group(2).strip()

                current_question["choices"].append({
                    "label": label_num,
                    "body": body
                })

# 마지막 문제 저장
if current_question:
    questions.append(current_question)

# 2. 정답 매핑
for q in questions:
    try:
        correct_label = answer_keys[q["number"] - 1]
        for c in q["choices"]:
            c["isCorrect"] = (c["label"] == correct_label)
    except IndexError:
        print(f"정답 없음: {q['number']}번")

# 3. 이미지 추출 (옵션)
Path(IMG_DIR).mkdir(exist_ok=True)
for i, page in enumerate(doc):
    for img_index, img in enumerate(page.get_images(full=True)):
        base_image = doc.extract_image(img[0])
        ext = base_image['ext']
        image_bytes = base_image['image']
        image_path = os.path

# 1. 총 문제 수
print(f"총 문제 수: {len(questions)}")

# 2. 과목별 문제 수
subject_count = defaultdict(int)
for q in questions:
    subject_count[q['subject']] += 1

print("\n[과목별 문제 수]")
for subject, count in subject_count.items():
    print(f"- {subject}: {count}문제")

# 3. 보기 개수 이상/이하 확인
incomplete_choices = [q for q in questions if len(q['choices']) != 4]
if incomplete_choices:
    print(f"\n⚠️ 보기 개수가 4개가 아닌 문제 수: {len(incomplete_choices)}")
    for q in incomplete_choices[:5]:  # 앞부분만 예시 출력
        print(f"  - {q['number']}번 ({q['subject']}), 보기 수: {len(q['choices'])}")
else:
    print("\n✅ 모든 문제가 4개의 보기를 가짐")

# 4. 정답 정보 누락 확인
no_answer = []
for q in questions:
    has_correct = any(c.get("isCorrect") for c in q['choices'])
    if not has_correct:
        no_answer.append(q['number'])

if no_answer:
    print(f"\n⚠️ 정답 정보 누락된 문제 수: {len(no_answer)}")
    print("  누락 문제 번호:", no_answer[:10], "...")
else:
    print("\n✅ 모든 문제에 정답 매핑 완료")

# 5. 샘플 문제 1~3개 출력
print("\n[샘플 문제 출력]")
for q in questions[:11]:
    print(f"{q['number']}번 ({q['subject']}): {q['body']}")
    for c in q['choices']:
        mark = "✅" if c.get("isCorrect") else ""
        print(f"  - {c['label']}: {c['body']} {mark}")