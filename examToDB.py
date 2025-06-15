import fitz  # PyMuPDF
import re
import pytesseract
from PIL import Image
import io
import cv2
import numpy as np
from collections import defaultdict

# ocr-env_examToDB\Scripts\activate

pytesseract.pytesseract.tesseract_cmd = r'D:\yunzi\Tesseract-OCR\tesseract.exe'
PDF_PATH = "가스기사20200606.pdf"

subject_re = re.compile(r'(\d+)\s*과목\s*[:\-]?\s*(.+)')
question_re = re.compile(r'^(\d{1,3})\.\s*(.*)')
choice_re = re.compile(r'([①②③④❶❷❸❹])\s*([^①②③④❶❷❸❹]+)')
label_map = {"①": 1, "②": 2, "③": 3, "④": 4, "❶": 1, "❷": 2, "❸": 3, "❹": 4}

answer_keys = [  # 문제 번호 순서
    2, 3, 3, 3, 1, 1, 1, 1, 2, 4, 2, 2, 1, 4, 1, 4, 3, 4, 3, 3,
    2, 3, 2, 2, 4, 1, 1, 4, 1, 4, 1, 2, 3, 2, 4, 2, 3, 2, 3, 3,
    2, 2, 1, 3, 1, 2, 2, 1, 4, 3, 1, 4, 1, 4, 2, 2, 1, 4, 4, 3,
    2, 2, 1, 3, 4, 1, 1, 2, 2, 2, 1, 4, 3, 1, 1, 3, 3, 2, 3, 1,
    2, 2, 2, 3, 3, 1, 4, 3, 1, 1, 2, 2, 2, 3, 3, 1, 4, 3, 1, 1
]

questions = []
current_question = None
current_subject = "알 수 없음"
seen_question_numbers = set()
ocr_needed = []

doc = fitz.open(PDF_PATH)

for page_index, page in enumerate(doc):
    lines = page.get_text("text").splitlines()
    for line in lines:
        line = line.strip()
        if not line:
            continue

        if "과목" in line:
            subj_match = subject_re.search(line)
            if subj_match:
                current_subject = subj_match.group(2).strip()
                continue

        q_match = question_re.match(line)
        if q_match:
            q_num = int(q_match.group(1))
            if q_num in seen_question_numbers:
                continue
            seen_question_numbers.add(q_num)

            if current_question:
                questions.append(current_question)

            current_question = {
                "number": q_num,
                "body": q_match.group(2).strip(),
                "subject": current_subject,
                "choices": [],
                "page": page_index
            }
            continue

        if current_question:
            for c_match in choice_re.finditer(line):
                label = label_map.get(c_match.group(1))
                body = c_match.group(2).strip()
                current_question["choices"].append({
                    "label": label,
                    "body": body
                })

        if current_question and not question_re.match(line) and not choice_re.search(line):
            current_question["body"] += " " + line.strip()

if current_question:
    questions.append(current_question)

# OCR 처리 대상 필터링
for q in questions:
    if len(q["choices"]) < 4:
        ocr_needed.append(q)

# OCR 추출 함수
def extract_choices_with_ocr(page, bbox=None):
    # 이미지로 페이지 추출
    pix = page.get_pixmap(dpi=300)
    img = Image.open(io.BytesIO(pix.tobytes()))
    if bbox:
        img = img.crop(bbox)

    img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
    ocr_text = pytesseract.image_to_string(img_cv, lang="eng+kor")
    results = []
    for match in choice_re.finditer(ocr_text):
        label = label_map.get(match.group(1))
        body = match.group(2).strip()
        results.append({
            "label": label,
            "body": body
        })
    return results

# OCR 수행
for q in ocr_needed:
    page = doc[q["page"]]
    new_choices = extract_choices_with_ocr(page)
    if new_choices:
        q["choices"] = new_choices

# 정답 매핑
for q in questions:
    try:
        correct_label = answer_keys[q["number"] - 1]
        for c in q["choices"]:
            c["isCorrect"] = (c["label"] == correct_label)
    except IndexError:
        for c in q["choices"]:
            c["isCorrect"] = False

# 검증 요약
summary = {
    "total_questions": len(questions),
    "subjects": {},
    "ocr_fixed": [q["number"] for q in ocr_needed if len(q["choices"]) == 4],
    "incomplete_choices": [q["number"] for q in questions if len(q["choices"]) != 4],
    "no_answer": [q["number"] for q in questions if not any(c.get("isCorrect") for c in q["choices"])],
    "samples": questions[:5]
}
for q in questions:
    summary["subjects"].setdefault(q["subject"], 0)
    summary["subjects"][q["subject"]] += 1

summary["incomplete_choices_count"] = len(summary["incomplete_choices"])
summary["no_answer_count"] = len(summary["no_answer"])
summary

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
for q in questions[:20]:
    print(f"{q['number']}번 ({q['subject']}): {q['body']}")
    for c in q['choices']:
        mark = "✅" if c.get("isCorrect") else ""
        print(f"  - {c['label']}: {c['body']} {mark}")