import fitz  # PyMuPDF: pdf 읽기
import re # 정규식
import pytesseract # OCR
from PIL import Image
import io
import cv2 # 이미지 처리
import numpy as np
from collections import defaultdict

# ocr-env_examToDB\Scripts\activate

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
PDF_PATH = "가스기사20200606_3.pdf"

subject_re = re.compile(r'(\d+)\s*과목\s*[:\-]?\s*(.+)')
question_re = re.compile(r'^(\d{1,3})\.\s*(.*)')
choice_re = re.compile(r'([①②③④❶❷❸❹])\s*([^①②③④❶❷❸❹]+)')
label_map = {"①": 1, "②": 2, "③": 3, "④": 4, "❶": 1, "❷": 2, "❸": 3, "❹": 4}

answer_keys = [  # 문제 번호 순서
    2,3,3,3,1,1,1,1,2,4,
    2,2,1,4,1,4,3,4,3,1,
    2,3,2,2,4,1,1,4,1,4,
    2,3,2,2,4,1,1,4,1,4,
    2,2,3,2,4,2,3,2,3,3,
    2,2,1,3,1,2,2,1,4,3,
    1,4,1,4,2,2,1,4,4,3,
    2,2,1,3,4,1,1,2,1,2,
    1,4,1,1,4,3,3,2,3,4,
    2,3,2,3,3,1,4,3,4,4,
    1,2,2,2,3,3,4,3,1,1
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

# 이미지 처리 대상 필터링
for q in questions:
    if len(q["choices"]) < 4:
        ocr_needed.append(q)

# 이미지 저장 함수
def save_cropped_box_image(page, q_number):
    # 1. PDF 페이지를 이미지로 변환
    pix = page.get_pixmap(dpi=300)
    img = Image.open(io.BytesIO(pix.tobytes(output="png")))

    # 2. PIL → OpenCV 이미지로 변환
    img_cv = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)

    # 3. 그레이스케일 + 이진화 (검은 테두리 강조)
    gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

    # 4. 윤곽선(컨투어) 찾기
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    # 5. 네모 중 크기 큰 것 1개만 선택
    largest_box = None
    max_area = 0
    for cnt in contours:
        approx = cv2.approxPolyDP(cnt, 0.02 * cv2.arcLength(cnt, True), True)
        area = cv2.contourArea(cnt)

        if len(approx) == 4 and area > 10000 and area > max_area:
            x, y, w, h = cv2.boundingRect(cnt)
            largest_box = (x, y, w, h)
            max_area = area

    # 6. 박스를 잘라서 저장
    if largest_box:
        x, y, w, h = largest_box
        cropped = img_cv[y:y+h, x:x+w]
        save_path = f"cropped_q{q_number}_page{page.number}.png"
        cv2.imwrite(save_path, cropped)
        print(f"✅ 잘라낸 박스 저장됨: {save_path}")
    else:
        print(f"⚠️ {q_number}번: 박스 인식 실패. 전체 저장.")
        # fallback: 전체 저장
        full_path = f"full_q{q_number}_page{page.number}.png"
        pix.save(full_path)


# 이미지 저장
for q in ocr_needed:
    page = doc[q["page"]]
    save_cropped_box_image(page, q["number"])

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
# print("\n[샘플 문제 출력]")
# for q in questions[:10]:
#     print(f"{q['number']}번 ({q['subject']}): {q['body']}")
#     for c in q['choices']:
#         mark = "✅" if c.get("isCorrect") else ""
#         print(f"  - {c['label']}: {c['body']} {mark}")