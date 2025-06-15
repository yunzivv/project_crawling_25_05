import fitz  # PyMuPDF
import re
import pytesseract
from PIL import Image
import io
import cv2
import numpy as np
from collections import defaultdict
from docx import Document

# ocr-env_examToDB\Scripts\activate

doc = Document("가스기사20200606.docx")

subject_re = re.compile(r'(\d+)\s*과목\s*[:\-]?\s*(.+)')
question_re = re.compile(r'^(\d{1,3})\.\s*(.+)')
choice_re = re.compile(r'([①②③④❶❷❸❹])\s*(.+)')
label_map = {'①': 1, '②': 2, '③': 3, '④': 4, '❶': 1, '❷': 2, '❸': 3, '❹': 4}

questions = []
current_subject = "알 수 없음"
current_question = None

for para in doc.paragraphs:
    text = para.text.strip()
    if not text:
        continue

    # 과목명
    subj_match = subject_re.match(text)
    if subj_match:
        current_subject = subj_match.group(2).strip()
        continue

    # 문제 시작
    q_match = question_re.match(text)
    if q_match:
        if current_question:
            questions.append(current_question)
        current_question = {
            "number": int(q_match.group(1)),
            "body": q_match.group(2).strip(),
            "subject": current_subject,
            "choices": [],
            "correct": None
        }
        continue

    # 보기
    for c_match in choice_re.finditer(text):
        label = label_map.get(c_match.group(1))
        current_question["choices"].append({
            "label": label,
            "body": c_match.group(2).strip()
        })

# 마지막 문제 저장
if current_question:
    questions.append(current_question)

# 정답 추출 (문서 마지막 정답표)
answers = []
for para in reversed(doc.paragraphs):
    line = para.text.strip()
    if re.match(r'^\d{1,3}$', line) or re.match(r'^[①②③④❶❷❸❹\s]+$', line):
        answers.insert(0, line)
    if len(answers) > 3:  # 예상 정답 줄 수 확보되면 중단
        break

# 정답 매핑
answer_labels = re.findall(r'[①②③④❶❷❸❹]', " ".join(answers))
for i, q in enumerate(questions):
    if i < len(answer_labels):
        correct = label_map[answer_labels[i]]
        q["correct"] = correct
        for c in q["choices"]:
            c["isCorrect"] = (c["label"] == correct)
    else:
        for c in q["choices"]:
            c["isCorrect"] = False

# 검증
print(f"총 문제 수: {len(questions)}")
for q in questions[:5]:
    print(f"{q['number']}번 ({q['subject']}): {q['body']}")
    for c in q["choices"]:
        mark = "✅" if c.get("isCorrect") else ""
        print(f"  - {c['label']}: {c['body']} {mark}")


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