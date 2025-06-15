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

# --- 설정 ---
FILE = "가스기사20200606.docx"

subject_re = re.compile(r'(\d+)\s*과목\s*[:\-]?\s*(.+)')
question_re = re.compile(r'^(\d{1,3})\.\s*(.+)')
choice_re = re.compile(r'(①|②|③|④|❶|❷|❸|❹)\s*([^①②③④❶❷❸❹]+)')
label_map = {'①': 1, '②': 2, '③': 3, '④': 4, '❶': 1, '❷': 2, '❸': 3, '❹': 4}

# --- 문서 열기 ---
doc = Document(FILE)
all_paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
full_text = "\n".join(all_paragraphs)

# --- 1. 과목 추출 (문서 전체에서) ---
subject_positions = []  # [(문자 인덱스, 과목명)]
for m in subject_re.finditer(full_text):
    subject_positions.append((m.start(), m.group(2).strip()))
subject_positions.sort()

# --- 2. 문제/보기 추출 ---
questions = []
current_question = None

# 전체 라인 순회
for idx, line in enumerate(all_paragraphs):
    line = line.strip()
    if not line:
        continue

    # 문제
    q_match = question_re.match(line)
    if q_match:
        if current_question:
            questions.append(current_question)
        current_question = {
            "number": int(q_match.group(1)),
            "body": q_match.group(2).strip(),
            "choices": [],
            "subject": "알 수 없음",
            "correct": None
        }
        continue

    # 보기
    if current_question:
        for c_match in choice_re.finditer(line):
            label = label_map[c_match.group(1)]
            body = c_match.group(2).strip()
            current_question["choices"].append({
                "label": label,
                "body": body
            })
        # 본문 계속 추가
        if not question_re.match(line) and not choice_re.search(line):
            current_question["body"] += " " + line

# 마지막 문제 저장
if current_question:
    questions.append(current_question)

# --- 3. 정답 추출 ---
answer_lines = []
for para in reversed(doc.paragraphs):
    text = para.text.strip()
    if re.search(r'[①②③④❶❷❸❹]', text):
        answer_lines.insert(0, text)
    if len(answer_lines) > 2:
        break

answer_symbols = re.findall(r'[①②③④❶❷❸❹]', " ".join(answer_lines))
for i, q in enumerate(questions):
    if i < len(answer_symbols):
        correct = label_map[answer_symbols[i]]
        q["correct"] = correct
        for c in q["choices"]:
            c["isCorrect"] = (c["label"] == correct)

# --- 4. 과목 ↔ 문제 매핑 ---
# 각 과목은 전체 텍스트 기준 위치가 있고, 각 문제는 문장 시작 단어 기준 위치로 추정 가능
line_positions = []
char_idx = 0
for line in all_paragraphs:
    line_positions.append((char_idx, line))
    char_idx += len(line) + 1  # +1 for newline

def find_subject_for_line(text_line):
    pos = full_text.find(text_line)
    selected = "알 수 없음"
    for i in range(len(subject_positions)):
        if pos >= subject_positions[i][0]:
            selected = subject_positions[i][1]
        else:
            break
    return selected

# 각 문제에 과목 할당
for q in questions:
    body_line = q["body"].split()[0] if q["body"] else ""
    q["subject"] = find_subject_for_line(body_line)

# --- 5. 출력 요약 ---
print(f"총 문제 수: {len(questions)}\n")

# 과목별 개수
subject_count = defaultdict(int)
for q in questions:
    subject_count[q['subject']] += 1
print("[과목별 문제 수]")
for s, cnt in subject_count.items():
    print(f"- {s}: {cnt}문제")

# 보기 부족
incomplete = [q for q in questions if len(q['choices']) != 4]
if incomplete:
    print(f"\n⚠️ 보기 4개가 아닌 문제 수: {len(incomplete)}")
    for q in incomplete[:5]:
        print(f"  - {q['number']}번 ({q['subject']}), 보기 수: {len(q['choices'])}")
else:
    print("\n✅ 모든 문제가 보기 4개")

# 정답 누락
no_answer = [q for q in questions if not any(c.get("isCorrect") for c in q['choices'])]
if no_answer:
    print(f"\n⚠️ 정답 누락 문제 수: {len(no_answer)}")
    print("  문제 번호 예시:", [q["number"] for q in no_answer[:10]])
else:
    print("\n✅ 모든 문제에 정답 포함")

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