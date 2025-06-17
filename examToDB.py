import re
import os
import difflib
import requests
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.document import Document as _Document
from docx.table import _Cell
# ocr-env_examToDB\Scripts\activate
# python examToDB.py

# 문서 블록 순회
def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Unsupported parent")
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

# 제목 정보 추출
def extract_title_info(filename):
    basename = os.path.basename(filename)
    name, _ = os.path.splitext(basename)
    match = re.match(r'(.+?)(\d{8})$', name)
    if match:
        return match.group(1).strip(), match.group(2)
    return name, None

# 정답 테이블 추출
def extract_answers_from_table(table):
    answers = {}
    rows = table.rows
    for i in range(0, len(rows), 2):
        if i + 1 >= len(rows):
            break
        q_nums = [cell.text.strip() for cell in rows[i].cells]
        q_ans = [cell.text.strip() for cell in rows[i + 1].cells]
        for q, a in zip(q_nums, q_ans):
            if q.isdigit():
                answers[int(q)] = a
    return answers

# 문제 시작 위치 탐색
def find_question_start(paragraphs, q_text, from_idx):
    q_start_text = q_text[:20].replace(" ", "")
    best_score = 0
    best_idx = -1
    for i in range(from_idx, len(paragraphs)):
        para_text = paragraphs[i].text.replace(" ", "")
        score = difflib.SequenceMatcher(None, q_start_text, para_text[:len(q_start_text)]).ratio()
        if score > best_score:
            best_score = score
            best_idx = i
        if score > 0.85:
            return i
    return best_idx if best_score > 0.6 else -1

# 이미지 업로드
def upload_image_to_imgur(image_bytes):
    CLIENT_ID = '00ff8e726eb9eb8'
    url = "https://api.imgur.com/3/image"
    headers = {'Authorization': f'Client-ID {CLIENT_ID}'}
    response = requests.post(url, headers=headers, files={"image": image_bytes})
    if response.status_code == 200:
        return response.json()['data']['link']
    return None

# 선택지 추출
def extract_choices(text):
    split_choices = re.split(r"(①|②|③|④|❶|❷|❸|❹)", text)
    choices = []
    for i in range(1, len(split_choices) - 1, 2):
        choices.append({
            "number": (i // 2) + 1,
            "text": split_choices[i + 1].strip(),
            "has_image": False,
            "image_url": None
        })
    while len(choices) < 4:
        choices.append({"number": len(choices) + 1, "text": "", "has_image": False, "image_url": None})
    return choices

# 문제와 선택지 분리
def split_question_and_choices(text):
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if re.search(r"[①②③④❶❷❸❹]", line):
            return " ".join(lines[:i]), extract_choices(" ".join(lines[i:]))
    return text, extract_choices("")

# 본문 파싱
def parse_exam(doc):
    blocks = list(iter_block_items(doc))
    paragraphs = []
    para_to_index = {}

    for idx, b in enumerate(blocks):
        if isinstance(b, Paragraph):
            paragraphs.append(b)
            para_to_index[id(b)] = idx
        elif isinstance(b, Table):
            for row in b.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        paragraphs.append(para)
                        para_to_index[id(para)] = idx

    print(f"📄 전체 문단 수: {len(paragraphs)}")

    # 과목 추출
    subjects = []
    for idx, b in enumerate(blocks):
        if isinstance(b, Table) and len(b.rows) == 1 and len(b.rows[0].cells) == 1:
            text = b.rows[0].cells[0].text.strip()
            print(f"🔍 과목 후보 텍스트: '{text}'")
            m = re.match(r"^(\d)과목\s*[:：]\s*(.+)$", text)
            if m:
                print(f"✅ 과목 인식 성공: {text}")
                subjects.append((int(m.group(1)), m.group(2).strip(), idx))

    if not subjects:
        print("❌ 과목을 찾지 못했습니다. 파싱이 실패했을 수 있습니다.")
        return {"subjects": []}

    # 마지막 테이블을 정답표로 사용
    answer_table = None
    for b in reversed(blocks):
        if isinstance(b, Table):
            answer_table = b
            break

    answers = extract_answers_from_table(answer_table) if answer_table else {}

    subject_starts = [s[2] for s in subjects] + [len(blocks)]  # 끝 인덱스 포함
    data = {"subjects": []}

    for i in range(len(subjects)):
        subj_num, subj_name, start_idx = subjects[i]
        end_idx = subject_starts[i + 1]
        subj_blocks = blocks[start_idx:end_idx]

        questions = []
        current_q_num = None
        current_text = ""

        for b in subj_blocks:
            if isinstance(b, Paragraph):
                text = b.text.strip()
                if not text:
                    continue
                bold = any(run.bold for run in b.runs if run.text.strip())
                is_question = re.match(r"^\d+[.)]", text)
                if bold and is_question:
                    if current_q_num:
                        q_text, choices = split_question_and_choices(current_text)
                        questions.append({
                            "question_number": current_q_num,
                            "question_text": q_text,
                            "choices": choices,
                            "question_has_image": False,
                            "question_image_url": None,
                            "answer": answers.get(current_q_num, '')
                        })
                    current_q_num = int(is_question.group(0)[:-1])
                    current_text = text
                else:
                    current_text += "\n" + text

        if current_q_num:
            q_text, choices = split_question_and_choices(current_text)
            questions.append({
                "question_number": current_q_num,
                "question_text": q_text,
                "choices": choices,
                "question_has_image": False,
                "question_image_url": None,
                "answer": answers.get(current_q_num, '')
            })

        data["subjects"].append({
            "subject_number": subj_num,
            "subject_name": subj_name,
            "questions": questions
        })

    return data


# 요약 출력
def print_exam_summary(data):
    for subj in data['subjects']:
        print(f"\n📘 과목: {subj['subject_number']}과목 : {subj['subject_name']} - 총 {len(subj['questions'])}문제")
        for q in subj['questions'][8:11]:  # 처음 3문제만 확인
            print(f"  - {q['question_number']}번: {q['question_text'][:50]}... (정답: {q['answer']}, 이미지: {'O' if q['question_has_image'] else 'X'})")

# 메인 실행
def main(path):
    title, date = extract_title_info(path)
    print(f"\n📄 문서: {os.path.basename(path)}")
    doc = Document(path)
    exam_data = parse_exam(doc)

    # ✅ subjects가 비었는지 확인
    if not exam_data["subjects"]:
        print("❌ 과목을 찾지 못했습니다. 파싱이 실패했을 수 있습니다.")
        return

    print_exam_summary(exam_data)

if __name__ == "__main__":
    main("가스기사20200606.docx")