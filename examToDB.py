import re
import os
import requests
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph

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

# 문제/선택지 분리
def split_question_and_choices(text):
    lines = text.splitlines()
    for i, line in enumerate(lines):
        if re.search(r"[①②③④❶❷❸❹]", line):
            return " ".join(lines[:i]), extract_choices(" ".join(lines[i:]))
    return text, extract_choices("")

# 본문 파싱
def parse_exam(doc):
    paragraphs = []
    for b in iter_block_items(doc):
        if isinstance(b, Paragraph):
            paragraphs.append(b)
        elif isinstance(b, Table):
            for row in b.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        paragraphs.append(para)

    blocks = list(iter_block_items(doc))
    tables = [b for b in blocks if isinstance(b, Table)]
    answer_table = tables[-1] if tables else None
    answers = extract_answers_from_table(answer_table) if answer_table else {}

    subjects = []
    for b in blocks:
        if isinstance(b, Table) and len(b.rows) == 1 and len(b.rows[0].cells) == 1:
            cell_text = b.rows[0].cells[0].text.strip()
            m = re.match(r"^(\d)과목\s*[:：]\s*(.+)$", cell_text)
            print(m)
            if m:
                for para in b.rows[0].cells[0].paragraphs:
                    para = b.rows[0].cells[0].paragraphs[0]
                    subjects.append((int(m.group(1)), m.group(2), para))

    subject_indices = []
    for (_, _, para) in subjects:
        for idx, p in enumerate(paragraphs):
            if p.text.strip() == para.text.strip():  # 텍스트로 비교
                subject_indices.append(idx)
                break
    subject_indices.append(len(paragraphs))

    data = {"subjects": []}
    for i in range(len(subjects)):
        print(i)
        number, name, _ = subjects[i]
        start = subject_indices[i]
        end = subject_indices[i + 1]

        questions = []
        current_q = None
        current_text = ""
        for p in paragraphs[start:end]:
            text = p.text.strip()
            if not text:
                continue
            bold = any(run.bold for run in p.runs if run.text.strip())
            is_q = re.match(r"^(\d+)\.\s", text)
            if bold and is_q:
                if current_q:
                    qt, choices = split_question_and_choices(current_text)
                    questions.append({
                        "question_number": current_q,
                        "question_text": qt,
                        "choices": choices,
                        "question_has_image": False,
                        "question_image_url": None,
                        "answer": answers.get(current_q, '')
                    })
                current_q = int(is_q.group(1))
                current_text = text
            else:
                current_text += "\n" + text

        if current_q:
            qt, choices = split_question_and_choices(current_text)
            questions.append({
                "question_number": current_q,
                "question_text": qt,
                "choices": choices,
                "question_has_image": False,
                "question_image_url": None,
                "answer": answers.get(current_q, '')
            })

        data["subjects"].append({"subject_number": number, "subject_name": name, "questions": questions})

    return data

# 요약 출력
def print_exam_summary(data, from_q=9, to_q=11):
    for subj in data['subjects']:
        print(f"\n📘 {subj['subject_number']}과목: {subj['subject_name']}")
        for q in subj['questions']:
            if from_q <= q['question_number'] <= to_q:
                print(f"  - {q['question_number']}번: {q['question_text'][:60]}... (정답: {q['answer']}, 이미지: {'O' if q['question_has_image'] else 'X'})")

# 메인 실행
def main(path):
    title, date = extract_title_info(path)
    print(f"\n📄 문서: {os.path.basename(path)}")
    doc = Document(path)
    data = parse_exam(doc)
    print_exam_summary(data, from_q=9, to_q=11)

if __name__ == "__main__":
    main("가스기사20200606.docx")
