import re
import os
import pandas as pd
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph

# 블록 순회
def iter_block_items(parent):
    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

# 문단이 표 안에 있는지 여부
def is_paragraph_in_table(paragraph: Paragraph):
    parent = paragraph._element
    while parent is not None:
        if parent.tag.endswith("tbl"):
            return True
        parent = parent.getparent()
    return False

def extract_answer_map_from_table(table):
    CHOICE_MAP = {"①": 1, "②": 2, "③": 3, "④": 4}
    answer_map = {}
    rows = table.rows
    for i in range(0, len(rows), 2):  # 두 행씩 묶기
        if i + 1 >= len(rows):
            break  # 짝이 안 맞는 경우 방지

        number_cells = rows[i].cells
        answer_cells = rows[i + 1].cells

        for nc, ac in zip(number_cells, answer_cells):
            qnum_text = nc.text.strip()
            print(qnum_text)
            ans_text = ac.text.strip()
            print(ans_text)

            if not qnum_text:
                continue

            match = re.search(r"\d+", qnum_text)
            if not match:
                continue
            qnum = int(match.group())

            if ans_text in CHOICE_MAP:
                answer_map[qnum] = CHOICE_MAP[ans_text]
            else:
                try:
                    answer_map[qnum] = int(ans_text)
                except ValueError:
                    continue  # 정답값이 이상한 경우 skip

    return answer_map


# 정답표 추출 함수 (마지막 표)
def extract_answer_map(doc):
    tables = [tbl for tbl in iter_block_items(doc) if isinstance(tbl, Table)]
    if not tables:
        return {}

    answer_table = tables[-1]
    answers = {}
    for row in answer_table.rows:
        if len(row.cells) < 2:
            continue
        qnum_cell = row.cells[0].text.strip()
        ans_cell = row.cells[1].text.strip()

        try:
            qnum = int(qnum_cell)
            match = re.search(r"[①-④]", ans_cell)
            if match:
                answers[qnum] = "①②③④".index(match.group(0)) + 1
        except:
            continue
    return answers

# 문제, 선택지, 과목 파싱
def parse_exam_doc(doc_path):
    doc = Document(doc_path)
    paragraphs = [p for p in iter_block_items(doc) if isinstance(p, Paragraph)]
    answer_map = extract_answer_map_from_table(doc.tables[-1])  # 마지막 표가 정답표

    results = []
    current_subject = None
    current_subject_number = None
    current_question = {}
    is_question_block = False

    for para in paragraphs:
        text = para.text.strip()

        # 과목
        if text.startswith("(Subject)") and text.endswith("(Subject)"):
            subject_content = text.replace("(Subject)", "").strip()
            match = re.match(r"(\d+과목)\s*:\s*(.+)", subject_content)
            if match:
                current_subject_number = match.group(1)
                current_subject = match.group(2)
            continue

        # 문제 시작
        if text == "<<<QUESTION>>>":
            if current_question:
                results.append(current_question)
            current_question = {
                "subject_number": current_subject_number,
                "subject": current_subject,
                "question_number": None,
                "question_text": "",
                "has_image": False,
                "choices": []
            }
            is_question_block = True
            continue

        # 문제 본문
        if is_question_block:
            if para.text.strip().startswith(tuple("①②③④")) or "[choice]" in para.text:
                is_question_block = False
            else:
                if current_question["question_number"] is None:
                    match = re.match(r"^(\d+)\.\s*(.*)", text)
                    if match:
                        current_question["question_number"] = int(match.group(1))
                        current_question["question_text"] += match.group(2).strip() + " "
                else:
                    current_question["question_text"] += text + " "

                if any("graphic" in run._element.xml for run in para.runs):
                    current_question["has_image"] = True
            continue

        # 선택지
        if "[choice]" in text or text.startswith(("①", "②", "③", "④")):
            choice_text = para.text.strip()
            match = re.match(r"(\[choice\])?\s*(①|②|③|④)\s*(.*)", choice_text)
            if match:
                num = "①②③④".index(match.group(2)) + 1
                content = match.group(3).strip()
                if content:
                    current_question["choices"].append((num, content))

    if current_question:
        results.append(current_question)

    # 정답 매핑 추가
    for q in results:
        qnum = q["question_number"]
        q["answer_number"] = answer_map.get(qnum)
        new_choices = []
        for num, text in q["choices"]:
            is_correct = (num == q["answer_number"])
            new_choices.append((num, text, is_correct))
        q["choices"] = new_choices

    return results


# 실행
docx_path = "marked00_가스기사20200606.docx"
parsed_data = parse_exam_doc(docx_path)

df = pd.DataFrame([
    {
        "과목번호": q["subject_number"],
        "과목명": q["subject"],
        "문제번호": q["question_number"],
        "문제텍스트": q["question_text"].strip(),
        "이미지포함": "true" if q["has_image"] else "false",
        "선택지번호": num,
        "선택지내용": text,
        "정답번호": q["answer_number"],
        "정답여부": "true" if is_correct else "false"
    }
    for q in parsed_data for num, text, is_correct in q["choices"]
])

df.to_excel("parsed_exam.xlsx", index=False)
print("✅ Excel 파일 저장 완료")
