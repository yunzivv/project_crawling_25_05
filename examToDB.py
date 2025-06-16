import fitz  # PyMuPDF: pdf 읽기
import re # 정규식
import pytesseract # OCR
from PIL import Image
import io
import cv2 # 이미지 처리
import numpy as np
import docx
from collections import defaultdict

# ocr-env_examToDB\Scripts\activate

import docx
import re
import os
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph


# ✅ 문서 안의 문단과 표를 "원래 순서대로" 순회
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

# ✅ 파일명에서 제목, 날짜 추출
def extract_title_info(filename):
    basename = os.path.basename(filename)
    name, _ = os.path.splitext(basename)
    match = re.match(r'(.+?)(\d{8})$', name)
    if match:
        return match.group(1).strip(), match.group(2)
    return name, None

# ✅ 텍스트 순서대로 리스트로 추출 (문단 + 표)
def extract_all_text_ordered(doc):
    texts = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text:
                texts.append(text)
        elif isinstance(block, Table):
            for row in block.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        text = para.text.strip()
                        if text:
                            texts.append(text)
    return texts

# ✅ 시험지 파서
def parse_exam(texts):
    data = {
        "subjects": []
    }

    current_subject = None
    question_buffer = []
    question_number = 0

    subject_pattern = re.compile(r"^\s*(\d+)과목\s*[:：]\s*(.+)$")
    question_pattern = re.compile(r"^(\d+)[.\\)]")
    choice_pattern = re.compile(r"[①②③④❶❷❸❹]")

    for i, text in enumerate(texts):
        subj_match = subject_pattern.match(text)
        if subj_match:
            if current_subject:
                if question_buffer:
                    current_subject["questions"].append({
                        "question_number": question_number,
                        "question_text": " ".join(question_buffer)
                    })
                    question_buffer = []
                data["subjects"].append(current_subject)

            current_subject = {
                "subject_number": int(subj_match.group(1)),
                "subject_name": subj_match.group(2).strip(),
                "questions": []
            }
            continue

        q_match = question_pattern.match(text)
        if q_match:
            if current_subject is None:
                print(f"⚠️ 과목 없이 문제 발견 (문단 {i}): {text}")
                continue

            if question_buffer:
                current_subject["questions"].append({
                    "question_number": question_number,
                    "question_text": " ".join(question_buffer)
                })

            question_number = int(q_match.group(1))
            question_buffer = [text]
            continue

        if question_buffer:
            question_buffer.append(text)

    if current_subject and question_buffer:
        current_subject["questions"].append({
            "question_number": question_number,
            "question_text": " ".join(question_buffer)
        })
        data["subjects"].append(current_subject)

    return data

# ✅ 메인 실행
def main(docx_path):
    title, date = extract_title_info(docx_path)
    print(f"제목: {title}, 날짜: {date if date else '날짜 없음'}")

    doc = docx.Document(docx_path)
    texts = extract_all_text_ordered(doc)

    exam_data = parse_exam(texts)

    for subj in exam_data['subjects']:
        print(f"\n📘 {subj['subject_number']}과목: {subj['subject_name']}")
        print(f"총 {len(subj['questions'])}문제")
        for q in subj['questions'][:2]:
            print(f"  - {q['question_number']}번: {q['question_text'][:60]}...")

    return exam_data



# ✅ 파일 실행 (변경 가능)
if __name__ == "__main__":
    docx_file = "가스기사20200606.docx"
    main(docx_file)
