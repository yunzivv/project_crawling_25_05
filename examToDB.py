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

# ✅ 파일명에서 자격증명과 날짜 추출
def extract_title_info(filename):
    basename = os.path.basename(filename)
    name, _ = os.path.splitext(basename)
    match = re.match(r'(.+?)(\d{8})$', name)
    if match:
        return match.group(1).strip(), match.group(2)
    return name, None

# ✅ docx 문서에서 본문 추출
def read_docx(filepath):
    doc = docx.Document(filepath)
    return [p.text.strip() for p in doc.paragraphs if p.text.strip()]

# ✅ 시험지 분석: 과목, 문제, 보기 추출
def parse_exam(paragraphs):
    data = {
        "subjects": []
    }

    current_subject = None
    question_buffer = []
    question_number = 0

    subject_pattern = re.compile(r"^\s*(\d+)과목\s*[:：]\s*(.+)$")
    question_pattern = re.compile(r"^(\d+)[.\\)]")
    choice_pattern = re.compile(r"[①②③④❶❷❸❹]")

    for i, para in enumerate(paragraphs):
        # 과목 줄 감지
        subj_match = subject_pattern.match(para)
        if subj_match:
            # 이전 과목 저장
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

        # 문제 번호 감지
        q_match = question_pattern.match(para)
        if q_match:
            if current_subject is None:
                print(f"⚠️ 과목 없이 문제 발견 (문단 {i}): {para}")
                continue

            # 이전 문제 저장
            if question_buffer:
                current_subject["questions"].append({
                    "question_number": question_number,
                    "question_text": " ".join(question_buffer)
                })

            question_number = int(q_match.group(1))
            question_buffer = [para]
            continue

        # 선택지 또는 일반 문단 → 현재 문제에 이어 붙임
        if question_buffer:
            question_buffer.append(para)

    # 마지막 문제 저장
    if current_subject and question_buffer:
        current_subject["questions"].append({
            "question_number": question_number,
            "question_text": " ".join(question_buffer)
        })
        data["subjects"].append(current_subject)

    return data

# ✅ 메인 처리 흐름
def main(docx_path):
    title, date = extract_title_info(docx_path)
    print(f"제목: {title}, 날짜: {date if date else '날짜 없음'}")

    paragraphs = read_docx(docx_path)
    exam_data = parse_exam(paragraphs)

    # 요약 출력
    for subj in exam_data['subjects']:
        print(f"\n📘 {subj['subject_number']}과목: {subj['subject_name']}")
        print(f"총 {len(subj['questions'])}문제")
        for q in subj['questions'][:2]:  # 미리보기 2문제
            print(f"  - {q['question_number']}번: {q['question_text'][:60]}...")

    return exam_data

# ✅ 파일 실행 (변경 가능)
if __name__ == "__main__":
    docx_file = "가스기사20200606.docx"
    main(docx_file)
