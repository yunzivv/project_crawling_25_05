import fitz  # PyMuPDF: pdf 읽기
import re # 정규식
import pytesseract # OCR
from PIL import Image # 이미지 분석
import cv2 # 이미지 자르기
import numpy as np # 이미지 데이터 배열 저장
import docx # 워드 문서 다루기
from collections import defaultdict # 딕셔너리 같은 데이터 구조
import os # 파일, 디렉토리 관리
import requests
from io import BytesIO # 이미지 플래그 및 업로드
from PIL import Image
from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
# ocr-env_examToDB\Scripts\activate
# python examToDB.py


# 문서 내 요소 순회
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

# 파일명에서 자격증명과 날짜 추출
def extract_title_info(filename):
    basename = os.path.basename(filename)
    name, _ = os.path.splitext(basename)
    match = re.match(r'(.+?)(\d{8})$', name)
    if match:
        return match.group(1).strip(), match.group(2)
    return name, None

# 정답 테이블 추출 (지그재그 구조)
def extract_answers_from_zigzag_table(table):
    answers = {}
    rows = table.rows
    for i in range(0, len(rows), 2):
        if i + 1 >= len(rows):
            break
        q_nums = [cell.text.strip() for cell in rows[i].cells]
        q_ans = [cell.text.strip() for cell in rows[i+1].cells]
        for q, a in zip(q_nums, q_ans):
            if q.isdigit():
                answers[int(q)] = a
    return answers

# 선택지 추출
def extract_choices_from_lines(lines):
    joined = " ".join(lines)
    
    # 선택지 번호로 split
    split_choices = re.split(r"(①|②|③|④|❶|❷|❸|❹)", joined)

    choice_chunks = []
    for i in range(1, len(split_choices) - 1, 2):
        number = split_choices[i]
        text = split_choices[i + 1].strip()
        choice_chunks.append((number, text))

    # 딱 4개로 보정
    while len(choice_chunks) < 4:
        choice_chunks.append(("?", ""))

    choices = []
    for idx, (_, text) in enumerate(choice_chunks):
        choices.append({
            "number": idx + 1,
            "text": text,
            "has_image": False,
            "image_url": None
        })
    return choices

# 문제 텍스트와 보기 분리
def split_question_and_choices(lines):
    choice_pattern = re.compile(r"[①②③④❶❷❸❹]")
    question_part = []
    for line in lines:
        if choice_pattern.search(line):
            cut = choice_pattern.search(line).start()
            question_part.append(line[:cut].strip())
            break
        question_part.append(line.strip())
    qt = " ".join(question_part).strip()
    choices = extract_choices_from_lines(lines)
    return qt, choices

# 시험지 파서 개선
def parse_exam(texts, answer_table=None):
    data = {"subjects": [], "answers": {}}
    if answer_table:
        data["answers"] = extract_answers_from_zigzag_table(answer_table)

    current_subject = None
    question_buffer = []
    question_number = 0

    subject_pattern = re.compile(r"^(\d+)과목\s*[:：]\s*(.+)$")
    question_pattern = re.compile(r"^(\d+)[.\)]")

    for i, text in enumerate(texts):
        subj_match = subject_pattern.match(text)
        if subj_match:
            if current_subject:
                if question_buffer:
                    qt, choices = split_question_and_choices(question_buffer)
                    current_subject["questions"].append({
                        "question_number": question_number,
                        "question_text": qt,
                        "choices": choices,
                        "question_has_image": False,
                        "question_image_url": None,
                        "answer": data['answers'].get(question_number, '')
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
                qt, choices = split_question_and_choices(question_buffer)
                current_subject["questions"].append({
                    "question_number": question_number,
                    "question_text": qt,
                    "choices": choices,
                    "question_has_image": False,
                    "question_image_url": None,
                    "answer": data['answers'].get(question_number, '')
                })
            question_number = int(q_match.group(1))
            question_buffer = [text]
            continue

        if question_buffer:
            question_buffer.append(text)

    if current_subject and question_buffer:
        qt, choices = split_question_and_choices(question_buffer)
        current_subject["questions"].append({
            "question_number": question_number,
            "question_text": qt,
            "choices": choices,
            "question_has_image": False,
            "question_image_url": None,
            "answer": data['answers'].get(question_number, '')
        })
        data["subjects"].append(current_subject)

    return data

# 이미지 여부
def has_image(paragraph):
    for run in paragraph.runs:
        if run._element.xpath(".//w:drawing"):
            return True
    return False

# imgur 업로드
def upload_image_to_imgur(image_bytes):
    CLIENT_ID = '00ff8e726eb9eb8'
    url = "https://api.imgur.com/3/image"
    headers = {'Authorization': f'Client-ID {CLIENT_ID}'}
    response = requests.post(url, headers=headers, files={"image": image_bytes})
    if response.status_code == 200:
        return response.json()['data']['link']
    return None


def assign_image_flags(doc, exam_data):
    paragraphs = list(doc.paragraphs)
    image_indices = {}

    # 이미지 포함 문단 수집
    for i, para in enumerate(paragraphs):
        for run in para.runs:
            drawing = run._element.xpath(".//*[local-name()='drawing']")
            if drawing:
                blip = drawing[0].xpath(".//*[local-name()='blip']")
                if blip:
                    rId = blip[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    image_part = doc.part.related_parts[rId]
                    image_bytes = image_part.blob
                    image_indices[i] = image_bytes

    used_image_indices = set()
    current_index = 0

    for subj in exam_data["subjects"]:
        questions = subj["questions"]
        for qidx, q in enumerate(questions):
            # 문제 시작 위치 탐색
            found = False
            for i in range(current_index, len(paragraphs)):
                if q["question_text"][:10] in paragraphs[i].text:
                    start = i
                    found = True
                    break
            if not found:
                continue
            current_index = start

            # 다음 문제 시작 전까지 범위 제한
            next_q_text = questions[qidx + 1]["question_text"][:10] if qidx + 1 < len(questions) else None
            end_index = len(paragraphs)
            if next_q_text:
                for j in range(start + 1, len(paragraphs)):
                    if next_q_text in paragraphs[j].text:
                        end_index = j
                        break

            # 문제 이미지 탐색
            q["question_has_image"] = False
            q["question_image_url"] = None
            for idx_img in range(start, end_index):
                if idx_img in image_indices and idx_img not in used_image_indices:
                    q["question_has_image"] = True
                    img_url = upload_image_to_imgur(image_indices[idx_img])
                    if img_url:
                        q["question_image_url"] = img_url
                        used_image_indices.add(idx_img)
                    break

            # 선택지 이미지 초기화
            for ch in q["choices"]:
                ch["has_image"] = False
                ch["image_url"] = None

            # 텍스트 없는 이미지 문단 우선 탐색
            image_only = [
                i for i in range(start, end_index)
                if i in image_indices and not paragraphs[i].text.strip() and i not in used_image_indices
            ]

            image_ptr = 0
            for ch in q["choices"]:
                if image_ptr < len(image_only):
                    ch["has_image"] = True
                    ch["image_url"] = upload_image_to_imgur(image_indices[image_only[image_ptr]])
                    used_image_indices.add(image_only[image_ptr])
                    image_ptr += 1
                elif not ch["text"].strip():
                    # 보기 텍스트가 없는 경우: 아무 이미지라도 사용
                    for i in range(start, end_index):
                        if i in image_indices and i not in used_image_indices:
                            ch["has_image"] = True
                            ch["image_url"] = upload_image_to_imgur(image_indices[i])
                            used_image_indices.add(i)
                            break



# 메인 실행
def main(docx_path):
    title, date = extract_title_info(docx_path)
    print(f"제목: {title}, 날짜: {date if date else '날짜 없음'}")
    doc = Document(docx_path)
    blocks = list(iter_block_items(doc))
    texts = []
    tables = []
    for b in blocks:
        if isinstance(b, Paragraph):
            t = b.text.strip()
            if t:
                texts.append(t)
        elif isinstance(b, Table):
            tables.append(b)
            for row in b.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        t = para.text.strip()
                        if t:
                            texts.append(t)

    # 마지막 테이블을 정답표로 지정
    answer_table = tables[-1] if tables else None

    exam_data = parse_exam(texts, answer_table)
    assign_image_flags(doc, exam_data)

    for subj in exam_data['subjects']:
        print(f"\n📘 {subj['subject_number']}과목: {subj['subject_name']}")
        print(f"총 {len(subj['questions'])}문제")
        for q in subj['questions'][8:11]:
            print(f"  - {q['question_number']}번 문제: {q['question_text'][:60]}... (정답: {q['answer']}, 이미지: {'O' if q['question_has_image'] else 'X'})")
            if q['question_image_url']:
                print(f"    문제 이미지 URL: {q['question_image_url']}")
            for ch in q['choices']:
                if ch['has_image']:
                    print(f"    {ch['number']} 이미지: {ch['image_url']}")
                else:
                    print(f"    {ch['number']} {ch['text'][:40]}")
    return exam_data

if __name__ == "__main__":
    main("가스기사20200606.docx")