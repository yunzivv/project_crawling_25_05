# ocr-env_examToDB\Scripts\activate
# python parsedExam.py

import re
import os
import pandas as pd
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image
from io import BytesIO
import base64
import requests
import time
import pyimgur
import cloudinary
import cloudinary.uploader


# 계정 정보 리스트
CLOUDINARY_CREDENTIALS = [
    {'cloud_name': 'dc12fahac', 'api_key': '622374682885518', 'api_secret': '_lJN1N_PoOkviAcZ77RMsUnIbfQ'},
    {'cloud_name': 'dnpiyrk6n', 'api_key': '692131124971581', 'api_secret': '6PVShhV1FNoqNy5IBrjVzwtITxw'},
    {'cloud_name': 'duyepqc2e', 'api_key': '336161591162649', 'api_secret': 'o7xQn166trb0WHu7tsOTZstJeDM'},
]

# 계정 순환 인덱스
current_account_index = 0

# def get_next_cloudinary_account():
#     global current_account_index
#     cred = CLOUDINARY_CREDENTIALS[current_account_index]
#     current_account_index = (current_account_index + 1) % len(CLOUDINARY_CREDENTIALS)
#     return cred

# def upload_image_to_cloudinary(image_bytes, max_retries=3):
#     try:
#         # Pillow로 이미지 포맷 변환 (안정성 ↑)
#         image = Image.open(BytesIO(image_bytes)).convert("RGB")
#         buffer = BytesIO()
#         image.save(buffer, format="PNG")
#         image_data = buffer.getvalue()

#         for attempt in range(1, max_retries + 1):
#             try:
#                 # 계정 설정
#                 cred = get_next_cloudinary_account()
#                 cloudinary.config(
#                     cloud_name=cred['cloud_name'],
#                     api_key=cred['api_key'],
#                     api_secret=cred['api_secret']
#                 )

#                 result = cloudinary.uploader.upload(BytesIO(image_data), resource_type="image")
#                 url = result.get("secure_url")
#                 if url:
#                     print(f"✅ 이미지 업로드 성공: {url}")
#                     time.sleep(2) 
#                     return url
#                 else:
#                     print("❌ 업로드 실패")
#                     return None

#             except cloudinary.exceptions.Error as e:
#                 print(f"⚠️ Cloudinary 오류 발생 (시도 {attempt}/{max_retries}): {e}")
#                 time.sleep(5)  # 서버 오류 or 일시적 문제 대응

#     except Exception as e:
#         print("❌ 이미지 처리 오류 (PIL 등):", e)

#     return None

def iter_block_items(parent):
    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def is_paragraph_in_table(paragraph: Paragraph):
    parent = paragraph._element
    while parent is not None:
        if parent.tag.endswith("tbl"):
            return True
        parent = parent.getparent()
    return False


# def extract_image_url_from_paragraph(paragraph):
#     for run in paragraph.runs:
#         drawing = run._element.find(".//w:drawing", namespaces=run._element.nsmap)
#         if drawing is not None:
#             blip = drawing.find(".//a:blip", namespaces={"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
#             if blip is None:
#                 print("❌ blip (a:blip) not found in drawing")
#                 continue
#             rId = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
#             if not rId or rId not in paragraph.part.related_parts:
#                 print(f"❌ rId {rId} not found in related_parts")
#                 continue
#             image_part = paragraph.part.related_parts[rId]
#             image_bytes = image_part.blob
#             url = upload_image_to_cloudinary(image_bytes)
#             return url
#     return None


def extract_answer_map_from_table(table):
    answer_map = {}
    rows = table.rows
    for i in range(0, len(rows), 2):
        q_row = rows[i]
        a_row = rows[i + 1] if i + 1 < len(rows) else None
        if not a_row:
            continue
        for q_cell, a_cell in zip(q_row.cells, a_row.cells):
            q_text = q_cell.text.strip()
            a_text = a_cell.text.strip()
            if q_text.isdigit() and a_text in "①②③④":
                answer_map[int(q_text)] = "①②③④".index(a_text) + 1
    return answer_map

def parse_exam_doc(doc_path):
    doc = Document(doc_path)
    paragraphs = [p for p in iter_block_items(doc) if isinstance(p, Paragraph)]
    answer_map = extract_answer_map_from_table(doc.tables[-1])

    results = []
    current_subject = None
    current_subject_number = None
    current_question = {}
    is_question_block = False
    is_choice_block = False 

    for para in paragraphs:
        text = para.text.strip()

        if text.startswith("(Subject)") and text.endswith("(Subject)"):
            subject_content = text.replace("(Subject)", "").strip()
            match = re.match(r"(\d+\uACFC\uBAA9)\s*:\s*(.+)", subject_content)
            if match:
                current_subject_number = match.group(1)
                current_subject = match.group(2)
            continue

        if text == "<<<QUESTION>>>":
            if current_question:
                results.append(current_question)
            current_question = {
                "subject_number": current_subject_number,
                "subject": current_subject,
                "question_number": None,
                "question_text": "",
                "has_image": False,
                "image_url": None,
                "choices": [],
            }
            is_question_block = True
            is_choice_block = False
            continue

        if is_question_block:
            match = re.match(r"^(\d+)\.\s*(.*)", text)
            if match:
                current_question["question_number"] = int(match.group(1))
                current_question["question_text"] += match.group(2).strip() + " "
            elif text == "[choice]":
                is_question_block = False
                is_choice_block = True
            else:
                current_question["question_text"] += text + " "
            continue

        if is_choice_block:
            match = re.match(r"(①|②|③|④)\s*(.*)", text)
            if match:
                num = "①②③④".index(match.group(1)) + 1
                content = match.group(2).strip()
                if content:
                    current_question["choices"].append((num, content))
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

                # # 이미지 개수 세기
                # image_count = sum("graphic" in run._element.xml for run in para.runs)
                # if image_count > 0:
                #     current_question["image_count"] += image_count

                # # 이미지가 2개 이상이면 이 문제 건너뛰기
                # if current_question.get("image_count", 0) > 1:
                #     current_question = {}
                #     is_question_block = False
                #     continue

                # # 이미지가 1개인 경우 업로드
                # if current_question.get("image_count", 0) == 1 and not current_question["has_image"]:
                #     current_question["has_image"] = True
                #     img_url = extract_image_url_from_paragraph(para)
                #     current_question["image_url"] = img_url if img_url else "UPLOAD_FAILED"


                if "[choice]" in text or text.startswith(("①", "②", "③", "④")):
                    if current_question and "choices" in current_question:
                        choice_text = para.text.strip()
                        match = re.match(r"(\[choice\])?\s*(①|②|③|④)\s*(.*)", choice_text)
                        if match:
                            num = "①②③④".index(match.group(2)) + 1
                            content = match.group(3).strip()
                            if content:
                                current_question["choices"].append((num, content))
                    else:
                        print(f"⚠️ 선택지를 만났지만 current_question이 비정상 상태입니다: \"{text}\"")
                
    # ✅✅✅ 이 아래 코드 반드시 추가!
    if current_question and current_question.get("question_number"):
        if len(current_question["choices"]) == 4:
            results.append(current_question)
        else:
            print(f"⚠️ 선택지 누락 - 문제 {current_question['question_number']} 건너뜀 (선택지 {len(current_question['choices'])}개)")


    # if current_question and current_question.get("question_number"):
    #     if len(current_question["choices"]) == 4:
    #         results.append(current_question)
    #     else:
    #         print(f"⚠️ 선택지 누락 - 문제 {current_question['question_number']} 건너뜀 (선택지 {len(current_question['choices'])}개)")


    for q in results:
        qnum = q["question_number"]
        q["answer_number"] = answer_map.get(qnum)
        new_choices = []
        for num, text in q["choices"]:
            is_correct = (num == q["answer_number"])
            new_choices.append((num, text, is_correct))
        q["choices"] = new_choices

    return results

def process_all_exam_files(input_folder, start_index=0, end_index=250):
    # all_questions = []
    all_choices = []

    # 기존 파일 로딩
    if os.path.exists("questions.xlsx") and os.path.exists("choices.xlsx"):

        # df_questions_existing = pd.read_excel("questions.xlsx")
        df_choices_existing = pd.read_excel("choices.xlsx")
        print("📂 기존 엑셀 파일 로드 완료")

        # last_exam_id = df_questions_existing["시험ID"].max()
        # last_question_id = df_questions_existing["문제ID"].max()
        last_exam_id = 0
        last_question_id = 0

    else:
        # df_questions_existing = pd.DataFrame()
        df_choices_existing = pd.DataFrame()
        last_exam_id = 0
        last_question_id = 0

    exam_id = last_exam_id + 1
    question_id_counter = last_question_id + 1

    filenames = sorted([f for f in os.listdir(input_folder) if f.endswith('.docx')])[start_index:end_index]

    for filename in filenames:
        filepath = os.path.join(input_folder, filename)
        print(filename)
        parsed_questions = parse_exam_doc(filepath)

        basename = os.path.splitext(filename)[0]
        match = re.match(r"([^\d]+)(\d{8})", basename)
        if match:
            cert_name = match.group(1)
            exam_date = match.group(2)
        else:
            cert_name = ""
            exam_date = ""

        print(f"▶️ 파일: {filename}, 자격증명: {cert_name}, 시험일자: {exam_date}, 문제 수: {len(parsed_questions)}")

        for q in parsed_questions:
            current_qid = question_id_counter

            # all_questions.append({
            #     "자격증명": cert_name,
            #     "시험일자": exam_date,
            #     "시험ID": exam_id,
            #     "문제ID": current_qid,
            #     "과목번호": q["subject_number"],
            #     "과목명": q["subject"],
            #     "문제번호": q["question_number"],
            #     "문제텍스트": q["question_text"].strip(),
            #     "이미지포함": "true" if q["has_image"] else "false",
            #     "이미지URL": q["image_url"] or ""
            # })

            for num, text, is_correct in q["choices"]:
                all_choices.append({
                    "자격증명": cert_name,
                    "시험일자": exam_date,
                    "시험ID": exam_id,
                    "문제ID": current_qid,
                    "선택지번호": num,
                    "선택지내용": text,
                    "정답여부": "true" if is_correct else "false"
                })

            question_id_counter += 1

        exam_id += 1

    # 새로운 데이터프레임 생성
    # df_new_questions = pd.DataFrame(all_questions)
    df_new_choices = pd.DataFrame(all_choices)

    # 기존 데이터와 병합
    # df_questions_final = pd.concat([df_questions_existing, df_new_questions], ignore_index=True)
    df_choices_final = pd.concat([df_choices_existing, df_new_choices], ignore_index=True)

    # df_questions_final.to_excel("questions.xlsx", index=False)
    df_choices_final.to_excel("choices.xlsx", index=False)
    print("✅ 추가 데이터 저장 완료: questions.xlsx, choices.xlsx")


if __name__ == "__main__":
    process_all_exam_files("기출문제포맷", start_index=0, end_index=250) 
