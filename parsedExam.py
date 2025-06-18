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
import pyimgur
from PIL import Image
from io import BytesIO
import base64
import requests
import time

IMGUR_CLIENT_IDS = ["6b9042903a1fea3", "a1608f84c2725a5"]
current_imgur_index = 0  # 전역 변수

def get_next_client_id():
    global current_imgur_index
    client_id = IMGUR_CLIENT_IDS[current_imgur_index]
    current_imgur_index = (current_imgur_index + 1) % len(IMGUR_CLIENT_IDS)
    return client_id

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

# def upload_image_to_imgur(image_bytes):
#     try:
#         image = Image.open(BytesIO(image_bytes)).convert("RGB")
#         buffer = BytesIO()
#         image.save(buffer, format="PNG")
#         encoded = base64.b64encode(buffer.getvalue()).decode("utf-8")

#         client_id = get_next_client_id()
#         headers = {"Authorization": f"Client-ID {client_id}"}
#         data = {
#             'image': encoded,
#             'type': 'base64',
#             'name': 'upload.png',
#         }
#         response = requests.post("https://api.imgur.com/3/image", headers=headers, data=data)
#         if response.status_code == 200:
#             # print("✅ 이미지 업로드 성공", response.json()['data']['link'])
#             return response.json()['data']['link']
#         else:
#             print("❌ 업로드 실패:", response.status_code, response.text)
#             return None
#     except Exception as e:
#         print("❌ 이미지 처리 실패:", e)
#         return None

def extract_image_url_from_paragraph(paragraph):
    for run in paragraph.runs:
        drawing = run._element.find(".//w:drawing", namespaces=run._element.nsmap)
        if drawing is not None:
            blip = drawing.find(".//a:blip", namespaces={"a": "http://schemas.openxmlformats.org/drawingml/2006/main"})
            if blip is None:
                print("❌ blip (a:blip) not found in drawing")
                continue
            rId = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if not rId:
                print("❌ No embed ID (r:embed) found in blip")
                continue
            if rId not in paragraph.part.related_parts:
                print(f"❌ rId {rId} not found in related_parts")
                continue
            image_part = paragraph.part.related_parts[rId]
            image_bytes = image_part.blob 
            url = upload_image_to_imgur(image_bytes)
            print(f"✅ 이미지 업로드 성공: {url}")
            time.sleep(5)
            return url
    return None

# 업로드 실패 시 대기 
def upload_image_to_imgur(image_bytes, max_retries=3):
    try:
        image = Image.open(BytesIO(image_bytes)).convert("RGB")
        buffer = BytesIO()
        image.save(buffer, format="PNG")
        encoded = base64.b64encode(buffer.getvalue()).decode("utf-8")

        client_id = get_next_client_id()
        headers = {"Authorization": f"Client-ID {client_id}"}
        data = {'image': encoded, 'type': 'base64', 'name': 'upload.png'}

        for attempt in range(1, max_retries + 1):
            response = requests.post("https://api.imgur.com/3/image", headers=headers, data=data)
            if response.status_code == 200:
                url = response.json()['data']['link']
                print(f"✅ 이미지 업로드 성공: {url}")
                return url
            elif response.status_code == 429:
                print("🚫 업로드 제한 도달. 60초 대기 후 재시도...")
                time.sleep(60)
            elif response.status_code >= 500:
                print(f"⚠️ 서버 오류 ({response.status_code}). {attempt}/{max_retries}회 재시도 중...")
                time.sleep(5)
            else:
                print(f"❌ 업로드 실패: {response.status_code}\n{response.text}")
                break

    except Exception as e:
        print("❌ 이미지 처리 오류:", e)

    return None


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
                    img_url = extract_image_url_from_paragraph(para)
                    current_question["image_url"] = img_url if img_url else "UPLOAD_FAILED"
                    # if any("graphic" in run._element.xml for run in para.runs): # 이미지 업로드 잠시 중단
                    #     current_question["has_image"] = True


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

    for q in results:
        qnum = q["question_number"]
        q["answer_number"] = answer_map.get(qnum)
        new_choices = []
        for num, text in q["choices"]:
            is_correct = (num == q["answer_number"])
            new_choices.append((num, text, is_correct))
        q["choices"] = new_choices

    return results

def process_all_exam_files(input_folder):
    all_questions = []
    all_choices = []

    exam_id = 1
    question_id_counter = 1

    filenames = sorted([f for f in os.listdir(input_folder) if f.endswith('.docx')])

    for filename in filenames:
        filepath = os.path.join(input_folder, filename)
        print(filename)
        parsed_questions = parse_exam_doc(filepath)

        # 시험명과 날짜 추출 (예: '가스기사20200606' → '가스기사', '20200606')
        basename = os.path.splitext(filename)[0]
        match = re.match(r"([^\d]+)(\d{8})", basename)
        if match:
            cert_name = match.group(1)
            exam_date = match.group(2)
        else:
            cert_name = ""
            exam_date = ""

        print(f"▶️ 현재 파일: {filename}, 자격증명: {cert_name}, 시험일자: {exam_date}, 문제 수: {len(parsed_questions)}")

        for q in parsed_questions:
            current_qid = question_id_counter

            all_questions.append({
                "자격증명": cert_name,
                "시험일자": exam_date,
                "시험ID": exam_id,
                "문제ID": current_qid,
                "과목번호": q["subject_number"],
                "과목명": q["subject"],
                "문제번호": q["question_number"],
                "문제텍스트": q["question_text"].strip(),
                "이미지포함": "true" if q["has_image"] else "false",
                "이미지URL": q["image_url"] or ""
            })

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

    df_questions = pd.DataFrame(all_questions)
    df_choices = pd.DataFrame(all_choices)

    df_questions.to_excel("questions.xlsx", index=False)
    df_choices.to_excel("choices.xlsx", index=False)
    print("✅ 전체 시험 Excel 저장 완료: questions.xlsx, choices.xlsx")


if __name__ == "__main__":
    process_all_exam_files("기출문제포맷")    
