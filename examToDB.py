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

# 이미지 처리
def assign_images(doc, exam_data):
    paragraphs = []
    for b in iter_block_items(doc):
        if isinstance(b, Paragraph):
            paragraphs.append(b)
        elif isinstance(b, Table):
            for row in b.rows:
                for cell in row.cells:
                    paragraphs.extend(cell.paragraphs)

    image_indices = {}
    for i, p in enumerate(paragraphs):
        for run in p.runs:
            drawing = run._element.xpath(".//*[local-name()='drawing']")
            if drawing:
                blip = drawing[0].xpath(".//*[local-name()='blip']")
                if blip:
                    rId = blip[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                    image_part = doc.part.related_parts[rId]
                    image_bytes = image_part.blob
                    image_indices[i] = image_bytes

    used = set()
    for subj in exam_data["subjects"]:
        for q in subj["questions"]:
            start_idx = next((i for i, p in enumerate(paragraphs) if q["question_text"][:10].strip() in p.text), -1)
            if start_idx == -1:
                continue
            # 문제 이미지
            for offset in range(5):
                idx = start_idx + offset
                if idx in image_indices and idx not in used:
                    q["question_has_image"] = True
                    q["question_image_url"] = upload_image_to_imgur(image_indices[idx])
                    used.add(idx)
                    break

            # 선택지 이미지
            img_candidates = [i for i in range(start_idx + 1, start_idx + 10) if i in image_indices and not paragraphs[i].text.strip() and i not in used]
            for i, ch in enumerate(q["choices"]):
                if ch["text"]:
                    continue
                if i < len(img_candidates):
                    ch["has_image"] = True
                    ch["image_url"] = upload_image_to_imgur(image_indices[img_candidates[i]])
                    used.add(img_candidates[i])

# 문서 파싱

def parse_exam(doc):
    paragraphs = []
    for b in iter_block_items(doc):
        if isinstance(b, Paragraph):
            paragraphs.append(b)
        elif isinstance(b, Table):
            for row in b.rows:
                for cell in row.cells:
                    paragraphs.extend(cell.paragraphs)

    blocks = list(iter_block_items(doc))
    tables = [b for b in blocks if isinstance(b, Table)]
    answer_table = tables[-1] if tables else None
    answers = extract_answers_from_table(answer_table) if answer_table else {}

    subjects = []
    for b in blocks:
        if isinstance(b, Table) and len(b.rows) == 1 and len(b.rows[0].cells) == 1:
            text = b.rows[0].cells[0].text.strip()
            print(f"🔍 과목 후보 텍스트: '{text}'")
            m = re.match(r"(\d)과목\s*[:：]\s*(.+)", text)
            if m:
                print(f"✅ 과목 인식 성공: {text}")
                subjects.append((int(m.group(1)), m.group(2), b))

    if not subjects:
        print("❌ 과목을 찾지 못했습니다. 파싱이 실패했을 수 있습니다.")
        return {"subjects": []}

    subject_indices = []
    for (_, _, table) in subjects:
        found = False
        for para in table.rows[0].cells[0].paragraphs:
            if para in paragraphs:
                subject_indices.append(paragraphs.index(para))
                found = True
                break
        if not found:
            subject_indices.append(-1)

    # 음수 제거 + 안전한 종료 지점 추가
    subject_indices = [i for i in subject_indices if i >= 0]
    while len(subject_indices) < len(subjects):
        subject_indices.append(len(paragraphs))  # ⚠️ 부족한 경우 안전하게 문서 끝으로 대체
    subject_indices.append(len(paragraphs))  # ✅ 마지막 과목의 끝 경계 추가

    data = {"subjects": []}
    for i in range(len(subjects)):
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
            q_match = re.match(r"^(\d+)[.)]", text)
            if bold and q_match:
                if current_q:
                    q_text, choices = split_question_and_choices(current_text)
                    questions.append({
                        "question_number": current_q,
                        "question_text": q_text,
                        "choices": choices,
                        "question_has_image": False,
                        "question_image_url": None,
                        "answer": answers.get(current_q, '')
                    })
                current_q = int(q_match.group(1))
                current_text = text
            else:
                current_text += "\n" + text
        if current_q:
            q_text, choices = split_question_and_choices(current_text)
            questions.append({
                "question_number": current_q,
                "question_text": q_text,
                "choices": choices,
                "question_has_image": False,
                "question_image_url": None,
                "answer": answers.get(current_q, '')
            })

        data["subjects"].append({"subject_number": number, "subject_name": name, "questions": questions})

    return data

# 출력

def print_exam_summary(data, q_start=9, q_end=11):
    for subj in data['subjects']:
        print(f"\n📘 {subj['subject_number']}과목: {subj['subject_name']}")
        for q in subj['questions']:
            if q_start <= q['question_number'] <= q_end:
                print(f"  - {q['question_number']}번 문제: {q['question_text'][:60]}... (정답: {q['answer']}, 이미지: {'O' if q['question_has_image'] else 'X'})")
                if q['question_image_url']:
                    print(f"    문제 이미지 URL: {q['question_image_url']}")
                for ch in q['choices']:
                    if ch['has_image']:
                        print(f"    {ch['number']}번 보기 이미지 URL: {ch['image_url']}")
                    else:
                        print(f"    {ch['number']}번 보기 텍스트: {ch['text'][:40]}")

# 실행

def main(path):
    title, date = extract_title_info(path)
    print(f"\n📄 문서: {os.path.basename(path)}")
    doc = Document(path)
    data = parse_exam(doc)
    assign_images(doc, data)
    print_exam_summary(data, 9, 11)

if __name__ == "__main__":
    main("가스기사20200606.docx")