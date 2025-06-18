import re
import os
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement

# ocr-env_examToDB\Scripts\activate
# python examToDB.py

# 문서 블록 순회
def iter_block_items(parent):
    parent_elm = parent.element.body
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

# 문단 앞에 새로운 문단 삽입
def insert_paragraph_before(paragraph, text):
    new_p = OxmlElement("w:p")
    paragraph._element.addprevious(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    new_para.add_run(text)
    return new_para

# 채워진 번호를 비채워진 번호로 변환
def replace_filled_numbers(paragraph):
    filled_to_unfilled = {
        "❶": "①", "❷": "②", "❸": "③", "❹": "④",
    }
    for run in paragraph.runs:
        original_text = run.text
        for filled, unfilled in filled_to_unfilled.items():
            if filled in run.text:
                run.text = run.text.replace(filled, unfilled)
                run.bold = False  # 굵기 제거
        if original_text != run.text:
            print(f"🔄 변환됨: '{original_text}' → '{run.text}'")

# <<<QUESTION>>> 삽입 + 숫자 변환
def insert_question_and_subject_markers(doc):
    paragraphs = []
    for b in iter_block_items(doc):
        if isinstance(b, Paragraph):
            paragraphs.append(b)
        elif isinstance(b, Table):
            for row in b.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        paragraphs.append(para)

    print(f"\n📄 전체 문단 수: {len(paragraphs)}")

    for p in paragraphs:
        text = p.text.strip()
        if not text:
            continue

        # 채워진 숫자 → 비채워진 숫자 (굵기 제거 포함)
        replace_filled_numbers(p)

        # 문제 번호 표시
        bold = any(run.bold for run in p.runs if run.text.strip())
        if bold and re.match(r"^\d+\.\s", text):
            insert_paragraph_before(p, "<<<QUESTION>>>")

# 메인 실행
def main(path):
    title, date = extract_title_info(path)
    print(f"\n📄 문서: {os.path.basename(path)}")
    doc = Document(path)
    insert_question_and_subject_markers(doc)
    output_path = f"marked3_{os.path.basename(path)}"
    doc.save(output_path)
    print(f"✅ 저장 완료: {output_path}")

if __name__ == "__main__":
    main("가스기사20200606.docx")