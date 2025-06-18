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
def replace_filled_numbers(paragraph, counter):
    filled_to_unfilled = {
        "❶": "①", "❷": "②", "❸": "③", "❹": "④",
    }
    for run in paragraph.runs:
        for filled, unfilled in filled_to_unfilled.items():
            if filled in run.text:
                run.text = run.text.replace(filled, unfilled)
                run.bold = False  # 굵기 제거
                counter[0] += 1

# <<<QUESTION>>> 삽입
def insert_question_markers(paragraphs):
    for p in paragraphs:
        text = p.text.strip()
        if not text:
            continue
        bold = any(run.bold for run in p.runs if run.text.strip())
        if bold and re.match(r"^\d+\.\s", text):
            insert_paragraph_before(p, "<<<QUESTION>>>")

# 안내문 삭제
def remove_cbt_notice(paragraphs):
    start_idx, end_idx = None, None
    for i, p in enumerate(paragraphs):
        text = p.text.strip()
        if start_idx is None and text.startswith("전자문제집 CBT"):
            start_idx = i
        if start_idx is not None and text.endswith("확인하세요."):
            end_idx = i
            break

    if start_idx is not None and end_idx is not None:
        for i in range(start_idx, end_idx + 1):
            paragraphs[i]._element.getparent().remove(paragraphs[i]._element)
        print(f"🗑️ 안내문 삭제 완료: 문단 {start_idx} ~ {end_idx}")
    else:
        print("⚠️ 안내문 텍스트를 찾지 못했습니다.")


# 과목 텍스트에 (Subject) 추가
def is_paragraph_in_table(paragraph):
    parent = paragraph._element
    while parent is not None:
        if parent.tag.endswith("tbl"):
            return True
        parent = parent.getparent()
    return False

def mark_subject_titles(paragraphs):
    subject_count = 0
    for para in paragraphs:
        if is_paragraph_in_table(para):  # 표 내부에 있는 문단인지 확인
            is_bold = any(run.bold for run in para.runs if run.text.strip())
            if is_bold and para.text.strip():
                para.text = f"(Subject) {para.text.strip()} (Subject)"
                subject_count += 1
    print(f"🏷️ 과목 마킹 완료: 총 {subject_count}개")


# 굵은 문단 수 세기 + (Bold) 표시
def count_bold_paragraphs(paragraphs):
    count = 0
    for para in paragraphs:
        if any(run.bold for run in para.runs if run.text.strip()):
            para.add_run(" (Bold)")
            count += 1
    print(f"📝 굵은 글씨체 문단 수: {count}")

# 메인 실행
def main(path):
    title, date = extract_title_info(path)
    print(f"\n📄 문서: {os.path.basename(path)}")
    doc = Document(path)

    # ✅ 문단 리스트 추출
    paragraphs = []
    for b in iter_block_items(doc):
        if isinstance(b, Paragraph):
            paragraphs.append(b)

    # 1. 안내문 삭제
    remove_cbt_notice(paragraphs)

    # 2. 문단 다시 추출 (삭제 후 반영)
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

    # 3. 채워진 번호 변환
    counter = [0]
    for p in paragraphs:
        replace_filled_numbers(p, counter)
    print(f"✅ 숫자 변환: 총 {counter[0]}개")

    # 4. <<<QUESTION>>> 삽입
    insert_question_markers(paragraphs)

    # 5. 과목 표시 (Subject) 삽입
    mark_subject_titles(paragraphs)

    # 6. 굵기 문단 수 세기
    count_bold_paragraphs(paragraphs)

    # 7. 저장
    output_path = f"marked6_{os.path.basename(path)}"
    doc.save(output_path)
    print(f"✅ 저장 완료: {output_path}")

if __name__ == "__main__":
    main("가스기사20200606.docx")