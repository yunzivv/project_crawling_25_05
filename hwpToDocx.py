import win32com.client
import os
import re

# ocr-env_docx\Scripts\activate
# python hwpToDocx.py
input_dir = r"D:\yunzi\academy\기출문제hwp"
output_dir = r"D:\yunzi\academy\hwpToDocx"

def clean_filename(filename):
    return re.sub(r'[\\/:*?"<>|]', '_', filename)


# 폴더가 없다면 생성
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

# 한글 및 Word COM 객체 준비
hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False  # Word 창 안 띄움

file_list = [f for f in os.listdir(input_dir) if f.lower().endswith(".hwp")]

for filename in file_list:
    print(f"처리 중: {filename}")

    # ✅ HWP를 매 파일마다 새로 열기
    try:
        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

        file_path = os.path.join(input_dir, filename)
        hwp.Open(file_path, "HWP", "forceopen:true")

        # 클립보드 복사
        hwp.HAction.Run("SelectAll")
        hwp.HAction.Run("Copy")
        hwp.Quit()

        # Word 문서 붙여넣기 및 저장
        doc = word.Documents.Add()
        doc.Content.Paste()

        safe_name = clean_filename(os.path.splitext(filename)[0])
        docx_path = os.path.join(output_dir, safe_name + ".docx")

        doc.SaveAs(docx_path, FileFormat=16)
        doc.Close()

    except Exception as e:
        print(f"❌ 오류 발생 - {filename}: {e}")

word.Quit()
print("✅ 모든 파일 처리 완료.")