import win32com.client
import os

hwp = win32com.client.Dispatch("HWPFrame.HwpObject")

input_dir = r"C:\HWP_INPUT"      # HWP 파일들이 들어있는 폴더
output_dir = r"C:\DOCX_OUTPUT"   # DOCX로 저장할 폴더

# 한글 보안모듈 무시 설정
hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")

file_list = [f for f in os.listdir(input_dir) if f.endswith(".hwp")]

for filename in file_list:
    hwp.Open(os.path.join(input_dir, filename))
    
    docx_name = os.path.splitext(filename)[0] + ".docx"
    save_path = os.path.join(output_dir, docx_name)

    hwp.SaveAs(save_path, "DOCX")  # 저장 형식: DOCX
    hwp.Quit()