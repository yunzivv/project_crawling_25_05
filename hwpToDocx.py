import pyautogui
import time
import win32com.client
import os
import subprocess

# ocr-env_docx\Scripts\activate
# python hwpToDocx.py

# 입력 폴더와 출력 폴더 설정
input_folder = r"D:\yunzi\academy\기출hwp"
output_folder = r"D:\yunzi\academy\기출문제Docx"

# 출력 폴더가 없으면 생성
os.makedirs(output_folder, exist_ok=True)

# Word Application 객체 생성 (한 번만)
word = win32com.client.Dispatch("Word.Application")
word.Visible = False

# hwp 파일들 반복 처리
for filename in os.listdir(input_folder):
    if filename.lower().endswith('.hwp'):
        input_path = os.path.join(input_folder, filename)
        output_filename = os.path.splitext(filename)[0] + '.docx'
        output_path = os.path.join(output_folder, output_filename)

        print(f"▶ 변환 중: {filename}")

        # HWP 파일 실행
        subprocess.Popen(['start', '', input_path], shell=True)

        # 파일 열리는 시간 대기 (필요시 조절)
        time.sleep(3)

        # 전체 선택 및 복사
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.3)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(1)

        # 새 Word 문서에 붙여넣기
        doc = word.Documents.Add()
        doc.Content.Paste()
        doc.SaveAs(output_path, FileFormat=16)  # FileFormat=16: docx
        doc.Close()

         # 한글 종료 (Alt+F4)
        pyautogui.hotkey('alt', 'f4')
        time.sleep(1)  # 창 닫힐 시간 대기


        print(f"✅ 완료: {output_filename}")

# Word 종료
word.Quit()

print("\n🎉 모든 변환이 완료되었습니다.")