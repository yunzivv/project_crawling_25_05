import pyautogui
import time
import win32com.client
import os
import subprocess

# ocr-env_docx\Scripts\activate
# python hwpToDocx.py

# 파일 경로
input_path = r"D:\yunzi\academy\기출문제hwp\건축기사20210307.hwp"
output_path = r"D:\yunzi\academy\hwpToDocx\건축기사20210307.docx"

# 한글 앱으로 직접 실행
subprocess.Popen(['start', '', input_path], shell=True)

# 파일 열리는 시간 기다리기
time.sleep(3)

# 복사
pyautogui.hotkey('ctrl', 'a')
time.sleep(0.2)
pyautogui.hotkey('ctrl', 'c')
time.sleep(1)

# 워드 붙여넣기
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
doc = word.Documents.Add()
doc.Content.Paste()
doc.SaveAs(output_path, FileFormat=16)
doc.Close()
word.Quit()

print("✅ 변환 완료:", output_path)