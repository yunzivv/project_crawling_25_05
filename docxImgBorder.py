import zipfile
import os
from docx import Document
from PIL import Image, ImageOps
import io
import shutil

# 파일 경로 지정
docx_path = "가스기사20200606.docx"           # 원본 docx 파일명
output_path = "테두리적용.docx"   # 최종 결과물 저장 위치
temp_dir = "temp_docx"           # 임시 압축 해제 디렉토리

# 1. 임시 디렉토리 초기화
if os.path.exists(temp_dir):
    shutil.rmtree(temp_dir)
os.makedirs(temp_dir, exist_ok=True)

# 2. docx 파일 압축 해제 (docx는 zip 구조)
with zipfile.ZipFile(docx_path, 'r') as zip_ref:
    zip_ref.extractall(temp_dir)

# 3. 이미지 처리 (검은색 테두리 입히기)
media_dir = os.path.join(temp_dir, "word", "media")
if os.path.exists(media_dir):
    for img_name in os.listdir(media_dir):
        if img_name.lower().endswith((".png", ".jpg", ".jpeg")):
            img_path = os.path.join(media_dir, img_name)
            img = Image.open(img_path)
            img_with_border = ImageOps.expand(img, border=3, fill='black')  # 테두리 3px
            img_with_border.save(img_path)
else:
    print("⚠️ 이미지(media) 폴더가 없습니다.")

# 4. 다시 zip으로 압축 후 .docx로 저장
shutil.make_archive("temp_modified", 'zip', temp_dir)
shutil.move("temp_modified.zip", output_path)
print(f"✅ 처리 완료: {output_path}")