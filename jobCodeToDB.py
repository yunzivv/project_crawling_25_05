import pandas as pd
import re

# 엑셀 불러오기
df = pd.read_excel("qnet_certifications.xlsx")

# 등급 우선순위
priority = {"기술사": 1, "기능장": 2, "기사": 3, "산업기사": 4, "기능사": 5}

# 정확한 등급 추출 함수
def extract_grade(name):
    for grade in priority:
        if isinstance(name, str) and re.search(f'{grade}$', name):
            return grade
    return None

# 자격증군 추출 (등급 제거한 앞부분)
def extract_group(name):
    grade = extract_grade(name)
    if grade and isinstance(name, str):
        return name.replace(grade, '')
    return name

# 적용
df['seriesNm'] = df['자격증명'].apply(extract_grade)
df['자격증군'] = df['자격증명'].apply(extract_group)
df['등급순위'] = df['seriesNm'].map(priority)

# 정렬
df_sorted = df.sort_values(by=['자격증군', '등급순위'])

# 저장
df_sorted.to_excel("자격증_정렬완료.xlsx", index=False)