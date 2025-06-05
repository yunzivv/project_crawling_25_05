import pandas as pd
from difflib import get_close_matches

# ocr-env_preprocess
# 불러오기
collected = pd.read_excel("job-db.xlsx")   # 수집된 자격증 데이터
official = pd.read_excel("자격증_정렬완료.xlsx")     # 정식 자격증 목록

# 이름만 추출
collected_names = collected['자격증명'].dropna().unique()
official_names = official['자격증명'].dropna().unique()

# 유사한 국가자격증 매핑
matches = {}
for name in collected_names:
    match = get_close_matches(name, official_names, n=1, cutoff=0.8)
    matches[name] = match[0] if match else None

# 결과 저장
result = pd.DataFrame(list(matches.items()), columns=['원래명칭', '정식명칭'])
result.to_excel("자격증_매핑결과.xlsx", index=False)
