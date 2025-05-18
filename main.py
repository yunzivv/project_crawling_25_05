import requests
import bs4
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import time

# 크롬 드라이버 설정
# options = webdriver.ChromeOptions()
# options.add_argument("--headless")
# options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36")
# options.add_argument("--disable-blink-features=AutomationControlled")
# service = Service('./drivers/chromedriver.exe')
# driver = webdriver.Chrome(service=service)

# 설정
options = webdriver.ChromeOptions()
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/117 Safari/537.36")
options.add_argument("--disable-blink-features=AutomationControlled")
driver = webdriver.Chrome(options=options)

# 자격/취득 포함된 링크 저장용
qualified_links = []

# IP 주소 확인
import requests
ip = requests.get("https://api.ipify.org").text
print(f"현재 IP 주소: {ip}")


# 키워드
keywords = ["채용"]

# 잡코리아 첫 페이지
base_url = "https://www.jobkorea.co.kr/recruit/joblist?page="

# 크롬 옵션 설정
options = webdriver.ChromeOptions()
options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
options.add_argument("--disable-blink-features=AutomationControlled")
driver = webdriver.Chrome(options=options)

# 링크 저장
matched_links = []

# 1~3페이지 예시로 진행
for page in range(1, 4):
    print(f"\n🔄 [{page} 페이지 검색 중]")
    driver.get(base_url + str(page))
    time.sleep(3)

    # 공고 목록의 href만 먼저 추출
    job_links = driver.find_elements(By.CSS_SELECTOR, "td.tplTit > div.titBx > strong > a")
    hrefs = []

    for link in job_links:
        try:
            href = link.get_attribute("href")
            if href:
                hrefs.append(href)
        except Exception as e:
            print("❌ 링크에서 href 추출 실패:", e)

    # 각 href에 접속해서 '채용' 키워드 찾기
    for href in hrefs:
        try:
            driver.get(href)
            time.sleep(2)
            page_text = driver.page_source

            if any(keyword in page_text for keyword in keywords):
                print(f"✅ {href}")
                matched_links.append(href)
        except Exception as e:
            print(f"❌ 링크 문제: {href}", e)

# 결과 출력
print(f"\n📌 총 {len(matched_links)}개의 공고에서 키워드 발견됨.")
# for url in matched_links:
#     print(url)

driver.quit()