import requests
import bs4
from selenium import webdriver
from selenium.webdriver.chrome.service import Service

# 크롬드라이버 경로
service = Service('./drivers/chromedriver.exe')
driver = webdriver.Chrome(service=service)

driver.get("https://www.google.com")
print(driver.title)
driver.quit()

