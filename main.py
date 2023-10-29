import time
import os
os.environ['WDM_LOG'] = '0'

import requests
import pandas as pd
import tkinter as tk
import traceback

from typing import List
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager


def login(driver: WebDriver) -> List[dict]:
  try:
    driver.get('https://hogangnono.com')
    popup_button = WebDriverWait(driver, 10).until(
      EC.presence_of_element_located((
        By.CSS_SELECTOR, 'a[data-ga-event="intro,closeBtn"]'))
    )
    popup_button.click()

    login_button = WebDriverWait(driver, 10).until(
      EC.presence_of_element_located((
        By.CSS_SELECTOR, 'a.btn-login'))
    )
    login_button.click()

    print('\n[*] 로그인을 진행하세요 ')
    print('(웹 브라우저 창에서 로그인을 완료하면 자동으로 진행됩니다.)')

    WebDriverWait(driver, 60 * 5).until(
      EC.url_to_be('https://hogangnono.com/my')
    )
    print("=> 로그인이 완료되었습니다.")

    cookies = driver.get_cookies()
    return cookies
  
  except Exception as e:
    print("[!] 로그인 중 에러 발생:")
    print(str(e))
    traceback.print_exc()  # 에러 정보를 출력
    return []


def create_session(cookies: List[dict]) -> requests.Session:
  session = requests.Session()
  for cookie in cookies:
    session.cookies.set(cookie['name'], cookie['value'])

  return session


def get_my_info(session: requests.Session) -> dict:
  response = session.get('https://hogangnono.com/api/me')
  return response.json()


def search(session: requests.Session, query: str) -> dict:
  search_url = f'https://hogangnono.com/api/v2/searches/new?query={query}'
  search_response = session.get(search_url)
  return search_response.json()


def get_apt_id(search_data: dict) -> int:
  if 'data' not in search_data or search_data['data'] is None:
    return None

  return search_data['data']['matched']['apt']['list'][0]['id']


def get_reviews(session: requests.Session, apt_id: int) -> List[dict]:
  review_url = f'https://hogangnono.com/api/v2/apts/{apt_id}/reviews'
  review_params = {
    'orderType': 1,
    'page': 1,
    'showResidentReviewOnly': 0
  }

  reviews_list = []
  while True:
    review_response = session.get(review_url, params=review_params)
    review_data = review_response.json()

    if 'data' not in review_data or not review_data['data']:
      break

    reviews = review_data['data'].get('data', [])
    for review in reviews:
      formatted_review = {
        '닉네임': review.get('name', '익명'),
        '내용': review.get('content'),
        '도움돼요': review.get('countUp'),
        '작성날짜': review.get('date'),
        '댓글': '\n'.join(
          [f"{comment['name']}: {comment['content']}" for comment in review.get('comments', [])])
          if review.get('comments') else None
      }
      reviews_list.append(formatted_review)

    review_params['page'] += 1

  return reviews_list


def retrieve_file_path() -> str:
  root = tk.Tk()
  root.withdraw()
  file_path = filedialog.askopenfilename()
  root.destroy()

  return file_path


def main():
  try:
    print('[!] 주의: 실행 시 크롬 드라이버 설치를 위해 시간이 소요될 수 있습니다.')
    print('=> 다음 메시지가 출력될 때까지 기다려주세요.')
    print('=> 브라우저가 자동으로 열리면 로그인을 진행하세요.')
    time.sleep(3)

    # Chrome 드라이버 옵션 설정
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

    with webdriver.Chrome(service=ChromeService(ChromeDriverManager().install(), options=chrome_options)) as driver:
      cookies = login(driver)

    if not cookies:
      print("로그인에 실패하였습니다.")
      return

    session = create_session(cookies)
    my_info = get_my_info(session)
    if not my_info:
      print("[!] 내 정보를 가져오는데 실패하였습니다.")
      return

    print(f"=> {my_info['data']['nickname']}님 환영합니다.")
    
    print("\n[*] 검색할 파일을 선택하세요.")
    excel_file_path = retrieve_file_path()
    if not excel_file_path:
      print("[!] 파일을 선택하지 않았습니다.")
      return

    print("\n[*] 검색을 시작합니다.")
    df = pd.read_excel(excel_file_path, header=None)
    search_queries = df.iloc[:, 0].dropna().tolist()

    with pd.ExcelWriter(excel_file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
      for query in search_queries:
        search_data = search(session, query)
        apt_id = get_apt_id(search_data)
        if apt_id is not None:
          print(f"=> {query} 검색 결과: 아파트 ID {apt_id}")
          reviews = get_reviews(session, apt_id)
          reviews_df = pd.DataFrame(reviews)

          reviews_df.to_excel(writer, sheet_name=query, index=False)
        else:
          print(f"=> {query} 검색 결과: 결과가 없습니다.")
    
    print("\n=> 🐎 검색이 모두 완료되었습니다.")

  except Exception as e:
    print("[!] 오류 발생:")
    print(str(e))
    traceback.print_exc()


if __name__ == '__main__':
  main()
