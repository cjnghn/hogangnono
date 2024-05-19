import os
import time
import traceback
from typing import List, Dict, Optional

import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager

os.environ['WDM_LOG'] = '0'


def setup_driver() -> webdriver.Chrome:
  chrome_options = webdriver.ChromeOptions()
  chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
  return webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)


def login(driver: webdriver.Chrome) -> Optional[List[dict]]:
  try:
    driver.get('https://hogangnono.com')
    dismiss_popup(driver)
    initiate_login(driver)
    print_login_prompt()
    wait_for_login(driver)
    return driver.get_cookies()
  except Exception as e:
    log_error("로그인 중 에러 발생", e)
    return None


def dismiss_popup(driver: webdriver.Chrome):
  click_element(driver, 'a[data-ga-event="intro,closeBtn"]')


def initiate_login(driver: webdriver.Chrome):
  click_element(driver, 'a.btn-login')


def print_login_prompt():
  print('\n[*] 로그인을 진행하세요 ')
  print('(웹 브라우저 창에서 로그인을 완료하면 자동으로 진행됩니다.)')


def wait_for_login(driver: webdriver.Chrome):
  WebDriverWait(driver, 60 * 5).until(EC.url_to_be('https://hogangnono.com/my'))
  print("=> 로그인이 완료되었습니다.")


def click_element(driver: webdriver.Chrome, css_selector: str, timeout: int = 10):
  element = WebDriverWait(driver, timeout).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, css_selector))
  )
  element.click()


def log_error(message: str, exception: Exception):
  print(f"[!] {message}:")
  print(str(exception))
  traceback.print_exc()


def create_session(cookies: List[dict]) -> requests.Session:
  session = requests.Session()
  for cookie in cookies:
    session.cookies.set(cookie['name'], cookie['value'])
  return session


def fetch_data(session: requests.Session, url: str, params: dict = None) -> dict:
  response = session.get(url, params=params)
  response.raise_for_status()
  return response.json()


def get_my_info(session: requests.Session) -> dict:
  return fetch_data(session, 'https://hogangnono.com/api/me')


def search(session: requests.Session, query: str) -> dict:
  search_url = f'https://hogangnono.com/api/v2/searches/new?query={query}'
  return fetch_data(session, search_url)


def extract_apt_id(search_data: dict) -> Optional[int]:
  try:
    return search_data['data']['matched']['apt']['list'][0]['id']
  except (KeyError, IndexError, TypeError):
    return None


def fetch_reviews(session: requests.Session, apt_id: int) -> List[Dict[str, str]]:
  reviews = []
  page = 1

  while True:
    review_data = request_reviews(session, apt_id, page)
    if not review_data.get('data'):
      break

    reviews.extend(format_reviews(review_data['data']['data']))
    page += 1

  return reviews


def request_reviews(session: requests.Session, apt_id: int, page: int) -> dict:
  review_url = f'https://hogangnono.com/api/v2/apts/{apt_id}/reviews'
  review_params = {'orderType': 1, 'page': page, 'showResidentReviewOnly': 0}
  return fetch_data(session, review_url, review_params)


def format_reviews(reviews: List[dict]) -> List[Dict[str, str]]:
  formatted_reviews = []
  for review in reviews:
    formatted_reviews.append({
      '닉네임': review.get('name', '익명'),
      '내용': review.get('content'),
      '도움돼요': review.get('countUp'),
      '작성날짜': review.get('date'),
      '댓글': format_comments(review.get('comments', []))
    })
  return formatted_reviews


def format_comments(comments: List[dict]) -> str:
  return '\n'.join([f"{comment['name']}: {comment['content']}" for comment in comments])


def retrieve_file_path() -> str:
  root = tk.Tk()
  root.withdraw()
  file_path = filedialog.askopenfilename()
  root.destroy()
  return file_path


def save_reviews_to_excel(file_path: str, search_queries: List[str], session: requests.Session):
  with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    for query in search_queries:
      process_query(writer, query, session)


def process_query(writer: pd.ExcelWriter, query: str, session: requests.Session):
  search_data = search(session, query)
  apt_id = extract_apt_id(search_data)
  if apt_id:
    print(f"=> {query} 검색 결과: 아파트 ID {apt_id}")
    reviews = fetch_reviews(session, apt_id)
    reviews_df = pd.DataFrame(reviews)
    reviews_df.to_excel(writer, sheet_name=query, index=False)
  else:
    print(f"=> {query} 검색 결과: 결과가 없습니다.")


def main():
  try:
    display_initial_messages()

    with setup_driver() as driver:
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

    excel_file_path = retrieve_file_path()
    if not excel_file_path:
      print("[!] 파일을 선택하지 않았습니다.")
      return

    search_queries = load_search_queries_from_excel(excel_file_path)
    save_reviews_to_excel(excel_file_path, search_queries, session)
    print("\n=> 🐎 검색이 모두 완료되었습니다.")

  except Exception as e:
    log_error("오류 발생", e)


def display_initial_messages():
  print('[!] 주의: 실행 시 크롬 드라이버 설치를 위해 시간이 소요될 수 있습니다.')
  print('=> 다음 메시지가 출력될 때까지 기다려주세요.')
  print('=> 브라우저가 자동으로 열리면 로그인을 진행하세요.')
  time.sleep(3)


def load_search_queries_from_excel(file_path: str) -> List[str]:
  df = pd.read_excel(file_path, header=None)
  return df.iloc[:, 0].dropna().tolist()


if __name__ == '__main__':
  main()
