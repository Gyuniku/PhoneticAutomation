from ctypes import c_short

from selenium import webdriver
import chromedriver_binary
import os
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import re
import pandas as pd
from pykakasi import kakasi
import traceback

# WebDriver設定
# chromeのWebDriverのパス
driver_path = 'testpath\webdriver\chrome_78.0\chromedriver_win32\chromedriver.exe'
# Chromeユーザの設定フォルダのパス
chrome_option = '--user-data-dir=C:\\Users\\~~~~~\\AppData\\Local\\Google\\Chrome\\User Data'
# テスト用Chromeユーザの設定フォルダ名
chrome_testuser_settings_dir = '--profile-directory=folder_name'
# 初期画面
default_page = 'https://outlook.office.com/people/'
# 検索ボックスのXPath
find_box_id = "//div[@id='owaSearchBox']/div/div/div/div/input"
# 検索条件
find_keyword = ''
# 検索ボタンのXPath
find_execute_button = "//div[2]/div/div/div/button/span/i"
# 検索結果0件メッセージが含まれる箇所のXPath
find_result_area = "//div[@id='app']/div/div[2]/div/div[2]/div/div[2]/div/div/div/div[2]/p"
# 姓名表示がされている箇所のXPath
name_area = "//div[@id='app']/div/div[2]/div/div[2]/div/div[2]/div/div[3]/div/div/div/div/div/header/div[2]/div[2]/div/div"
# Chromeのテストユーザ名
test_user = 'selenium_chrome_user2'

# データ設定
# 読み込むCSVのファイル名
csv_import_file = './import.csv'


def wait_cond(driver) -> bool:
    print("判定開始")
    try:
        if '別のキーワードをお試しください' in driver.page_source:
            print("検索結果0件")
            return True
        if driver.find_element_by_xpath(name_area):
            print("検索結果あり")
            return True
    except Exception as e:
        traceback.format_exc()
        return False
    return False


# インポートデータ読み込み
df = pd.read_csv(csv_import_file, sep=',', encoding='UTF-8')
print(df)

# 既存のChromeの設定を読み込む（認証情報含む）
os.makedirs('selenium_chrome_user2', exist_ok=True)

# 結果格納用データフレーム
result_df = pd.DataFrame(index=[], columns=['outlook_name', 'kakasi_name'])

try:
    options = webdriver.ChromeOptions()
    options.add_argument('--user-data-dir=' + test_user)
    # options.add_argument(chrome_option)
    options.add_argument(chrome_testuser_settings_dir)

    # 初期ページへ遷移
    driver = webdriver.Chrome(options=options, executable_path=driver_path)
    driver.get(default_page)

    # 表示されるまで待機
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, find_box_id))
    )

    # データを1件ずつ処理
    for index, row in df.iterrows():
        # Outlookに登録されているふりがな
        outlook_name = ''
        # PyKakasiを用いて出力したふりがな
        kakasi_name = ''

        try:
            # 検索ボックスを押して検索条件を入れる
            find_box = driver.find_element_by_xpath(find_box_id)
            find_box.click()
            # 検索ボックスの中身を空にする（クリックしてキーボード操作で削除しないと反応しない）
            find_box.send_keys(Keys.CONTROL + 'a')
            find_box.send_keys(Keys.DELETE)
            find_box.send_keys(row['name'])
            driver.find_element_by_xpath(find_execute_button).click()

            # 検索結果が表示されるまで待機
            # 表示されない場合(検索結果なし)の場合は次に進む
            WebDriverWait(driver, 10).until(wait_cond)

            # Outlookに登録されている名前
            outlook_name = driver.find_element_by_xpath(name_area).text
            # 正規表現で英字とスペースのみにする
            pattern1 = re.compile(r"[^a-z\s]+")
            outlook_name = pattern1.sub('', outlook_name)
            # 先頭のスペースは削除
            pattern2 = re.compile(r"^ +")
            outlook_name = pattern2.sub('', outlook_name)
            # 姓と名を入れ替える
            pattern3 = re.compile(r"(?P<g1>.+)\s(?P<g2>.+)")
            outlook_name = pattern3.sub(r'\g<g2> \g<g1>', outlook_name)

        except Exception as e:
            # 検索結果なしのときは空白を入れて次へ
            print(e)
            outlook_name = ''
            print("検索結果なし")

        finally:
            try:
                # PyKakasiを用いた処理
                kks = kakasi()
                # 入力された名前をカタカナにする
                kks.setMode('H', 'K')
                kks.setMode('K', 'K')
                kks.setMode('J', 'K')
                kks.setMode('a', 'K')
                conv = kks.getConverter()
                kakasi_name = conv.do(row['name'])

                # CSVに追加
                series = pd.Series([outlook_name, kakasi_name], index=['outlook_name', 'kakasi_name'])
                result_df = result_df.append(series, ignore_index=True)
                print(result_df)
            except Exception as e:
                print("エラーが発生しました。次の文字列を処理します。")
                print(traceback.format_exc())

finally:
    driver.close()
    # CSVを出力
    # 読み込んだデータに結果を追加
    df = pd.concat([df, result_df], axis=1)
    print(df)
    df.to_csv('export.csv', index=False)


