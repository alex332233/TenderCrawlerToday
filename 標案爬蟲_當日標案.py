#!/usr/bin/env python
# coding: utf-8

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.alert import Alert
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import datetime
import os

# 設定 Chrome Driver的執行檔路徑
options = Options()
options.chrome_executable_path = (
    r"C:\Users\USER\Projects\pengFE\Python\chromedriver.exe"
)

# chromedriver_autoinstaller.install()  # Check if the current version of chromedriver exists
# and if it doesn't exist, download it automatically,
# then add chromedriver to path

# 建立Driver物件實體，用程式操作瀏覽器操作
driver = webdriver.Chrome(options=options)
print("start chrome...")
# init_date = datetime.date(2022, 11, 27)

driver.get("https://web.pcc.gov.tw/prkms/tender/common/basic/indexTenderBasic")
tenderName = driver.find_element(By.NAME, "tenderName")
print("get tender name...")
# tenderName.select_by_visible_text("創業") #輸入標案查尋關鍵字
tenderName.clear()  # 清除輸入框內容
tenderName.send_keys("創業")  # 需要輸入string
# tenderName.send_keys(Keys.ENTER) #輸入完按Enter關掉日曆
print("send keys...")
search = driver.find_element(
    By.XPATH,
    '//*[@id="basicTenderSearchForm"]/table/tbody/tr[11]/td/div/a',
).click()
print("search...")
# //*[@id="basicTenderSearchForm"]/table/tbody/tr[10]/td/div/a
html = driver.page_source

df = pd.read_html(html)[4]
print(df)

view_buttons = driver.find_elements(
    By.XPATH,
    '//*[@id="tpam"]/tbody/tr/td[10]/div/a | //*[@id="tpam"]/tbody/tr[position() > 1]/td[10]/div/a',
)
print("view button...")

# 檢查是否找到連結
if len(view_buttons) > 0:
    # 如果只有一列搜尋結果
    if len(view_buttons) == 1:
        # 取得該按鈕的 href 屬性
        href = view_buttons[0].get_attribute("href")
        print(href)
        # 將 href 加到結果表格的最後一欄
        df.loc[0, "連結"] = href
    else:
        # 如果有多列搜尋結果，使用 for 迴圈逐一處理
        for i, view_button in enumerate(view_buttons):
            href = view_button.get_attribute("href")
            print(href)
            df.loc[i, "連結"] = href
    # 在這裡加入您要執行的動作
else:
    # 如果沒有找到連結，就執行其他動作
    print("找不到連結")
    df.loc[0, "連結"] = "無連結"

df["類別"] = "創業"
result = df

# keywords = ['新創','職業體驗','青年','競賽']
keywords = [
    "新創",
    "職業體驗",
    "青年",
    "競賽",
    "課程",
    "輔導",
    "講座",
    "孵化",
    "投資",
    "Demoday",
    "媒合",
    "創育",
    "育成",
    "培育",
    "人才",
    "加速",
    "職涯發展",
    "市集",
    "市場",
    "社會創新",
    "社會企業",
]
print("list")
for keyword in keywords:
    tenderName = driver.find_element(By.NAME, "tenderName")
    tenderName.clear()  # 清除輸入框內容
    tenderName.send_keys(keyword)  # 需要輸入string

    search = driver.find_element(
        By.XPATH,
        '//*[@id="basicTenderSearchForm"]/table/tbody/tr[11]/td/div/a',
    ).click()
    print(keyword)
    #     search.click()
    html = driver.page_source

    df = pd.read_html(html)[4]
    print(df)

    # 找到所有的檢視按鈕
    view_buttons = driver.find_elements(
        By.XPATH,
        '//*[@id="tpam"]/tbody/tr/td[10]/div/a | //*[@id="tpam"]/tbody/tr[position() > 1]/td[10]/div/a',
    )

    # 檢查是否找到連結
    if len(view_buttons) > 0:
        #         # 如果找到連結，就點擊第一個連結
        #         view_buttons[0].click()

        # 如果只有一列搜尋結果
        if len(view_buttons) == 1:
            # 取得該按鈕的 href 屬性
            href = view_buttons[0].get_attribute("href")
            print(href)
            # 將 href 加到結果表格的最後一欄
            df.loc[0, "連結"] = href
        else:
            # 如果有多列搜尋結果，使用 for 迴圈逐一處理
            for i, view_button in enumerate(view_buttons):
                href = view_button.get_attribute("href")
                print(href)
                df.loc[i, "連結"] = href
    # 在這裡加入您要執行的動作
    else:
        # 如果沒有找到連結，就執行其他動作
        df.loc[0, "連結"] = "無連結"
        print("找不到連結")

    df["類別"] = keyword
    result = pd.concat([result, df], axis=0, ignore_index=True)

    # 先拿到 column list
cols = result.columns.to_list()
# 交換 column 順序
cols = cols[-1:] + cols[:-1]
# 套用新的 column 順序
result = result[cols]
today = datetime.date.today().strftime("%Y-%m-%d")
directory = "C:/Users/USER/Documents/Projects/pengFE/Python/標案紀錄"
if not os.path.exists(directory):
    os.makedirs(directory)
result.to_csv(directory + "/當日標案_" + today + ".csv", encoding="utf_8_sig")
