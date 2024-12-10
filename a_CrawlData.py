import time
from datetime import datetime

import xlwt
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
from selenium.webdriver.common.by import By
import re


# 0. Function List
# scrape data from the content page
def content_analyse(driver, news_link):
    driver.execute_script(f"window.open('{news_link}','_blank');")
    new_window = [window for window in driver.window_handles if window != original_window][0]
    driver.switch_to.window(new_window)
    news_html = driver.page_source
    news_soup = BeautifulSoup(news_html, 'html.parser')
    global news_content
    news_content = news_soup.find(class_='content_area')
    if news_content:
        news_content = news_soup.find(class_='content_area').text
    else:
        news_content = news_soup.find(class_='text_area')
        if news_content:
            news_content = news_soup.find(class_='text_area').text
        else:
            news_content = 0
    # time.sleep(3)
    driver.close()
    driver.switch_to.window(original_window)

    return 0


# handle the publication time
def times_standard(string):
    date_format = '%Y/%m/%d'
    pattern = r'\d{4}/\d{2}/\d{2}'
    date = re.search(pattern, string).group()
    date_object = datetime.strptime(date, date_format)

    return date_object


def save_data():
    # save in the list
    data = {
        "news_title": news_title,
        "news_brief": news_brief,
        "news_link": news_link,
        "news_time": news_time,
        "news_content": news_content
    }
    data_list.append(data)
    # save in the csv
    global n
    sheet.write(n, 0, news_title)
    sheet.write(n, 1, news_brief)
    sheet.write(n, 2, news_link)
    sheet.write(n, 3, data["news_time"].strftime("%Y-%m-%d"))
    sheet.write(n, 4, news_content)
    n += 1

    return data_list


# 1. open the webpage
url = "https://news.cctv.com/world/?spm=C94212.P4YnMod9m2uD.EWZW7h07k3Vs.3"
path_chrome_testing = r"F:\chrometestingwin64\chrome-win64\chrome.exe"
path_chrome_driver = r"F:\chrometestingwin64\chrome-win64\chromedriver.exe"
options = webdriver.ChromeOptions()
options.binary_location = path_chrome_testing
options.add_experimental_option('detach', True)  # To avoid close the page automatically
service = Service(path_chrome_driver)
driver = webdriver.Chrome(service=service, options=options)
driver.get(url)
original_window = driver.current_window_handle

# 2. load all items
end_load = driver.find_elements(By.CLASS_NAME, "more.cur.nomore")
load = driver.find_element(By.XPATH, '//*[@id="SUBD1565174589203138"]/div[3]/div/div[3]/p')
while not end_load:
    load.click()
    end_load = driver.find_elements(By.CLASS_NAME, "more.cur.nomore")
    print("not ending")
    time.sleep(1)
print("All loaded")

# 3. scrape news & save in csv
data_list = []  # save the entire news list
data = {}  # save one news item

# 3.1 create a csv file
excl = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = excl.add_sheet('China_international_news', cell_overwrite_ok=True)
sheet.write(0, 0, 'title')
sheet.write(0, 1, 'brief')
sheet.write(0, 2, 'link')
sheet.write(0, 3, 'time.')  # string
sheet.write(0, 4, 'content')
n = 1  # csv driver

html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')
# 3.2 head news
head_news = soup.find(class_='right_text')
news_title = head_news.find('h3').text
news_brief = head_news.find('p').text
news_link = head_news.find('h3').find('a')['href']
news_time = times_standard(news_link)
content_analyse(driver, news_link)
save_data()

# 3.3 list news
news_list = soup.find_all('li', style="display: list-item;")
for news in news_list:
    news_title = news.find(class_='text_con').find(class_='title').text
    news_brief = news.find(class_='text_con').find(class_='brief').string
    news_link = news.find(class_='titlekapain').find('a')['href']
    news_time = times_standard(news_link)
    content_analyse(driver, news_link)
    save_data()
excl.save("ChinaNewsIntList.xls")
