# encoding: utf-8
import sys
import time
import xlwt
from bs4 import BeautifulSoup
from selenium import webdriver


def main():
    tag_video_tile_parent_element = 'ytd-playlist-video-renderer'
    tag_video_title_id = "#video-title"

    tag_list_title_parent_element = 'yt-formatted-string'
    tag_list_title_class = '.yt-simple-endpoint'

    list_url = 'https://www.youtube.com/playlist?list=PLfrV5gSr2fGzZkGrE9dsshdpxzhcmKh6N'

    if len(sys.argv) != 1:
        list_url = sys.argv[1]

    browser = webdriver.Firefox(executable_path=r"geckodriver.exe")
    time.sleep(3)
    browser.get(list_url)
    time.sleep(3)
    for i in range(10):  # 進行十次
        browser.execute_script('window.scrollTo(0, document.getElementById("content").clientHeight);')  # 重複往下捲動
        time.sleep(1)  # 每次執行打瞌睡一秒

    listPageSource = browser.page_source
    soup = BeautifulSoup(listPageSource, "lxml")

    list_name = soup.select('{} a{}'.format(tag_list_title_parent_element, tag_list_title_class))
    workbook = xlwt.Workbook(encoding='utf-8')
    if len(list_name) > 0:
        book_sheet = workbook.add_sheet(list_name[0].text, cell_overwrite_ok=True)
    else:
        book_sheet = workbook.add_sheet('Sheet 1', cell_overwrite_ok=True)

    book_sheet.write(0, 0, 'Serial Number')
    book_sheet.write(0, 1, 'Title')

    counter: int = 1

    for title in soup.select('{} span{}'.format(tag_video_tile_parent_element, tag_video_title_id)):
        title_name = '{} {}\n'.format(str(counter), title.text[1:])
        print(title_name)
        book_sheet.write(counter, 0, counter)
        book_sheet.write(counter, 1, title.text)
        counter = counter+1

    workbook.save('{}.xls'.format(time.strftime("%Y-%m-%d_%H-%M-%S", time.localtime())))

main()