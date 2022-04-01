import glob
import os
import sys
import time
from os import listdir
from os.path import isfile, isdir, join

import urllib
from urllib.request import urlretrieve

import requests
from bs4 import BeautifulSoup

headers = {
    'user-agent': 'Mozilla/5.0 (Macintosh Intel Mac OS X 10_13_4) AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/66.0.3359.181 Safari/537.36'}

web_text = "http://www.eyny.com/home.php?mod=space&uid=13320267&do=blog&view=me&classid=54423&from=space&page="
# web_num = range(1, 91)
web_num = range(1, 95)
web = web_text + str(web_num)
MovieName = "我來自北京之瑪尼堆的秋天"
Result = 0


for i in web_num:
    web_num_count = str(i)
    web = web_text + str(web_num_count)
    res = requests.get(web, headers=headers)
    soup = BeautifulSoup(res.text, 'html.parser')
    WebPage = str(i)
    print(WebPage)
    soup_str = str(soup)

    if MovieName in soup_str:
        print("找到關鍵字[" + MovieName+"]: 在第" + WebPage + "頁")
        Result = 1
        break

    time.sleep(0.6)

if Result == 0:
    print("都沒找到")



















# headers = {
#     'user-agent': 'Mozilla/5.0 (Macintosh Intel Mac OS X 10_13_4) AppleWebKit/537.36 (KHTML, like Gecko) '
#                   'Chrome/66.0.3359.181 Safari/537.36'}
#
# web = "http://www.eyny.com/home.php?mod=space&uid=11195773&do=blog&view=me&classid=54282&from=space&page=1"
#
# MovieName = "怨恨3"
# res = requests.get(web, headers=headers)
# soup = BeautifulSoup(res.text, 'html.parser')
# soup_text = str(soup)
#
# if "門神之決戰蛟龍" in soup_text:
#     print("YYYYY")
# else:
#     print("都沒找到")
#     print(soup_text)