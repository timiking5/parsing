from selenium import webdriver
from vtb import vtb
from alfa import alfa
from rshb import rshb
from gazprom import gazprom
from mkb import mkb


# service=Service(ChromeDriverManager().install())  -  для будущих разрабов
# options.add_argument('--start-maximized')  -  для некоторых браузеров/cайтов это лучше

options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--incognito')
options.add_argument('window-size=1920,1080')
options.add_argument('--headless')

PATH = "C:\\Program Files (x86)\\chromedriver.exe"  # путь до браузера
# Важно, чтобы оно совпадало с версией браузера на компе!!
browser = webdriver.Chrome(executable_path=PATH, chrome_options=options)

vtb(browser)
print('=' * 30)
alfa(browser)
print('=' * 30)
rshb(browser)
print('=' * 30)
gazprom(browser)
print('=' * 30)
mkb(browser)
print('=' * 30)
# sber(browser)
"""
У сбера на сайте, может выскочить страница "Была ли предыдущая страница полезной?"
либо он может просто запретить доступ роботу на сайт в режиме без окна, поэтому его
тяжело спарсить. Лучше это делать раз в день, тогда подозрительная активность может быть
не засечена
"""


browser.quit()
