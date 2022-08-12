from selenium import webdriver
from vtb import vtb
from alfa import alfa
from rshb import rshb
from gazprom import gazprom
from mkb import mkb
from sber import sber


# service=Service(ChromeDriverManager().install())  -  для будущих разрабов
# options.add_argument('--start-maximized')  -  для некоторых браузеров/cайтов это лучше

chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument('--ignore-certificate-errors')
chrome_options.add_argument('--incognito')
chrome_options.add_argument('window-size=1920,1080')
# chrome_options.add_argument('--headless')

firefox_options = webdriver.FirefoxOptions()
# firefox_options.add_argument('--headless')
firefox_options.add_argument('--ignore-certificate-errors')
firefox_options.add_argument('--incognito')

PATH_GOOGLE = "C:\\Program Files (x86)\\chromedriver.exe"  # путь до браузера
PATH_GECKO = "C:\\Program Files (x86)\\geckodriver.exe"
# Важно, чтобы оно совпадало с версией браузера на компе!!

browser_chrome = webdriver.Chrome(executable_path=PATH_GOOGLE, chrome_options=chrome_options)
browser_firefox = webdriver.Firefox(executable_path=PATH_GECKO, options=firefox_options)

vtb(browser_chrome)
print('=' * 30)
alfa(browser_chrome)
print('=' * 30)
rshb(browser_chrome)
print('=' * 30)
gazprom(browser_chrome)
print('=' * 30)
mkb(browser_chrome)
print('=' * 30)
sber(browser_firefox)
"""
У сбера на сайте, может выскочить страница "Была ли предыдущая страница полезной?"
либо он может просто запретить доступ роботу на сайт в режиме без окна, поэтому его
тяжело спарсить. Лучше это делать раз в день, тогда подозрительная активность может быть
не засечена
"""

browser_chrome.quit()
browser_firefox.quit()
