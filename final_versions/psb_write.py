import time
import re
import sys
import os
import warnings
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common import action_chains
from selenium.webdriver.remote.webelement import WebElement


def write_strong(driver: webdriver, active_sheet, row_f):
    """
    Одна таблица нестандартнее другой. Пришлось опять же выкручиваться.
    Функция записывает таблицу.
    :param driver: browser
    :param active_sheet: Эксель таблицу, куда записывать
    :param row_f: ряд, начиная с которого надо записывать
    :return:
    """
    driver.get('https://www.psbank.ru/Personal/Saving/Strong_bid')
    for i in range(2):
        scroll(driver)
        time.sleep(1)
    main = driver.find_element(By.CLASS_NAME, 'wrapper')
    table = main.find_element(By.TAG_NAME, 'table')
    headers = table.find_element(By.TAG_NAME, 'thead').find_elements(By.TAG_NAME, 'th')
    rows = table.find_element(By.TAG_NAME, 'tbody').find_elements(By.TAG_NAME, 'tr')
    active_sheet[f'{chr(65)}{row_f}'] = headers[1].text
    row_f += 1
    for i in range(len(rows)):
        tds = rows[i].find_elements(By.TAG_NAME, 'td')
        if i == 0:
            for j in range(len(tds) - 1, -1, -1):
                if tds[j].text.isdigit():
                    active_sheet[f'{chr(64 + j)}{row_f}'] = tds[j].text
                else:
                    active_sheet[f'{chr(64 + j)}{row_f}'] = 'Базовая'
                    active_sheet[f'{chr(64 + j)}{row_f + 5}'] = 'Повышенная'
                    break
        elif i == 1:
            continue
        else:
            basic = 0
            bonus = 0
            for j in range(len(tds)):
                if j == 0:
                    active_sheet[f'A{row_f}'] = tds[j].text
                    active_sheet[f'A{row_f + 5}'] = tds[j].text
                elif j % 2 == 0:
                    active_sheet[f'{chr(66 + basic)}{row_f}'] = tds[j].text
                    basic += 1
                elif j % 2 == 1:
                    active_sheet[f'{chr(66 + bonus)}{row_f + 5}'] = tds[j].text
                    bonus += 1
        row_f += 1


class Kostil:
    text = ''

    def get_attribute(self, asd):
        pass


def use_bar(main: WebElement, request):
    """
    Использует вводное поле
    :param main:
    :param request: то, что надо передать в поле
    :return:
    """
    bar = main.find_element(By.CSS_SELECTOR, '.range-label')
    bar.click()
    bar.send_keys(request, Keys.ESCAPE)


def scroll(driver):
    """
    Сайт не скролиться по-обычному, пришлось делать это так.
    :param driver:
    :return:
    """
    webdriver.common.action_chains.ActionChains(driver).scroll_by_amount(0, 1000).perform()


def use_slider(driver: webdriver, main: WebElement, offset):
    """
    Передвигает слайдер. При том слайдер проваливается в заготовленные периоды.
    :param driver: browser
    :param main:
    :param offset: сдвиг (можно эксперементировать)
    :return:
    """
    slider = main.find_elements(By.CSS_SELECTOR, '.irs-handle.single')[1]
    webdriver.common.action_chains.ActionChains(driver).drag_and_drop_by_offset(slider, offset, 0).perform()


def write_table(table: WebElement, active_sheet, row_f):
    """
    Забирает таблицы, но они довольно убогие. Пришлось использовать костыль
    :param table:
    :param active_sheet:
    :param row_f:
    :return:
    """
    head_rows = table.find_element(By.TAG_NAME, 'thead').find_elements(By.TAG_NAME, 'tr')
    info_rows = table.find_element(By.TAG_NAME, 'tbody').find_elements(By.TAG_NAME, 'tr')
    ths_1 = head_rows[0].find_elements(By.TAG_NAME, 'th')
    ths_2 = head_rows[2].find_elements(By.TAG_NAME, 'th')
    if len(ths_2) == 0:
        ths_2 = head_rows[2].find_elements(By.TAG_NAME, 'td')
    for i in range(1, 3):
        if i == 1:
            for j in range(len(ths_1) - 1):
                active_sheet[f'{chr(76 + j)}{row_f - 7}'] = ' '.join(ths_1[j].text.split('\n'))
        elif i == 2:
            for j in range(len(ths_2)):
                active_sheet[f'{chr(76 + j + len(ths_1) - 1)}{row_f - 7}'] = ths_2[j].text

    tds = [i.find_elements(By.TAG_NAME, 'td') for i in info_rows]
    for i in range(len(tds)):
        for j in range(len(tds[i])):
            active_sheet[f'{chr(76 + j)}{row_f - 6 + i}'] = tds[i][j].text

            if tds[i][j].get_attribute('rowspan') in ["2", "3"]:
                for h in range(1, len(tds)):
                    kostil = Kostil()
                    tds[h].insert(j, kostil)


def parse_psb(main: WebElement):
    """
    Забирает процент (он будет вторым представителем класса '.calculations-list__values')
    :param main:
    :return:
    """
    percent = main.find_elements(By.CSS_SELECTOR, '.calculations-list__values')[2].text
    return percent


def psb_write(driver, active_sheet, row_f):
    """
    В первых 2 ссылках есть калькуляторы, они парсятся, потом забираются таблицы. Из третьей же забирается только
    таблица. Забирает периоды из второго бара в калькуляторе - там есть аттрибут 'data-custom-values',
    если с ним сделать пару манипуляций, то можно получить сроки. Соответственно в калькуляторе максимальное значение -
    7млн. - важно понимать.
    :param driver: browser
    :param active_sheet: Эксель лист, куда записывать
    :param row_f: ряд, начиная с которого надо записывать
    :return:
    """
    hrefs = ['https://www.psbank.ru/Personal/Saving/Digital',
             'https://www.psbank.ru/Personal/Saving/MyProfit',
             'https://www.psbank.ru/Personal/Saving/CourseForProfit']
    amounts = ['500000', '1000000', '2000000', '3000000',
               '5000000', '7000000']
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    for k in range(len(hrefs)):
        driver.get(hrefs[k])
        print(f"    proccessing {k + 1}/4...")
        main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.CLASS_NAME, "wrapper"))
        )
        scroll(driver)
        time.sleep(1)
        if k in [0, 1]:
            periods = main.find_elements(By.CSS_SELECTOR, '.range-label')[1].get_attribute('data-custom-values')
            periods = re.sub("[\[\]\']", '', periods)
            periods = [period + ' мес.' for period in periods.replace(' ', '').split(',')]
            name = main.find_element(By.TAG_NAME, 'h1').text
            use_slider(driver, main, -400)
            slide = [0 if i == 0 else 100 for i in range(len(periods))]
            active_sheet[f'A{row_f}'] = name
            for i in range(len(periods)):
                active_sheet[f'{chr(66 + i)}{row_f}'] = periods[i]
            row_f += 1
            for i in range(len(amounts)):
                active_sheet[f'A{row_f}'] = amounts[i]
                use_bar(main, amounts[i])
                for j in range(len(periods)):
                    use_slider(driver, main, slide[j])
                    active_sheet[f'{chr(66 + j)}{row_f}'] = parse_psb(main)
                row_f += 1
                use_slider(driver, main, -500)
            row_f += 1
            scroll(driver)
        else:
            row_f = 27
            name = main.find_element(By.TAG_NAME, 'h1').text.split('\n')[0]
            active_sheet[f'L{row_f - 8}'] = name

        time.sleep(3)
        table = main.find_element(By.TAG_NAME, 'table')
        if table.text:
            write_table(table, active_sheet, row_f)


if __name__ == '__main__':
    wb = openpyxl.open('testing.xlsx')
    sheet = wb.worksheets[9]

    warnings.filterwarnings("ignore")
    PATH = "C:\\Program Files (x86)\\chromedriver.exe"

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--incognito')
    options.add_argument('window-size=1920,1080')
    options.add_argument('--ignore-certificate-errors')
    # options.add_experimental_option("excludeSwitches", ["enable-automation"])
    # options.add_experimental_option('useAutomationExtension', False)
    browser = webdriver.Chrome(executable_path=PATH, options=options)  # executable_path=PATH, chrome_options=options

    row = 1
    psb_write(browser, sheet, row)
    try:
        print(F"    proccessing 4/4...")
        write_strong(browser, sheet, 25)
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print(e)
    finally:
        wb.save('testing.xlsx')
        browser.quit()
