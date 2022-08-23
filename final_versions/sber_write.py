import time
import sys
import os
import openpyxl
import warnings
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common import action_chains
from selenium.webdriver.remote.webelement import WebElement


def parse_curr(driver: webdriver, active_sheet, row_f):
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    amounts = ["1000", "25000", "50000", "100000", "300000"]
    periods = ["3 мес", "6 мес", "9 мес", "12 мес", "18 мес", "2 года", "3 года"]
    slide_by = [0, 50, 30, 30, 60, 60, 120]
    char = ['₽', '$', '€']
    driver.get('https://www.sberbank.ru/ru/person/contributions/depositsnew')
    time.sleep(3)
    driver.execute_script(f"window.scrollBy(0, {200})")
    time.sleep(1)
    main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
        EC.presence_of_element_located((By.ID, "main-page"))
    )
    time.sleep(1)
    button_calc = main.find_element(By.CSS_SELECTOR, '.dc-menu__calc-wrapper')
    button_calc.click()
    bar = main.find_element(By.CSS_SELECTOR, '.dk-sbol-input.dk-sbol-input_size_md.dk-sbol-input-with-radio__input')
    buttons = main.find_elements(By.CSS_SELECTOR, '.dk-sbol-segmented__control.dk-sbol-segmented__control_size_md')
    slider = main.find_element(By.CSS_SELECTOR, '.dk-sbol-slider-track__handle')
    time.sleep(1)
    webdriver.common.action_chains.ActionChains(driver).drag_and_drop_by_offset(slider, -300, 0).perform()
    find = main.find_element(By.CSS_SELECTOR, '.dk-sbol-button__content')
    for k in [1, 2]:
        buttons[k].click()
        for i in range(len(periods)):
            active_sheet[f'{chr(66 + i)}{row_f}'] = periods[i]
        row_f += 1
        for i in range(len(amounts)):
            active_sheet[f'A{row_f}'] = amounts[i] + char[k]
            bar.send_keys([Keys.BACKSPACE for _ in range(9)], amounts[i])
            for j in range(len(periods)):
                webdriver.common.action_chains.ActionChains(driver).drag_and_drop_by_offset(slider, slide_by[j], 0).perform()
                find.click()
                active_sheet[f'{chr(66 + j)}{row_f}'] = main.find_element(By.CSS_SELECTOR, '.dk-sbol-heading.dk-sbol-heading_size_sm').text
            row_f += 1
            webdriver.common.action_chains.ActionChains(driver).drag_and_drop_by_offset(slider, -400, 0).perform()
        row_f += 1

def parse_in_advance(active_sheet, row_f):
    for i in range(8):
        for j in range(6):
            active_sheet[f'{chr(65 + j)}{row_f}'] = active_sheet[f'{chr(65 + j)}{row_f - 9}'].value
        row_f += 1
    # active_sheet[f'A{row_f - 8}'] = 'Пополняемый'
    return row_f


def navigate_upravl(main: WebElement, mode=1):
    box = main.find_elements(By.CSS_SELECTOR, '.dk-sbol-value-select__combobox.dk-sbol-value-select__combobox_size_md')
    box[1].click()
    time.sleep(1)
    if mode == 1:
        box[1].send_keys(Keys.ARROW_DOWN)
        box[1].send_keys(Keys.RETURN)
    elif mode == 0:
        while True:
            ammounts = main.find_elements(By.CSS_SELECTOR, '.dk-sbol-sellist__li.dk-sbol-sellist__li_size_md')
            sums = [ammount.text for ammount in ammounts if ammount.text != '']
            if sums:
                box[1].send_keys(Keys.RETURN)
                return sums
            time.sleep(1)


def find_buttons(main: WebElement):
    while True:
        buttons = main.find_elements(By.CSS_SELECTOR, '.dc-select-tab-input__button')
        if buttons:
            return [button for button in buttons if button.text != '']
        time.sleep(1)


def parse_upravlyai(driver: webdriver, active_sheet, row_f):
    driver.get('https://www.sberbank.ru/ru/person/contributions/promo-upravlyai')
    time.sleep(5)
    driver.execute_script(f"window.scrollBy(0, {1000})")
    time.sleep(1)
    main = driver.find_element(By.ID, 'main-page')
    bar = main.find_element(By.CSS_SELECTOR, '.dk-sbol-input.dk-sbol-input_size_md')
    bar.click()
    bar.send_keys([Keys.BACKSPACE for _ in range(9)], '1000', Keys.RETURN)
    amounts = navigate_upravl(main, 0)
    buttons = find_buttons(main)
    active_sheet[f'A{row_f}'] = 'Управляй+'
    for i in range(len(buttons)):
        active_sheet[f'{chr(66 + i)}{row_f}'] = buttons[i].text
    row_f += 1
    active_sheet[f'A{row_f}'] = amounts[0]
    for i in range(len(buttons)):
        buttons[i].click()
        active_sheet[F"{chr(66 + i)}{row_f}"] = parse_sber(main, None, 0, 0, 0) + '%'
    row_f += 1

    for i in range(1, len(amounts)):
        active_sheet[f'A{row_f}'] = amounts[i]
        navigate_upravl(main)
        for j in range(len(buttons)):
            buttons[j].click()
            active_sheet[f'{chr(66 + j)}{row_f}'] = parse_sber(main, None, 0, 0, 0) + '%'
        row_f += 1
    headers = main.find_elements(By.CSS_SELECTOR, '.ar-table__th')
    rows = main.find_elements(By.CSS_SELECTOR, '.ar-table__td')
    for i in range(len(headers)):
        active_sheet[f'{chr(74 + i)}{row_f - 5}'] = headers[i].text
    a = len(rows) // len(headers)
    for i in range(a):
        for j in range(len(headers)):
            if j == 0:
                active_sheet[f'{chr(74)}{row_f - 4 + i}'] = rows[len(headers) * i + j].text.split('\n')[0]
            else:
                active_sheet[f'{chr(74 + j)}{row_f - 4 + i}'] = rows[len(headers) * i + j].text.split()[1] + '%'


def count_sberprime(main: WebElement):
    percent_1 = parse_sber(main, None, 0, 0, 0)
    buttons = main.find_elements(By.CSS_SELECTOR, '.dk-sbol-switch__control.dk-sbol-switch__control_mode_off.dk-sbol-switch__control_size_lg')
    for button in buttons:
        button.click()
    time.sleep(1)
    percent_2 = parse_sber(main, None, 0, 0, 0)
    diff = str(float(percent_2.strip('%').replace(',', '.')) - float(percent_1.strip('%').replace(',', '.'))) + '%'
    return diff.replace('.', ',')


def get_from_list(main: WebElement, element: int, mode=0):
    while True:
        try:
            class_1 = '.dk-sbol-value-select__combobox.dk-sbol-value-select__combobox_size_md.dk-sbol-field__control.dk-sbol-field__control_size_md-combobox'
            selector = main.find_element(By.CSS_SELECTOR, class_1)
            selector.click()
            time.sleep(1)
            buttons = main.find_elements(By.CSS_SELECTOR, '.dk-sbol-sellist__li.dk-sbol-sellist__li_size_md')
            if mode == 1:
                return [button.text for button in buttons if button.text != '']
            elif mode == 0:
                buttons[element].click()
            break
        except Exception as g:
            print(g)
            time.sleep(1)


def parse_sber(main: WebElement, active_sheet, row_f, col_f, mode=1):
    percents = main.find_elements(By.CSS_SELECTOR, '.dk-sbol-an-number dk-sbol-an-number_size_md'.replace(' ', '.'))
    if mode == 1:
        active_sheet[f'{chr(col_f)}{row_f}'] = percents[1].text
    else:
        return percents[1].text


def use_bar(main, request):
    bar = main.find_element(By.CSS_SELECTOR, '.dk-sbol-input.dk-sbol-input_size_md.dk-sbol-field__control.dk-sbol-field__control_size_md')
    bar.click()
    bar.send_keys([Keys.BACKSPACE for _ in range(9)])
    bar.send_keys(request)


def switch_prime(main):
    class_1 = '.dk-sbol-switch__control.dk-sbol-switch__control_mode_on.dk-sbol-switch__control_size_lg'
    sber_prime_buttons = main.find_elements(By.CSS_SELECTOR, class_1)
    for button in sber_prime_buttons:
        button.click()


def parse_sbervklad(driver: webdriver, active_sheet, row_f):
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    driver.get('https://www.sberbank.ru/ru/person/contributions/deposits/sbervklad_a')
    time.sleep(10)
    driver.execute_script(f"window.scrollBy(0, {1800})")
    time.sleep(1)
    main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
        EC.presence_of_element_located((By.ID, "main-page"))
    )
    switch_prime(main)
    driver.execute_script(f"window.scrollBy(0, {-200})")
    time.sleep(1)
    buttons = get_from_list(main, 1, 1)  # находим все периоды
    active_sheet[f'A{row_f}'] = 'СберВклад'
    for i in range(len(buttons)):
        active_sheet[f'{chr(66 + i)}{row_f}'] = buttons[i]
    row_f += 1

    ammounts = ['500000', '1000000', '2000000', '3000000',
                '5000000', '10000000', '20000000']
    for i in range(len(ammounts)):
        active_sheet[f'A{row_f}'] = ammounts[i]
        use_bar(main, ammounts[i])
        for j in range(len(buttons)):
            get_from_list(main, j)
            parse_sber(main, active_sheet, row_f, 66 + j)
        row_f += 1

    headers = main.find_elements(By.TAG_NAME, 'th')
    rows = main.find_elements(By.TAG_NAME, 'td')
    for i in range(len(headers)):
        active_sheet[f'{chr(74 + i)}{row_f - 7}'] = headers[i].text

    a = len(rows) // len(headers)
    for i in range(a):
        for j in range(len(headers)):
            active_sheet[f'{chr(74 + j)}{row_f - 6 + i}'] = rows[i * len(headers) + j].text
    driver.execute_script(f"window.scrollBy(0, {200})")
    time.sleep(1)
    active_sheet[f'H{row_f - 1}'] = 'Надбавка за Сберпрайм'
    active_sheet[f'I{row_f - 1}'] = count_sberprime(main)

    return row_f


if __name__ == '__main__':
    wb = openpyxl.open(date.today().strftime("%d.%m.%y") + '.xlsx')
    sheet = wb.worksheets[6]
    sheet_1 = wb.worksheets[7]
    warnings.filterwarnings("ignore")
    PATH = "C:\\Program Files (x86)\\chromedriver.exe"

    options = webdriver.ChromeOptions()
    # options.add_argument('--headless')
    options.add_argument('--incognito')
    options.add_argument('window-size=1920,1080')

    options.add_argument('--ignore-certificate-errors')
    # options.add_experimental_option("excludeSwitches", ["enable-automation"])
    # options.add_experimental_option('useAutomationExtension', False)
    browser = webdriver.Chrome(executable_path=PATH, options=options)  # executable_path=PATH, chrome_options=options

    row = 1

    try:
        print('    processing 1/4...')
        row = parse_sbervklad(browser, sheet, row)
        print('    processing 2/4...')
        row = parse_in_advance(sheet, row + 1)
        print('    processing 3/4...')
        parse_upravlyai(browser, sheet, row + 1)
        print('    proccessing 4/4...')
        parse_curr(browser, sheet_1, 1)
    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        print(exc_type, fname, exc_tb.tb_lineno)
        print(e)
    finally:
        wb.save(date.today().strftime("%d.%m.%y") + '.xlsx')
        browser.quit()
