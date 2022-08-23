import time
import openpyxl
import warnings
from selenium import webdriver
from selenium.webdriver.common.by import By
# from selenium.common.exceptions import NoSuchElementException
# from selenium.common.exceptions import StaleElementReferenceException
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement


def write_prem(driver: webdriver, active_sheet, row_f):
    """
    Тут просто надо забрать таблицы
    :param driver: browser
    :param active_sheet: Эксель лист, куда записывать
    :param row_f: Ряд, с которого надо записывать
    :return:
    """
    hrefs = ['https://www.rshb.ru/natural/premium/ultra/deposit-dohodniy/',
             'https://www.rshb.ru/natural/premium/ultra/deposit-popolnaemiy/',
             'https://www.rshb.ru/natural/premium/ultra/deposit-komfortniy/']
    for k in range(len(hrefs)):
        driver.get(hrefs[k])
        print(f"    proccessing {k + 4}/6...")
        driver.execute_script(f"window.scrollBy(0, {1000})")
        time.sleep(1)
        main = driver.find_element(By.CLASS_NAME, 'rshb-layout')
        buttons = main.find_elements(By.CLASS_NAME, 'accordion__header')
        for button in buttons:
            if "капитализ" in button.text.lower():
                button.click()
                break

        rows = main.find_elements(By.TAG_NAME, 'table')[1].find_elements(By.TAG_NAME, 'tr')
        active_sheet[f'A{row_f}'] = main.find_element(By.TAG_NAME, 'h1').text
        row_f += 1
        for i in range(len(rows)):
            tds = rows[i].find_elements(By.TAG_NAME, 'td')
            for j in range(len(tds)):
                active_sheet[f'{chr(65 + j)}{row_f}'] = tds[j].text
            row_f += 1
        row_f += 1


def use_bar(main: WebElement, request):
    """
    Использует вводное поле, передаёт туда значение request
    :param main:
    :param request: то, что надо передать
    :return:
    """
    bar = main.find_element(By.CSS_SELECTOR, '.input-field-input')
    bar.send_keys([Keys.BACKSPACE for _ in range(9)], request)


def parse_rshb(main: WebElement):
    """
    Возвращает значение процента
    :param main:
    :return:
    """

    return main.find_element(By.CSS_SELECTOR, '.moe-calc-result-value').text


def write_rshb(driver: webdriver, active_sheet, row_f):
    """
    Парсит калькулятор, перебирает, месяцы, для каждой суммы. Сначала находит кнопки, потом начинает записывать.
    :param driver: Browser
    :param active_sheet: Эксель лист
    :param row_f: Ряд, с которого надо начинать записывать
    :return:
    """
    hrefs = ['https://www.rshb.ru/natural/deposits/dohodniy/']
    amounts = ['500000', '1000000', '2000000', '3000000',
               '5000000', '10000000', '20000000']
    for href in hrefs:
        print(f"    proccessing 3/6...")
        driver.get(href)
        browser.execute_script(f"window.scrollBy(0, {600})")
        time.sleep(1)
        main = driver.find_element(By.CSS_SELECTOR, '.moe')
        buttons = [i for i in main.find_elements(By.CSS_SELECTOR, '.toggle-buttons-button') if 'мес.' in i.text]
        name = main.find_elements(By.TAG_NAME, 'h1')[1].text.split()[1]
        active_sheet[f'A{row_f}'] = name
        for i in range(len(buttons)):
            active_sheet[f'{chr(66 + i)}{row_f}'] = buttons[i].text.split('\n')[0]
        row_f += 1
        for i in range(len(amounts)):
            active_sheet[f'A{row_f}'] = amounts[i]
            use_bar(main, amounts[i])
            for j in range(len(buttons)):
                buttons[j].click()
                active_sheet[f'{chr(66 + j)}{row_f}'] = parse_rshb(main)
            row_f += 1


def write_rshb_tables(driver: webdriver, active_sheet, row_f):
    """
    С некоторых сайтов можно забрать только таблицы.
    :param driver: browser
    :param active_sheet: Эксель лист, куда записывать
    :param row_f: Ряд, с которого надо начинать записывать
    :return:
    """
    hrefs = ['https://www.rshb.ru/natural/deposits/popolnaemiy/',
             'https://www.rshb.ru/natural/deposits/komfortniy/']
    for k in range(len(hrefs)):
        driver.get(hrefs[k])
        print(f"    proccessing {k + 1}/6...")
        main = driver.find_element(By.CSS_SELECTOR, '.page')
        name = main.find_element(By.TAG_NAME, 'h1')
        table = main.find_element(By.TAG_NAME, 'table')
        trs = table.find_elements(By.TAG_NAME, 'tr')  # Ряды таблицы - table rows

        active_sheet[f'A{row_f}'] = name.text
        row_f += 1
        for i in range(len(trs)):
            tds = trs[i].find_elements(By.TAG_NAME, 'td')
            if i == 0:
                active_sheet[f'A{row_f}'] = tds[0].text
                continue
            elif i == 1:
                for j in range(len(tds)):
                    active_sheet[f'{chr(66 + j)}{row_f}'] = tds[j].text
            else:
                for j in range(len(tds)):
                    active_sheet[f'{chr(65 + j)}{row_f}'] = tds[j].text
            row_f += 1
        row_f += 1
    return row_f


if __name__ == '__main__':
    wb = openpyxl.open('testing.xlsx')
    sheet = wb.worksheets[7]
    sheet_1 = wb.worksheets[8]

    warnings.filterwarnings("ignore")
    PATH = "C:\\Program Files (x86)\\chromedriver.exe"

    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    options.add_argument('--incognito')
    options.add_argument('window-size=1920,1080')
    options.add_argument('--ignore-certificate-errors')
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    browser = webdriver.Chrome(executable_path=PATH, options=options)  # executable_path=PATH, chrome_options=options

    row = 1
    row = write_rshb_tables(browser, sheet, row)
    write_rshb(browser, sheet, row)
    write_prem(browser, sheet_1, 1)
    wb.save('testing.xlsx')
    browser.quit()
