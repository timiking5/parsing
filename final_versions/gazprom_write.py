import time
import warnings
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common import action_chains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC


def sum_proc(proc, bonus):
    """
    Функция для расчёта процента премиальных процентов
    6,77% + 0,3% = 7,07%
    :param proc:
    :param bonus:
    :return:
    """
    return str(round(float(proc.strip('%').replace(',', '.')) + float(bonus.strip('%').replace(',', '.')), 3)).replace('.', ',') + '%'


def take_tables(driver, active_sheet, row_f):
    """
    Заходит на сайт, нажимает на вкладку проценты, забирает таблицу, где проценты с капитализацией.
    Сначала забирает верхнюю строчку таблицы, а потом остальные, записывает
    :param driver: browser
    :param active_sheet: Лист excel, куда записывать
    :param row_f: строчка, с которой записывать
    :return:
    """
    hrefs = ['https://www.gazprombank.ru/personal/increase/deposits/detail/6049',
             'https://www.gazprombank.ru/personal/increase/deposits/detail/2491',
             'https://www.gazprombank.ru/personal/increase/deposits/detail/1929']

    depos_name = ["\"Ваш успех\" с капитализ.",
                  "\"Копить\" с капитализ.",
                  "\"Управлять\" c капитализ."]
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    for _ in [0, 1, 2]:
        print(f"    processing {_ + 3}/5...")
        try:
            driver.get(hrefs[_])
            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.CLASS_NAME, "nr-layout"))
            )
            time.sleep(2)
            browser.execute_script(f"window.scrollBy(0, {900})")
            time.sleep(1)
            buttons = main.find_elements(By.CSS_SELECTOR, '.nr-tabs__el--inner')
            for button in buttons:
                if "Процентные ставки" in button.text:
                    press_button(driver, button)
                    break
            time.sleep(1)
            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.CLASS_NAME, "nr-layout"))
            )
            search = main.find_elements(By.TAG_NAME, 'h4')
            i = 0
            count = 0  # на всякий
            for elem in search:
                if 'капитал' in elem.text:
                    count = i
                i += 1
            table = main.find_elements(By.TAG_NAME, 'table')[count]
            headers = table.find_elements(By.TAG_NAME, 'th')
            words = table.find_elements(By.TAG_NAME, 'td')
            col = 65
            active_sheet[f'A{row_f}'] = depos_name[_]
            for i in range(1, len(headers)):
                active_sheet[f'{chr(col + i)}{row_f}'] = headers[i].text
            row_f += 1
            a = len(words) // (len(headers))

            for i in range(a):
                for j in range(len(headers)):
                    active_sheet[f'{chr(col + j)}{row_f}'] = words[i * (len(headers)) + j].text
                row_f += 1
            row_f += 1
        except Exception as e:
            print(e, 'failed', sep='\n')


def press_button(driver_, button):
    """
    Нажимает кнопку с помощью action_chains
    :param driver_: browser
    :param button: web element
    :return:
    """
    # print(button.text)
    action = webdriver.common.action_chains.ActionChains(driver_)
    action.move_to_element_with_offset(button, 2, 2)
    action.click()
    action.perform()


def parsing_gaz(driver_, row_, col_, active_sheet, filter_, query):
    """
    Забирает процент и вставляет в нужную ячейку в листе эксель.
    :param driver_: browser
    :param row_: строка
    :param col_: столбец
    :param active_sheet: Лист эксель, куда записывать
    :param filter_: By. ... . Как искать элемент
    :param query: "поисковой запрос" для элемента
    :return:
    """
    # ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    while True:
        main_ = driver_.find_element(By.CLASS_NAME, "nr-layout")

        search = main_.find_element(filter_, query)
        if search:
            active_sheet[f'{chr(col_)}{row_}'] = search.text
            break
        time.sleep(1)


def use_bar(request, times, main_):
    """
    Вводит нужное значение в поле ввода
    :param request: Что надо вставить
    :param times: По факту ненужный аргумент, можно просто поставить вместо него 9 или 10
    :param main_: Чтобы не прогружать main в этой функции, его можно передать из предыдущей
    :return:
    """
    bar = main_.find_element(By.CSS_SELECTOR, '.nr-form-mask.input_root__input--98347')
    if bar:
        bar.send_keys([Keys.BACK_SPACE for _ in range(times)], request)


def gazprom_write(driver: webdriver, active_sheet, row_f):
    """
    Переходит по ссылке, опускается до калькулятора (чтобы можно было нажимать кнопки), ищет все кнопки, потом вводит
    сумму и проходится по срокам (кнопкам), вводит новую сумму и т.д. Потом переходит в "Условия по вкладу", забирает
    оттуда надбавку за премиум у пересчитывает проценты для премиума.
    :param driver:
    :param active_sheet:
    :param row_f:
    :return:
    """
    hrefs = ['https://www.gazprombank.ru/personal/increase/deposits/detail/2491',
             'https://www.gazprombank.ru/personal/increase/deposits/detail/1929']
    ammounts = ['500000', '1000000', '2000000', '3000000',
                '5000000', '10000000', '20000000']
    scroll = [1200, 1400]
    times = [7, 6, 7, 7, 7, 7, 8]
    for b in [0, 1]:
        print(f"    processing {b + 1}/5...")
        browser.get(hrefs[b])
        time.sleep(2)
        browser.execute_script(f"window.scrollBy(0, {scroll[b]})")
        time.sleep(1)
        col = 66
        try:
            ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
            while True:
                main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "nr-layout"))
                )  # .chips_tag_group_root__tag--c5c08
                buttons = main.find_elements(By.CSS_SELECTOR, '.chips_tag_root--3b499')
                if buttons:
                    break
                time.sleep(1)
            for i in range(len(buttons)):
                active_sheet[f'{chr(col + i)}{row_f - 1}'] = buttons[i].text
            for j in range(7):
                active_sheet[f'A{row_f}'] = ammounts[j]
                use_bar(ammounts[j], times[j], main)
                for i in range(len(buttons)):
                    press_button(driver, buttons[i])
                    query = '.calculator_results_row_root__value--e5028.calculator_results_row_rate--e5028'
                    parsing_gaz(driver, row_f, col + i, active_sheet, By.CSS_SELECTOR, query)
                row_f += 1

            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "nr-layout"))
                )
            buttons_2 = main.find_elements(By.CSS_SELECTOR, '.nr-tabs__el--inner')
            for button in buttons_2:
                if 'Условия по вкладу' in button.text:
                    press_button(driver, button)
                    break
            time.sleep(2)
            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.CLASS_NAME, "nr-layout"))
            )
            search = main.find_elements(By.TAG_NAME, 'p')
            prem = "0,0%"
            for elem in search:
                if 'Премиальным' in elem.text:
                    txt = elem.text.split()
                    prem = txt[len(txt) - 1]
                    active_sheet['J29'] = elem.text
                    
            for j in range(7):
                for i in range(len(buttons)):
                    active_sheet[f'{chr(col + i + 9)}{row_f - j - 1}'] = sum_proc(active_sheet[f'{chr(col + i)}{row_f - j - 1}'].value, prem)
        except Exception as e:
            print(e)

        row_f += 1
    return row_f


if __name__ == '__main__':
    PATH = "C:\\Program Files (x86)\\geckodriver.exe"
    warnings.filterwarnings("ignore")

    options = webdriver.FirefoxOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--incognito')
    options.add_argument('window-size=1920,1080')
    options.add_argument('--headless')

    row = 3
    wb = openpyxl.open('testing.xlsx')
    sheet = wb.worksheets[3]

    browser = webdriver.Firefox(executable_path=PATH, options=options)

    row = gazprom_write(browser, sheet, row)
    take_tables(browser, sheet, row)
    browser.quit()
    wb.save('testing.xlsx')
