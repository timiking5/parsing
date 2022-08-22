import time
import openpyxl
# import sys, os
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common import action_chains


def max_perc(percents):
    mx = percents[0].text
    for k in range(len(percents)):
        if count_value(percents[k].text) > count_value(mx):
            mx = percents[k].text
    return mx


def count_value(percent):
    """
    Если вдруг придётся сравнивать проценты между собой
    6,667% --> 6.667
    """
    return float(percent.strip('%').replace(',', '.'))


def vtb_write(driver: webdriver, active_sheet, row_f, hrefs):
    """
    Проходит по всем ссылкам из href, и если название вклада не "Вклад в будущее" или "Большие возможности"
    (как раз в этих двух вкладах по процентам будет большая таблица), то выводит срок вклада и его максимальный процент.
    Пришлось поставить While True: try: ... . Потому что элементы на странице, бывает, не успевают загружаться
    и выводит ошибку (обычно элемент будет просто пустой). Остальное вроде понятно, хотя ВТБ иногда меняет названия
    классов (CSS_SELECTOR). Возможно, мне просто показалось, но такое вроде как уже случалось.
    """
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    class_1 = '.typographystyles__Box-foundation-kit__sc-14qzghz-0.hVDbVT.info-line-itemstyles__Title-info-line-item__sc-gswh8c-1.egXuTZ'
    class_2 = '.typographystyles__Box-foundation-kit__sc-14qzghz-0.kOPQGR.hero-blockstyles__Heading1Styled-hero-block__sc-124m6ob-3.klzwby'
    class_3 = '.typographystyles__Box-foundation-kit__sc-14qzghz-0.jEFSaq.numbersstyles__TypographyTitle-foundation-kit__sc-1xhbrzd-4.haHdlc'
    class_4 = '.typographystyles__Box-foundation-kit__sc-14qzghz-0.gGALTE.markdown-headingstyles__HeadingTypography-foundation-kit__sc-7uz79g-0.frxXnP'
    class_5 = '.typographystyles__Box-foundation-kit__sc-14qzghz-0.bIbiHl.table-cellstyles__HeadingTypography-table-cell__sc-pq68xl-2.eMAckD'
    # Да, можно сделать в виде списка, но разница небольшая
    for href in hrefs:
        driver.get(href)
        while True:
            try:
                main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                    EC.presence_of_element_located((By.ID, "root"))
                )
                account_name = main.find_element(By.CSS_SELECTOR, class_2)
                if account_name:

                    percents = main.find_elements(By.CSS_SELECTOR, class_1)
                    if 'Вклад «Большие возможности»' in account_name.text or 'Вклад в будущее' in account_name.text:
                        row_f += 1
                        active_sheet[f'A{row_f}'] = account_name.text
                        row_f += 1
                        percents = main.find_elements(By.CSS_SELECTOR, class_3)
                        rows = main.find_elements(By.CSS_SELECTOR, class_4)
                        headers = main.find_elements(By.CSS_SELECTOR, class_5)
                        for i in range(len(headers)):
                            active_sheet[f'{chr(65 + i)}{row_f}'] = headers[i].text
                        row_f += 1
                        for i in range(len(rows)):
                            active_sheet[f'A{row_f}'] = rows[i].text
                            for j in range(len(headers) - 1):
                                active_sheet[f'{chr(66 + j)}{row_f}'] = percents[(len(headers) - 1) * i + j].text
                                # Обращаемся с одномерным списком, как с двумерным
                                # При том "ширина" таблицы процентов на 1 меньше длины заголовков (см. сайт)
                            row_f += 1
                    elif "Управляемый" in account_name.text:
                        row_f += 1
                        active_sheet[f'A{row_f}'] = account_name.text
                        active_sheet[f'B{row_f}'] = "Если снимать"
                        active_sheet[f'C{row_f}'] = "Если оставлять"
                        row_f += 1
                        headers = main.find_elements(By.CSS_SELECTOR, class_5)
                        rows = main.find_elements(By.CSS_SELECTOR, class_4)
                        percents = main.find_elements(By.CSS_SELECTOR, class_3)
                        percents = [percent for percent in percents if percent.text != '']

                        rows = list(set([row_.text for row_ in rows if row_.text != '']))[::-1]
                        for i in range(2):
                            active_sheet[f'{chr(65 + i)}{row_f}'] = headers[i].text
                        row_f += 1
                        for i in range(len(rows)):
                            active_sheet[f'A{row_f + i}'] = rows[i]
                        for i in range(len(rows)):
                            for j in range(2):
                                active_sheet[f'{chr(66 + j)}{row_f}'] = percents[2 * i + j].text
                            row_f += 1
                    elif "Наличный" in account_name.text:
                        row_f += 1
                        active_sheet[f'A{row_f}'] = account_name.text
                        active_sheet[f'B{row_f}'] = 'Процент'
                        active_sheet[f'C{row_f}'] = 'Срок'
                        row_f += 1
                        for i in range(2):
                            if i == 1:
                                main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                                    EC.presence_of_element_located((By.ID, "root"))
                                )
                            nums = main.find_elements(By.CSS_SELECTOR, class_4)
                            num = nums[0]
                            for number in nums:
                                if '0' in number.text:
                                    num = number
                            active_sheet[f'A{row_f}'] = num.text
                            percents = main.find_elements(By.CSS_SELECTOR, class_3)
                            percents = [percent for percent in percents if percent.text != '']
                            active_sheet[f'B{row_f}'] = max_perc(percents)
                            periods = main.find_elements(By.CSS_SELECTOR, class_5)
                            for period in periods:
                                if 'день' in period.text or 'дней' in period.text:
                                    active_sheet[f'C{row_f}'] = period.text
                            if i == 0:
                                browser.execute_script(f"window.scrollBy(0, {750})")
                                time.sleep(1)
                                buttons = main.find_elements(By.CSS_SELECTOR, '.tabs-headerstyles__TabTitleHorizontal-foundation-kit__sc-1w1sfys-2.fNQQZu')
                                for button in buttons:
                                    if 'Доходность' in button.text:
                                        action = webdriver.common.action_chains.ActionChains(driver)
                                        action.move_to_element_with_offset(button, 2, 2)
                                        action.click()
                                        action.perform()
                                        time.sleep(1)
                            row_f += 1

                    elif percents:
                        active_sheet[f'A{row_f}'] = account_name.text
                        active_sheet[f'B{row_f}'] = 'Процент'
                        active_sheet[f'C{row_f}'] = 'Срок'
                        row_f += 1
                        for elem in percents:
                            if '%' in elem.text:
                                active_sheet[f'B{row_f}'] = elem.text.split()[1]
                            elif 'дней' in elem.text or 'день' in elem.text:
                                active_sheet[f'C{row_f}'] = elem.text

                    row_f += 1
                    break

            except Exception:
                # exc_type, exc_obj, exc_tb = sys.exc_info()
                # fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
                # print(exc_type, fname, exc_tb.tb_lineno)
                time.sleep(1)


def get_hrefs_vtb(driver: webdriver):
    """
    Собственно говоря находит все ссылки из раздела "вклады и счета", которые введут на вклады.
    Сначала смотрит под какими индексами названия имеют слово "вклад", а потом из всех ссылок "Подробные условия"
    выбирает те, что стоят под индексами, под которыми стоят названия, содержащие слово "вклад"
    1) [Счёт "Копилка", Вклад "Первый", Вклад "Второй", Счёт "Третий"] --> [1, 2]
    2) [Веб элементы "Подробные условия"]
    3) Из списка 2 берём элементы под индексами [1, 2]
    :param driver: browser
    """
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    driver.get('https://www.vtb.ru/personal/vklady-i-scheta/')
    time.sleep(3)  # чтобы не выдавало ошибку. Элементы не успевают прогружаться
    browser.execute_script(f"window.scrollBy(0, {1500})")
    time.sleep(1)
    main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
        EC.presence_of_element_located((By.ID, "root"))
    )
    hrefs_index = []
    hrefs = []
    buttons = main.find_elements(By.CSS_SELECTOR, '.buttonstyles__Content-foundation-kit__sc-sa2uer-0.bLTIdu')

    for button in buttons:
        if 'Показать еще' in button.text:
            action = webdriver.common.action_chains.ActionChains(driver)
            action.move_to_element_with_offset(button, 2, 2)
            action.click()
            action.perform()
    """Чтобы увидеть все вклады, нужно нажать "Показать ещё"."""
    while True:
        main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.ID, "root"))
        )
        depos_names = main.find_elements(By.CSS_SELECTOR, '.typographystyles__Box-foundation-kit__sc-14qzghz-0.gGALTE')
        if depos_names:
            for i in range(0, len(depos_names)):
                if 'вклад' in depos_names[i].text.lower() and '«' in depos_names[i].text:
                    hrefs_index.append(i)
            depos_hrefs = main.find_elements(By.LINK_TEXT, 'Подробные условия')
            for i in hrefs_index:
                hrefs.append(depos_hrefs[i].get_attribute("href"))
            break
    return hrefs


if __name__ == '__main__':
    PATH = "C:\\Program Files (x86)\\chromedriver.exe"

    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--incognito')
    options.add_argument('window-size=1920,1080')
    # options.add_argument('--headless')
    """Классическое предупреждение для будущих разработчиков: вроде как в селениуме собираются убрать параметр
    executable_path (см. DeprecationWarning при запуске), но сейчас же у меня, как бы я ни пытался, не получается
    использовать service=... . Возможно, в будущем это уже будет работать, так что вам придётся с этим запаритсья.
    Вместо chrome_options= можно использовать options= ... . Аргумент headless можно использовать по желанию, но
    иногда он приводит к тому, что программа не работает."""
    # service=Service(ChromeDriverManager().install())
    browser = webdriver.Chrome(executable_path=PATH, chrome_options=options)

    wb = openpyxl.open('testing.xlsx')
    sheet = wb.worksheets[1]
    """В строчке 146 в квадратных скобках указывается номер листа эксель файла.
    Название можно поменять, но и тогда название файла поменяйте."""
    row = 1
    vtb_write(browser, sheet, row, get_hrefs_vtb(browser))
    browser.quit()
    wb.save('testing.xlsx')
