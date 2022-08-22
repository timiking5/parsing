import time
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
    return str(round(float(proc.strip('%').replace(',', '.')) + float(bonus.strip('%').replace(',', '.')), 3)).replace('.', ',') + '%'


def take_tables(driver, active_sheet, row_f):
    hrefs = ['https://www.gazprombank.ru/personal/increase/deposits/detail/6049',
             'https://www.gazprombank.ru/personal/increase/deposits/detail/2491',
             'https://www.gazprombank.ru/personal/increase/deposits/detail/1929']

    depos_name = ["\"Ваш успех\" с капитализ.",
                  "\"Копить\" с капитализ.",
                  "\"Управлять\" c капитализ."]
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    for _ in [0, 1, 2]:
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
    # print(button.text)
    action = webdriver.common.action_chains.ActionChains(driver_)
    action.move_to_element_with_offset(button, 2, 2)
    action.click()
    action.perform()


def parsing_gaz(driver_, row_, col_, active_sheet, filter_, query):
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    while True:
        time.sleep(1)
        main_ = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.CLASS_NAME, "nr-layout"))
        )
        search = main_.find_element(filter_, query)
        if search:
            # print(f'{chr(col_)}{row_}', search.text)
            active_sheet[f'{chr(col_)}{row_}'] = search.text
            break


def use_bar(request, times, main_):
    # ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    # main_ = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
    #         EC.presence_of_element_located((By.CLASS_NAME, "nr-layout"))
    #     )
    bar = main_.find_element(By.CSS_SELECTOR, '.nr-form-mask.input_root__input--98347')
    if bar:
        bar.send_keys([Keys.BACK_SPACE for _ in range(times)], request)


def gazprom_write(driver: webdriver, active_sheet, row_f):
    hrefs = ['https://www.gazprombank.ru/personal/increase/deposits/detail/2491',
             'https://www.gazprombank.ru/personal/increase/deposits/detail/1929']
    ammounts = ['500000', '1000000', '2000000', '3000000',
                '5000000', '10000000', '20000000']
    scroll = [1200, 1400]
    times = [7, 6, 7, 7, 7, 7, 8]
    for b in [0, 1]:
        browser.get(hrefs[b])
        time.sleep(2)
        browser.execute_script(f"window.scrollBy(0, {scroll[b]})")
        time.sleep(1)
        col = 66
        try:
            ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
            while True:
                time.sleep(1)
                main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "nr-layout"))
                )  # .chips_tag_group_root__tag--c5c08
                buttons = main.find_elements(By.CSS_SELECTOR, '.chips_tag_root--3b499')
                if buttons:
                    break
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
            print(e, "Program crashed", sep='\n')

        row_f += 1
    return row_f


if __name__ == '__main__':
    PATH = "C:\\Program Files (x86)\\geckodriver.exe"

    options = webdriver.FirefoxOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--incognito')
    options.add_argument('window-size=1920,1080')
    # options.add_argument('--headless')

    row = 3
    wb = openpyxl.open('automated.xlsx')
    sheet = wb.worksheets[3]

    browser = webdriver.Firefox(executable_path=PATH, options=options)

    row = gazprom_write(browser, sheet, row)
    take_tables(browser, sheet, row)
    browser.quit()
    wb.save('automated.xlsx')
