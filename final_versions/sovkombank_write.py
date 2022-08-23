import warnings
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.remote.webelement import WebElement


def sovkom_write(driver: webdriver, active_sheet, row_f):
    hrefs = ['https://sovcombank.ru/deposits/spring-income',
             'https://sovcombank.ru/deposits/vklad-udobniy']
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    for k in [0, 1]:
        print(f'    processing {1 + k}/2...')
        driver.get(hrefs[k])
        main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".Layout__content"))
        )
        tables = main.find_elements(By.TAG_NAME, 'table')
        table = None
        for i in tables:
            if i.text != '':
                table = i

        if table:
            rows = table.find_elements(By.TAG_NAME, 'tr')
            for i in range(len(rows)):
                tds = rows[i].find_elements(By.TAG_NAME, 'td')
                if k == 1 and i == 0:
                    active_sheet[f'A{row_f}'] = 'Срок в месяцах'
                    active_sheet[f'B{row_f}'] = tds[0].text
                    active_sheet[f'C{row_f}'] = tds[1].text
                    active_sheet[F'D{row_f}'] = tds[2].text
                else:
                    for j in range(len(tds)):
                        active_sheet[f'{chr(65 + j)}{row_f}'] = tds[j].text
                row_f += 1
            row_f += 1
    return row_f


if __name__ == '__main__':
    wb = openpyxl.open('testing.xlsx')
    sheet = wb.worksheets[2]

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
    try:
        sovkom_write(browser, sheet, row)
    except Exception as e:
        print(e)
    finally:
        wb.save('testing.xlsx')
        browser.quit()
