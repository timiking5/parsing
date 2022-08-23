import time
import warnings
import openpyxl
from datetime import date
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement


ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)


def use_bar(bar: WebElement, request):
    bar.send_keys([Keys.BACKSPACE for _ in range(9)], request)


def alfa_write_curr(driver, active_table, row_f):
    global ignored_exceptions
    browser.get('https://alfabank.ru/make-money/deposits/alfa/?platformId=alfasite')
    time.sleep(3)
    # driver.execute_script(f"window.scrollBy(0, {800})")
    # time.sleep(1)
    char = ['$', '€']
    bar_selector = '.input__input_bmmfk.input__hasLabel_bmmfk.amount-input__input_1udd5'
    amounts = ["1000", "25000", "50000", "100000", "300000"]
    for k in [2, 3]:
        print(f"    proccesing {k}/3...")
        xpath_switch = f'//*[@id="calculator"]/div[2]/div/div[1]/button[{k}]'
        main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.ID, "alfa"))
            )
        button = main.find_element(By.XPATH, xpath_switch)
        button.click()
        div = main.find_elements(By.CSS_SELECTOR, '.r2lIQ')[k-1]
        buttons = div.find_elements(By.TAG_NAME, 'button')

        bar = main.find_elements(By.CSS_SELECTOR, bar_selector)[k - 1]

        for i in range(len(buttons)):
            active_table[f'{chr(66 + i)}{row_f}'] = buttons[i].text
            active_table[f'{chr(76 + i)}{row_f}'] = buttons[i].text
        row_f += 1
        for i in range(len(amounts)):
            active_table[f'{chr(65)}{row_f}'] = amounts[i] + char[k - 2]
            active_table[f'{chr(75)}{row_f}'] = amounts[i] + char[k - 2]
            use_bar(bar, amounts[i])
            for j in range(len(buttons)):
                buttons[j].click()
                active_table[f'{chr(66 + j)}{row_f}'] = main.find_elements(By.CSS_SELECTOR, '.g1f0h')[6 * (k - 1)].text
                active_table[f'{chr(66 + j + 10)}{row_f}'] = main.find_elements(By.CSS_SELECTOR, '.g1f0h.h1f0h')[k - 1].text
            row_f += 1
        row_f += 1
        time.sleep(2)


def alfa_write(driver: webdriver, active_sheet, row_f):
    global ignored_exceptions
    active_sheet['A1'] = "Розница"
    active_sheet['K1'] = "Премиум"
    driver.get('https://alfabank.ru/make-money/deposits/alfa/?platformId=alfasite')
    time.sleep(3)
    driver.execute_script(f"window.scrollBy(0, {800})")
    time.sleep(1)
    main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.ID, "alfa"))
        )
    bar = main.find_element(By.CSS_SELECTOR, '.input__input_bmmfk.input__hasLabel_bmmfk.amount-input__input_1udd5')
    buttons = find_buttons(main)
    amounts = ['500000', '1000000', '2000000', '3000000',
               '5000000', '10000000', '20000000']
    switch = main.find_elements(By.CSS_SELECTOR, '.i1HFB')[1]
    switch.click()  #
    change_terms = main.find_elements(By.CSS_SELECTOR, '.a1GLT.g1GLT.j1GLT.l1GLT')
    active_sheet[f'A{row_f}'] = 'Без опций'
    active_sheet[f'A{row_f + 8}'] = 'С пополнением'
    active_sheet[f'A{row_f + 16}'] = 'С пополнением и снятием'
    for i in range(len(buttons)):
        active_sheet[f'{chr(66 + i)}{row_f}'] = buttons[i].text
        active_sheet[f'{chr(66 + i + 10)}{row_f}'] = buttons[i].text

    row_f += 1
    for i in range(len(amounts)):
        use_bar(bar, amounts[i])
        active_sheet[f'A{row_f}'] = amounts[i]
        active_sheet[f'A{row_f + 8}'] = amounts[i]
        active_sheet[f'A{row_f + 16}'] = amounts[i]
        buttons = find_buttons(main)
        for j in range(len(buttons)):
            xpath = f'//*[@id="calculator"]/div[2]/div/div[2]/div/div/div[1]/div[2]/button[{j+1}]'
            click_button(main, xpath)
            a, b = find_percent(main)
            active_sheet[f'{chr(66 + j)}{row_f}'] = a
            active_sheet[f'{chr(66 + j + 10)}{row_f}'] = b
            if j <= 3:
                change_terms[0].click()
                a, b = find_percent(main)
                active_sheet[f'{chr(66 + j)}{row_f + 8}'] = a
                active_sheet[f'{chr(66 + j + 10)}{row_f + 8}'] = b
                change_terms[1].click()
                a, b = find_percent(main)
                active_sheet[f'{chr(66 + j)}{row_f + 16}'] = a
                active_sheet[f'{chr(66 + j + 10)}{row_f + 16}'] = b
                change_terms[1].click()
                change_terms[0].click()

        row_f += 1


def find_buttons(main: WebElement):
    buttons = []
    for i in range(7):
        xpath = f'//*[@id="calculator"]/div[2]/div/div[2]/div/div/div[1]/div[2]/button[{i+1}]'
        buttons.append(main.find_element(By.XPATH, xpath))
    return buttons


def click_button(main: WebElement, xpath):
    main.find_element(By.XPATH, xpath).click()


def find_percent(main: WebElement):
    percent = main.find_element(By.CSS_SELECTOR, '.g1f0h')
    percent_premium = main.find_element(By.CSS_SELECTOR, '.g1f0h.h1f0h')
    return percent.text, percent_premium.text


if __name__ == '__main__':
    wb = openpyxl.open(date.today().strftime("%d.%m.%y") + '.xlsx')
    table = wb.worksheets[8]
    table_1 = wb.worksheets[9]
    PATH = "C:\\Program Files (x86)\\chromedriver.exe"
    warnings.filterwarnings("ignore")

    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--incognito')
    options.add_argument('window-size=1920,1080')
    options.add_argument('--headless')

    browser = webdriver.Chrome(executable_path=PATH, chrome_options=options)

    browser.get('https://alfabank.ru/make-money/deposits/alfa/')
    row = 2
    print("    proccesing 1/3...")
    alfa_write(browser, table, row)
    alfa_write_curr(browser, table_1, row)
    #
    browser.quit()
    wb.save(date.today().strftime("%d.%m.%y") + '.xlsx')
