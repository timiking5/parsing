import time
import openpyxl
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common import action_chains


def use_button_alfa(driver_, filter_, class_name=".a28gX.k28gX.h28gX.l28gX.c28gX.cxDc6"):
    # Тут было что-то
    resp = ''
    while True:
        time.sleep(1)
        main = WebDriverWait(driver_, 10).until(
            EC.presence_of_element_located((By.ID, "alfa"))
        )
        button = main.find_element(filter_, class_name)
        if button:
            resp = button.text
            action = webdriver.common.action_chains.ActionChains(driver_)
            action.move_to_element_with_offset(button, 2, 2)
            action.click()
            action.perform()
            break

    return resp


def use_bar(driver, request, times):
    main = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "alfa"))
    )
    bar = main.find_element(By.CSS_SELECTOR, ".input__input_bmmfk.input__hasLabel_bmmfk.amount-input__input_1udd5")

    if bar:
        for _ in range(times):
            bar.send_keys(Keys.BACKSPACE)
        bar.send_keys(Keys.BACK_SPACE, request)


def parsing_alpha_2(driver_, row_, col_, active_sheet, class_name):
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    while True:
        time.sleep(1)
        main_ = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.ID, "alfa"))
        )
        search = main_.find_element(By.CLASS_NAME, class_name)
        if search:
            # print(f'{chr(col_)}{row_}', search.text)
            active_sheet[f'{chr(col_)}{row_}'] = search.text
            break


def alfa_write(driver: webdriver, sheet):
    try:
        ammounts = ['500000', '1000000', '2000000', '3000000',
                    '5000000', '10000000', '20000000']
        times = [6, 5, 6, 6, 6, 6, 7]
        driver.execute_script(f"window.scrollBy(0, {800})")
        time.sleep(1)
        row = 2
        col = 66

        for j in range(1, 8):
            xpath = f'//*[@id="calculator"]/div[2   ]/div/div[2]/div/div/div[1]/div[2]/button[{j}]'
            sheet[f'{chr(col + j - 1)}1'] = use_button_alfa(driver, By.XPATH, xpath)

        for j in range(7):
            print(f"Proccesing ({j + 1}/7)...")
            use_bar(driver, ammounts[j], times[j])
            sheet[f'A{row}'] = ammounts[j]
            for i in range(1, 8):
                xpath = f'//*[@id="calculator"]/div[2]/div/div[2]/div/div/div[1]/div[2]/button[{i}]'
                use_button_alfa(driver, By.XPATH, xpath)
                parsing_alpha_2(driver, row, col + i - 1, sheet, "g1f0h")
            row += 1

    except Exception as e:
        print(e, "program crashed", sep='\n')

    finally:
        print("program execution complete")


if __name__ == '__main__':
    wb = openpyxl.open('automated.xlsx')
    table = wb.worksheets[2]
    PATH = "C:\\Program Files (x86)\\chromedriver.exe"

    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--incognito')
    options.add_argument('window-size=1920,1080')
    options.add_argument('--headless')

    browser = webdriver.Chrome(executable_path=PATH, chrome_options=options)
    browser.get("https://alfabank.ru/make-money/deposits/alfa/?platformId=alfasite")

    alfa_write(browser, table)

    wb.save('automated.xlsx')
    browser.quit()
    