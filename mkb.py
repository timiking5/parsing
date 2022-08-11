import time
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common import action_chains


def use_button(driver_, filter_, query):
    while True:
        time.sleep(1)
        main = WebDriverWait(driver_, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
        )
        button = main.find_element(filter_, query)
        if button:
            action = webdriver.common.action_chains.ActionChains(driver_)
            action.move_to_element_with_offset(button, 2, 2)
            action.click()
            action.perform()
            break


def parsing_mkb(driver_, filter_, query, a=0, b=1):
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    while True:
        main = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
        )
        search = main.find_elements(filter_, query)

        if search:
            for i in range(a, b):
                print(search[i].text)
            break


def mkb(driver: webdriver):
    try:
        driver.get("https://mkb.ru/personal/deposits/advantage")
        parsing_mkb(driver, By.CLASS_NAME, "calculate-offer__calculation-value")

        print('-' * 20)

        driver.get("https://mkb.ru/personal/deposits/savings-account")
        driver.execute_script(f"window.scrollBy(0, {500})")
        time.sleep(1)
        use_button(driver, By.ID, 'tab-2')  # Don't forget to put smth here
        parsing_mkb(driver, By.CLASS_NAME, "table__column", 0, 9)

        print('-' * 20)

        driver.get("https://mkb.ru/personal/deposits/mega-online")
        driver.execute_script(f"window.scrollBy(0, {500})")
        time.sleep(1)
        use_button(driver, By.ID, 'tab-2')  # Don't forget to input
        parsing_mkb(driver, By.CLASS_NAME, "table__column", 0, 33)

    except Exception as e:
        print(e, 'Programm crashed', sep='\n')

    finally:
        print("mkb finished execution")


if __name__ == '__main__':
    PATH = "C:\\Program Files (x86)\\chromedriver.exe"

    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--incognito')
    options.add_argument('window-size=1920,1080')
    # options.add_argument('--headless')

    browser = webdriver.Chrome(executable_path=PATH, chrome_options=options)
    mkb(browser)
    browser.quit()
