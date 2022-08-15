import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def parsing_sber(driver_, filter_, query, a=0, b=2):
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    while True:
        time.sleep(1)
        main_ = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.ID, "main-page"))
        )
        try:
            search = main_.find_elements(filter_, query)
            if search:
                for i in range(a, b):
                    print(search[i].text)
                break
        except Exception:
            continue


def sber(driver: webdriver):
    try:
        driver.get("https://www.sberbank.ru/ru/person/contributions/deposits/nakopi")
        parsing_sber(driver, By.CLASS_NAME, "nswbonus-rates__card")

        print('-' * 20)

        driver.get("https://www.sberbank.ru/ru/person/contributions/deposits/sbervklad_a")
        parsing_sber(driver, By.TAG_NAME, "thead", b=1)
        parsing_sber(driver, By.TAG_NAME, "tbody", b=1)

        print('-' * 20)

        driver.get('https://www.sberbank.ru/ru/person/contributions/deposits/av')
        parsing_sber(driver, By.CLASS_NAME, "aage-rates__card")

    except Exception as e:
        print(e, "Program crashed", sep='\n')

    finally:
        print("sber finished execution")


if __name__ == '__main__':
    PATH = "C:\\Program Files (x86)\\geckodriver.exe"

    # options = webdriver.ChromeOptions()
    # options.add_argument('--ignore-certificate-errors')
    # options.add_argument('--incognito')
    # options.add_argument('--disable-gpu')
    # options.add_argument('window-size=1920,1080')
    # options.add_argument('--headless')
    options = webdriver.FirefoxOptions()
    options.add_argument('--headless')


    browser = webdriver.Firefox(executable_path=PATH, options=options)  # executable_path=PATH, chrome_options=options
    sber(browser)
    browser.quit()
