import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common import action_chains
from selenium.webdriver.support import expected_conditions as EC


def press_button(driver_, button):
    print(button.text)
    action = webdriver.common.action_chains.ActionChains(driver_)
    action.move_to_element_with_offset(button, 2, 2)
    action.click()
    action.perform()


def parsing_gaz(driver_, filter_, query):
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    while True:
        time.sleep(1)
        main_ = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.CLASS_NAME, "nr-layout"))
        )
        search = main_.find_element(filter_, query)
        if search:
            print(search.text)
            break


def gazprom(driver: webdriver):
    try:
        driver.get('https://www.gazprombank.ru/personal/increase/deposits/detail/2491')
        driver.execute_script(f"window.scrollBy(0, 1200)")
        time.sleep(1)
        ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
        for i in range(2):
            while True:
                time.sleep(1)
                main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "nr-layout"))
                )
                buttons = main.find_elements(By.CSS_SELECTOR, '.chips_tag_root--3b499')
                if buttons:
                    break
            for button in buttons:
                press_button(driver, button)
                query = '.calculator_results_row_root__value--e5028.calculator_results_row_rate--e5028'
                parsing_gaz(driver, By.CSS_SELECTOR, query)
            print('-' * 20)
            if i == 0:
                driver.get('https://www.gazprombank.ru/personal/increase/deposits/detail/1929')

    except Exception as e:
        print(e, "Program crashed", sep='\n')

    finally:
        print("gazprom finished excetuion")


if __name__ == '__main__':
    PATH = "C:\\Program Files (x86)\\chromedriver.exe"

    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--incognito')
    # options.add_argument('--start-maximized')
    options.add_argument('window-size=1920,1080')
    options.add_argument('--headless')
    browser = webdriver.Chrome(executable_path=PATH, chrome_options=options)
    gazprom(browser)
    browser.quit()
