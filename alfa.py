import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common import action_chains
from selenium.webdriver.support import expected_conditions as EC


def parsing_alpha(driver_, class_name):
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    while True:
        time.sleep(1)
        main_ = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.ID, "alfa"))
        )
        search = main_.find_elements(By.CLASS_NAME, class_name)
        if search:
            for item in search:
                print(item.text)
            break


def parsing_alpha_2(driver_, class_name):
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    while True:
        time.sleep(1)
        main_ = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.ID, "alfa"))
        )
        search = main_.find_element(By.CLASS_NAME, class_name)
        if search:
            print(search.text)
            break


def use_button_alfa(driver_, filter_, class_name=".a28gX.k28gX.h28gX.l28gX.c28gX.cxDc6"):
    # Тут было что-то
    while True:
        time.sleep(1)
        main = WebDriverWait(driver_, 10).until(
            EC.presence_of_element_located((By.ID, "alfa"))
        )
        button = main.find_element(filter_, class_name)
        if button:

            action = webdriver.common.action_chains.ActionChains(driver_)
            action.move_to_element_with_offset(button, 2, 2)
            action.click()
            action.perform()
            break


def alfa(driver: webdriver):
    try:
        classes = [".a28gX.k28gX.h28gX.l28gX.c28gX.cxDc6",
                   "e1Hrp",
                   "g1f0h"]
        driver.get('https://alfabank.ru/make-money/savings-account/alfa/')
        parsing_alpha(driver, classes[1])

        print("-" * 20)

        driver.execute_script(f"window.scrollBy(0, 500)")
        time.sleep(1)
        use_button_alfa(driver, By.CSS_SELECTOR)
        parsing_alpha(driver, classes[1])

        print("-" * 20)

        driver.get("https://alfabank.ru/make-money/deposits/alfa/?platformId=alfasite")
        driver.execute_script(f"window.scrollBy(0, {1200})")
        time.sleep(1)
        parsing_alpha_2(driver, classes[2])
        for i in range(6, 0, -1):
            xpath = f'//*[@id="calculator"]/div[2]/div/div[2]/div/div/div[1]/div[2]/button[{i}]'
            use_button_alfa(driver, By.XPATH, xpath)
            parsing_alpha_2(driver, classes[2])

    except Exception as e:
        print(e, "program crashed", sep='\n')

    finally:
        print("alfa finished execution")


if __name__ == '__main__':
    PATH = "C:\\Program Files (x86)\\chromedriver.exe"

    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--incognito')
    # options.add_argument('--start-maximized')
    options.add_argument('window-size=1920,1080')
    options.add_argument('--headless')
    browser = webdriver.Chrome(executable_path=PATH, chrome_options=options)
    alfa(browser)
    browser.quit()
