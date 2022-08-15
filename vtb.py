import time
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def parsing_vtb(driver_, class_name, a=19, b=25):
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    while True:
        time.sleep(1)
        main = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.ID, "root"))
        )
        try:
            search = main.find_elements(By.CSS_SELECTOR, class_name)
            if search:
                for i in range(a, b):
                    print(search[i].text)
                break
        except Exception:
            continue


def parsing_vtb_2(driver_, tag_name, class_name, a=6, b=12):
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    while True:
        time.sleep(1)
        main = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.ID, "root"))
        )
        try:
            search1 = main.find_elements(By.TAG_NAME, tag_name)
            search2 = main.find_elements(By.CSS_SELECTOR, class_name)
            if search1 and search2:
                for i in range(a):
                    print(search1[i].text)
                for i in range(b):
                    print(search2[i].text)
                break
        except Exception:
            pass


def vtb(driver: webdriver):
    try:
        classes = [".typographystyles__Box-foundation-kit__sc-14qzghz-0.gwYyod.markdown-paragraphstyles__ParagraphTypography-foundation-kit__sc-otngat-0.fTcTcY",
                   ".typographystyles__Box-foundation-kit__sc-14qzghz-0.fYipUR.numbersstyles__TypographyTitle-foundation-kit__sc-1xhbrzd-4.oZTKn"]
        driver.get('https://www.vtb.ru/personal/vklady-i-scheta/nakopitelny-schet-seif/')
        parsing_vtb(driver, classes[0])

        print('-' * 20)

        driver.get('https://www.vtb.ru/personal/vklady-i-scheta/nakopitelny-schet-kopilka/')
        parsing_vtb(driver, classes[0])

        print("-" * 20)

        driver.get('https://www.vtb.ru/personal/vklady-i-scheta/vklad-perviy/')
        parsing_vtb(driver, classes[0])
        parsing_vtb(driver, classes[1], 0, 8)

        print('-' * 20)

        driver.get('https://www.vtb.ru/personal/vklady-i-scheta/vklad-ypravlyaemiy/')
        parsing_vtb_2(driver, "h2", classes[1])

        print("-" * 20)

        driver.get('https://www.vtb.ru/personal/vklady-i-scheta/nalichniy/')
        parsing_vtb_2(driver, "h2", classes[1], 2, 2)

        print("-" * 20)

        driver.get('https://www.vtb.ru/personal/vklady-i-scheta/novoe-vremya/')
        parsing_vtb(driver, classes[0], 19, 25)

    except Exception as e:
        print(e)
        print("shit crashed")
    finally:
        print("vtb finished execution")


if __name__ == '__main__':
    PATH = "C:\\Program Files (x86)\\chromedriver.exe"
    # class_ = """//*[@id="seif-2"]/div/div[2]/div/div[1]/div/div/div[1]/div[2]/div[1]/div/div/div/div/div"""

    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    # options.add_argument('--start-maximized')
    options.add_argument('--incognito')
    options.add_argument('window-size=1920,1080')
    # options.add_argument('--headless')
    # service=Service(ChromeDriverManager().install())
    browser = webdriver.Chrome(executable_path=PATH, chrome_options=options)
    vtb(browser)
    browser.quit()
