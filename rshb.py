import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def rshb(driver: webdriver):
    try:
        driver.get("https://www.rshb.ru/natural/deposits/")
        while True:
            main = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "page"))
            )
            search = main.find_element(By.TAG_NAME, "table")
            if search:
                break
        soup = BeautifulSoup(search.get_attribute('innerHTML'), 'html.parser')
        rows = soup.find_all("tr")
        for i, row in enumerate(rows):
            tag = 'td'
            if i == 0:
                tag = 'th'
            values = row.find_all(tag)
            for j in range(len(values)):
                if i != 0 and (j == 1 or j == 5 or j == 6):
                    resp = values[j].find('span')
                    if resp:
                        print(resp.get('title'), end=' ')
                    else:
                        print("Не предусмотрено", end=' ')
                else:
                    print(values[j].getText(strip=True, separator=' '), end=' ')
            print('\n')

    except Exception as e:
        print(e, 'programm crashed', sep='\n')

    finally:
        print("rshb finished execution")


if __name__ == '__main__':
    PATH = "C:\\Program Files (x86)\\chromedriver.exe"
    # xpath = "/html/body/div[8]/div[3]/div[7]/table/tbody/tr[2]/td[9]/p"

    options = webdriver.ChromeOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--incognito')
    options.add_argument('window-size=1920,1080')
    options.add_argument('--headless')
    # /html/body/div[7]/div[3]/div[7]/table/tbody/tr[1]/th[1]
    browser = webdriver.Chrome(executable_path=PATH, chrome_options=options)
    rshb(browser)
    browser.quit()
