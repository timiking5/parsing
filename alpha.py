import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


PATH = "C:\\Program Files (x86)\\chromedriver.exe"


options = webdriver.ChromeOptions()
options.add_argument('--ignore-certificate-errors')
options.add_argument('--incognito')
options.add_argument('window-size=1920,1080')
options.add_argument('--headless')

driver = webdriver.Chrome(executable_path=PATH, chrome_options=options)
driver.get("https://alfabank.ru/")

try:
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Вклады"))
    )
    element.click()
    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Альфа-Счёт"))
    )
    element.click()
    main = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "alfa"))
    )
    search = main.find_elements(By.CLASS_NAME, "e1Hrp")
    if not search:
        print("elements not found")
    else:
        for item in search:
            print(item.text)

except Exception as e:
    print(e, "program crashed", sep='\n')

finally:
    print("program execution complete")
    time.sleep(3)
    driver.quit()
