import time
from selenium import webdriver
import openpyxl
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common import action_chains
from selenium.webdriver.support import expected_conditions as EC


row = 2


def gazprom_get_urls(driver, active_sheet):
    global row
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    hrefs = []
    while True:
        main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.CLASS_NAME, "nr-layout__top"))
        )
        search = main.find_elements(By.CSS_SELECTOR, '.nr-advantages-blocks-item__content-button')
        if search:
            for elem in search:
                if 'подробнее о счете' in elem.text.lower():
                    hrefs.append(elem.get_attribute('href'))
            break
    driver.set_window_size(1124, 850)
    for href in hrefs:
        driver.get(href)
        driver.execute_script(f"window.scrollBy(0, {700})")
        time.sleep(1)
        fl = 1
        percents = []
        while True:
            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.CLASS_NAME, "nr-layout__top"))
            )
            buttons = main.find_elements(By.CSS_SELECTOR, '.nr-tabs__el--inner')
            if buttons:
                for button in buttons:
                    if 'Процентные ставки' in button.text:
                        action = webdriver.common.action_chains.ActionChains(driver)
                        action.move_to_element_with_offset(button, 2, 2)
                        action.click()
                        action.perform()
                        break
                else:
                    fl = 0
                break
        time.sleep(1)
        if fl:
            while True:
                main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "nr-layout__top"))
                )
                try:
                    search_td = main.find_elements(By.TAG_NAME, 'td')
                    account_name = main.find_element(By.TAG_NAME, 'h1')
                    if account_name:
                        active_sheet[f'B{row}'] = account_name.text
                    if search_td:
                        for elem in search_td:
                            for word in elem.text.split():
                                if '%' in word and '+' not in word and len(word) >= 2:
                                    percents.append(word)
                        break
                except Exception:
                    pass

            if len(percents) == 1:
                active_sheet[f'C{row}'] = percents[0]
                active_sheet[f'D{row}'] = percents[0]
                row += 1
            elif len(percents) >= 2:
                if (percents[1] == '0,51%'):
                    percents[1], percents[3] = percents[3], percents[1]
                active_sheet[f'D{row}'] = percents[0]
                active_sheet[f'C{row}'] = percents[1]
                row += 1
            else:
                row += 1
        else:
            while True:
                main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "nr-layout__top"))
                )
                try:
                    search_info = main.find_elements(By.CSS_SELECTOR, '.nr-advantages-markers-el__subtitle')
                    account_name = main.find_element(By.TAG_NAME, 'h1')
                    if account_name:
                        active_sheet[f'B{row}'] = account_name.text
                    if search_info:
                        for elem in search_info:
                            for word in elem.text.split():
                                if '%' in word:
                                    percents.append(word)
                        break
                except Exception:
                    pass
            if len(percents) == 1:
                active_sheet[f'C{row}'] = percents[0]
                active_sheet[f'D{row}'] = percents[0]
                row += 1
            elif len(percents) >= 2:
                active_sheet[f'C{row}'] = percents[0]
                active_sheet[f'D{row}'] = percents[len(percents) - 1]
                row += 1
            else:
                row += 1



def rshb_get_urls(driver, active_sheet):
    global row
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    hrefs = ['https://www.rshb.ru/natural/deposits/savings_moneybox/']
    for href in hrefs:
        driver.get(href)
        percents = []
        while True:
            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.CLASS_NAME, "page"))
            )
            account_name = main.find_element(By.TAG_NAME, 'h1')
            search = main.find_elements(By.TAG_NAME, 'td')
            if account_name:
                active_sheet[f'B{row}'] = account_name.text
            if search:
                for elem in search:
                    if '%' in elem.text and '0,' not in elem.text:
                        percents.append(elem.text)
                break

    if len(percents) == 1:
        active_sheet[f'C{row}'] = percents[0]
        active_sheet[f'D{row}'] = percents[0]
        row += 1
    elif len(percents) >= 2:
        active_sheet[f'C{row}'] = percents[0]
        active_sheet[f'D{row}'] = percents[1]
        row += 1
    else:
        row += 1


def mkb_get_urls(driver, active_sheet):
    global row
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    hrefs = []
    while True:
        main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.CLASS_NAME, "section__container"))
        )
        search = main.find_elements(By.CSS_SELECTOR, '.deposits-item__title')
        if search:
            for elem in search:
                if 'счет' in elem.text.lower():
                    hrefs.append(elem.get_attribute('href'))
            break
    for href in hrefs:
        driver.get(href)
        percents = []
        driver.execute_script(f"window.scrollBy(0, {300})")
        time.sleep(1)
        while True:
            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
            )
            buttons = main.find_elements(By.TAG_NAME, 'li')
            if buttons:
                for button in buttons:
                    if 'Тарифы' in button.text:
                        action = webdriver.common.action_chains.ActionChains(driver)
                        action.move_to_element_with_offset(button, 2, 2)
                        action.click()
                        action.perform()
                break
        while True:
            main_1 = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
            )
            main_2 = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.CLASS_NAME, "main-banner__body"))
            )
            account_name = main_2.find_element(By.TAG_NAME, 'h2')
            search = main_1.find_elements(By.CSS_SELECTOR, '.table__column')
            if account_name:
                active_sheet[f'B{row}'] = account_name.text
            if search:
                for elem in search:
                    if '%' in elem.text:
                        percents.append(elem.text)
                break
        if len(percents) == 1:
            active_sheet[f'C{row}'] = percents[0]
            active_sheet[f'D{row}'] = percents[0]
            row += 1
        elif len(percents) >= 2:
            active_sheet[f'C{row}'] = percents[1]
            active_sheet[f'D{row}'] = percents[0]
            row += 1
        else:
            row += 1


def sber_get_urls(driver, active_sheet):
    global row
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    hrefs = []
    css_selector = '.dk-sbol-link.dk-sbol-link_size_md.dk-sbol-link_color_black'
    while True:
        main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.ID, "main-page"))
        )
        try:
            search = main.find_elements(By.CSS_SELECTOR, css_selector)
            if search:
                for elem in search:
                    if 'счёт' in elem.text:
                        hrefs.append(elem.get_attribute('href'))
                break
        except Exception:
            pass
    css_selector = '.dk-sbol-heading.dk-sbol-heading_size_sm.nswbonus-rates__rate'
    for href in hrefs:
        driver.get(href)
        percents = []
        while True:
            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.ID, "main-page"))
            )
            try:
                search = main.find_elements(By.CSS_SELECTOR, css_selector)
                account_name = main.find_element(By.TAG_NAME, 'h1')
                if account_name:
                    active_sheet[f'B{row}'] = account_name.text
                if search:
                    for elem in search:
                        if '%' in elem.text:
                            percents.append(elem.text)
                    break
            except Exception:
                pass
        if len(percents) == 1:
            active_sheet[f'C{row}'] = percents[0]
            active_sheet[f'D{row}'] = percents[0]
            row += 1
        elif len(percents) >= 2:
            active_sheet[f'C{row}'] = percents[0]
            active_sheet[f'D{row}'] = percents[1]
            row += 1
        else:
            row += 1


def alfa_get_urls(driver, active_sheet):
    global row
    css_selector = '.a1jIK.l1jIK.H1jIK.eG2mw.yG2mw.SG2mw.aaG2mw.i1Hrp.c1Hrp'
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    element = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
        EC.presence_of_element_located((By.LINK_TEXT, "Вклады"))
    )
    element.click()
    hrefs = []
    while True:
        main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.ID, "alfa"))
        )
        search = main.find_elements(By.CSS_SELECTOR, '.a1cAc.g1cAc.c1cAc')
        if search:
            for elem in search:
                if 'счёт' in elem.text.lower() and 'специальный' not in elem.text.lower():
                    hrefs.append(elem.get_attribute('href'))
            break
        time.sleep(1)
    for href in hrefs:
        driver.get(href)
        percents = []
        driver.execute_script(f"window.scrollBy(0, {700})")
        time.sleep(1)
        for k in range(10):
            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.ID, "alfa"))
            )
            search = main.find_elements(By.CSS_SELECTOR, '.a28gX.k28gX.h28gX.l28gX.c28gX.cxDc6')
            account_name = main.find_element(By.CSS_SELECTOR, '.a1q6H.c1q6H.n1q6H.r1q6H')
            if account_name:
                active_sheet[f'B{row}'] = account_name.text
            if search:
                for elem in search:
                    if 'не трачу' in elem.text.lower():
                        action = webdriver.common.action_chains.ActionChains(driver)
                        action.move_to_element_with_offset(elem, 2, 2)
                        action.click()
                        action.perform()
                break
            time.sleep(1)
        while True:
            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.ID, "alfa"))
            )
            search = main.find_elements(By.CSS_SELECTOR, css_selector)
            if search:
                for elem in search:
                    if '%' in elem.text:
                        percents.append(elem.text)
                break
        if len(percents) == 1:
            active_sheet[f'C{row}'] = percents[0]
            active_sheet[f'D{row}'] = percents[0]
            row += 1
        elif len(percents) >= 2:
            active_sheet[f'C{row}'] = percents[0]
            active_sheet[f'D{row}'] = percents[1]
            row += 1
        else:
            row += 1


def vtb_get_parse(driver, active_sheet):
    global row
    class_ = '.typographystyles__Box-foundation-kit__sc-14qzghz-0.fYipUR.numbersstyles__TypographyTitle-foundation-kit__sc-1xhbrzd-4.oZTKn'
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    # element = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
    #     EC.presence_of_element_located((By.LINK_TEXT, "Вклады и счета"))
    # )
    # element.click()
    hrefs = ['https://www.vtb.ru/personal/vklady-i-scheta/nakopitelny-schet-seif/',
             'https://www.vtb.ru/personal/vklady-i-scheta/nakopitelny-schet-kopilka/']
    for href in hrefs:
        percents = []
        driver.get(href)
        while True:
            time.sleep(1)
            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.ID, "root"))
            )
            try:
                account_name = main.find_element(By.TAG_NAME, 'h1')
                search = main.find_elements(By.CSS_SELECTOR, class_)
                if account_name:
                    active_sheet[f'B{row}'] = account_name.text
                if search:
                    for item in search:
                        if '%' in item.text:
                            percents.append(item.text)
                    break
            except Exception:
                continue

        if len(percents) == 1:
            active_sheet[f'C{row}'] = percents[0]
            active_sheet[f'D{row}'] = percents[0]
            row += 1
        elif len(percents) >= 2:
            active_sheet[f'C{row}'] = percents[1]
            active_sheet[f'D{row}'] = percents[0]
            row += 1
        else:
            row += 1


PATH = "C:\\Program Files (x86)\\geckodriver.exe"
urls = ['https://alfabank.ru/',
        'https://www.vtb.ru/',
        'https://www.sberbank.ru/ru/person/contributions/depositsnew',
        'https://mkb.ru/personal/deposits',
        'https://www.rshb.ru/natural/deposits/',
        'https://www.gazprombank.ru/personal/page/mob-accounts']
sites = ['alfa', 'vtb', 'sber', 'mkb', 'rshb', 'gazprom']
funcs = [alfa_get_urls, vtb_get_parse, sber_get_urls, mkb_get_urls, rshb_get_urls, gazprom_get_urls]

firefox_options = webdriver.FirefoxOptions()
firefox_options.add_argument('--ignore-certificate-errors')
firefox_options.add_argument('--incognito')
# firefox_options.add_argument('--headless')

firefox = webdriver.Firefox(executable_path=PATH, options=firefox_options)
firefox.set_window_size(1920, 10800)

wb = openpyxl.open('automated.xlsx')
sheet = wb.worksheets[1]

try:
    for i in [5, 2, 3, 4, 0]:  # range(len(urls))
        print(i)
        firefox.get(urls[i])
        sheet[f'A{row}'] = sites[i]
        funcs[i](firefox, sheet)
        time.sleep(1)

except Exception as e:
    print(e, "crash", sep='\n')

finally:
    print("complete")
    time.sleep(3)
    # input()
    wb.save('automated.xlsx')
    firefox.quit()
