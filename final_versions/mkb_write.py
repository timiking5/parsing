import time
import openpyxl
import warnings
from datetime import date
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common import action_chains
from selenium.webdriver.common.keys import Keys


def mkb_get_table_curr(driver: webdriver, active_sheet, row_f):
    driver.get("https://mkb.ru/personal/deposits/allinclusive")
    driver.execute_script(f"window.scrollBy(0, {1000})")
    time.sleep(1)
    main = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
        )
    buttons = main.find_elements(By.CSS_SELECTOR, '.navigation__item.js-event-markup')
    for button in buttons:
        if "Тарифы" in button.text:
            button.click()
            break
    time.sleep(1)
    driver.execute_script(f"window.scrollBy(0, {600})")
    time.sleep(1)
    # driver.execute_script(f"window.scrollBy(0, {600})")
    # time.sleep(1)

    currencies = main.find_elements(By.CSS_SELECTOR, '.field.field_radio')

    currencies[1].click()
    time.sleep(1)
    ths_1 = main.find_element(By.CSS_SELECTOR, '.table__header').find_elements(By.CSS_SELECTOR, '.table__column')
    trs_1 = main.find_elements(By.CSS_SELECTOR, '.table__row')

    currencies[2].click()
    time.sleep(1)
    ths_2 = main.find_element(By.CSS_SELECTOR, '.table__header').find_elements(By.CSS_SELECTOR, '.table__column')
    trs_2 = main.find_elements(By.CSS_SELECTOR, '.table__row')

    for i in range(len(ths_1)):
        active_sheet[f'{chr(75 + i)}{row_f}'] = ths_1[i].text
    row_f += 1

    for i in range(len(trs_1)):
        tds = trs_1[i].find_elements(By.CSS_SELECTOR, '.table__column')
        if tds:
            for j in range(len(tds)):
                active_sheet[f'{chr(75 + j)}{row_f}'] = tds[j].text
            row_f += 1

    row_f += 16

    for i in range(len(ths_2)):
        active_sheet[f'{chr(75 + i)}{row_f}'] = ths_2[i].text
    row_f += 1

    for i in range(len(trs_2)):
        tds = trs_2[i].find_elements(By.CSS_SELECTOR, '.table__column')
        if tds:
            for j in range(len(tds)):
                active_sheet[f'{chr(75 + j)}{row_f}'] = tds[j].text
            row_f += 1


def start_curr(driver: webdriver, active_sheet, row_f):
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    periods = ['3 мес.', '6 мес.', '1 год', '18 мес.', '2 года']
    char = ["$", "€"]
    slide_by = [0, 75, 75, 75, 75]
    slide = [0, -600]
    amounts = ["1000", "25000", "50000", "100000", "300000"]
    driver.get('https://mkb.ru/personal/deposits/allinclusive')
    driver.execute_script(f"window.scrollBy(0, {1000})")
    time.sleep(1)
    main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
        EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
    )
    switch = main.find_element(By.CSS_SELECTOR, '.vs__search')
    for i in range(2):
        switch.click()
        switch.send_keys(Keys.ARROW_DOWN, Keys.RETURN)
        main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
        )
        slider = main.find_elements(By.CSS_SELECTOR, '.vue-slider-dot-handle')[1]
        webdriver.common.action_chains.ActionChains(driver).drag_and_drop_by_offset(slider, slide[i], 0).perform()

        active_sheet[f'A{row_f - 1}'] = char[i] + " без опций"
        row_f = mkb_write(driver, active_sheet, row_f, slide_by, periods, 1, amounts)
        webdriver.common.action_chains.ActionChains(driver).drag_and_drop_by_offset(slider, slide[1], 0).perform()

        active_sheet[f'A{row_f - 1}'] = char[i] + " с пополнением"
        use_button(driver, "Хочу пополнять")
        row_f = mkb_write(driver, active_sheet, row_f, slide_by, periods, 1, amounts)
        webdriver.common.action_chains.ActionChains(driver).drag_and_drop_by_offset(slider, slide[1], 0).perform()

        active_sheet[f'A{row_f - 1}'] = char[i] + " с пополнением и снятием"
        use_button(driver, "Хочу снимать")
        row_f = mkb_write(driver, active_sheet, row_f, slide_by, periods, 1, amounts)

        use_button(driver, "Хочу пополнять")
        use_button(driver, "Хочу снимать")


def get_table(driver, active_sheet, row_f):
    """
    Заново проходится по всем сайтам и забирает таблицы. Это происходит очень быстро + мне не хотелось перегружать
    start, mkb_write, поэтому выделил в отдельную функцию. Я записываю в ряд 3 (если вы ничего не поменяли), но в
    столбце I, так удобнее, как мне кажется.
    :param driver: browser
    :param active_sheet: Лист excel, куда записывать
    :param row_f: ряд, в который записывать
    :return:
    """
    hrefs = ['https://mkb.ru/personal/deposits/mega-online',
             'https://mkb.ru/personal/deposits/allinclusive',
             'https://mkb.ru/personal/deposits/30years']
    col = 75
    for href in hrefs:
        try:
            driver.get(href)
            ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
            )
            driver.execute_script(f"window.scrollBy(0, {700})")
            buttons = main.find_elements(By.CSS_SELECTOR, '.navigation__item.js-event-markup')
            for button in buttons:
                if 'Тарифы' in button.text:
                    action = webdriver.common.action_chains.ActionChains(driver)
                    action.move_to_element_with_offset(button, 2, 2)
                    action.click()
                    action.perform()

            active_sheet[f'{chr(col)}{row_f}'] = driver.find_element(By.TAG_NAME, 'h2').text
            row_f += 1

            main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
                EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
            )
            headers = main.find_element(By.CSS_SELECTOR, '.table__header').find_elements(By.CSS_SELECTOR, '.table__column')
            rows = main.find_elements(By.CSS_SELECTOR, '.table__row')

            for i in range(len(headers)):
                active_sheet[f'{chr(col + i)}{row_f}'] = headers[i].text
            row_f += 1
            for i in range(len(rows)):
                elements = rows[i].find_elements(By.CSS_SELECTOR, '.table__column')
                elements = [element for element in elements if element.text != '']
                if elements:
                    for j in range(len(elements)):
                        active_sheet[f'{chr(col + j)}{row_f}'] = elements[j].text
                    row_f += 1
            row_f += 1

        except Exception as e:
            print(e)


def use_button(driver_, query):
    """
    Нажимает кнопку, имеющую соответствующий текст в ней
    :param driver_: browser
    :param query: Текст кнопки
    :return: Можете добавить в случае успеха 1, неудачи - 0. Но это может быть лишним.
    """
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    main = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
        EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
    )
    buttons = main.find_elements(By.CSS_SELECTOR, '.field__caption')
    for button in buttons:
        if query in button.text:
            action = webdriver.common.action_chains.ActionChains(driver_)
            action.move_to_element_with_offset(button, 2, 2)
            action.click()
            action.perform()
            # break


def use_bar(driver_, request):
    """
    Использует вводное поле.
    :param driver_: browser
    :param request: текст, который надо передать, передавайте, пожалуйста, только цифры, также смотрите, чтобы класс
    не поменялся
    :return:
    """
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    main = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
        EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
    )
    bar = main.find_element(By.CSS_SELECTOR, '.field__input')
    if bar:
        action = webdriver.common.action_chains.ActionChains(driver_)
        action.click(bar)
        action.send_keys([Keys.BACK_SPACE for _ in range(10)])
        # action.pause(1)
        action.send_keys(request)
        action.send_keys(Keys.RETURN)
        action.perform()
        """
        Такая убогая реализация, потому что нормальная не работала. Отправляет BACKSPACE 10 раз, чтобы поле было пустое,
        а потом вставляет сам запрос.
        """


def drag_slider(driver_, main, offset):
    """
    Двигает слайдер
    :param driver_: browser
    :param main: Слегка ускоряет процесс, потому что не приходится грузить его в этой функции заново.
    :param offset: кол-во пикселей смещения
    :return:
    """
    slider = main.find_elements(By.CSS_SELECTOR, '.vue-slider-dot-handle')[1]
    webdriver.common.action_chains.ActionChains(driver_).drag_and_drop_by_offset(slider, offset, 0).perform()


def parsing_mkb(driver_, active_sheet, row_f, col_f, index):
    """
    Находит нужный процент и записывает его по нужному адресу. Нужно следить, чтобы значение класса не устарело и не
    поменялось (CSS_SELECTOR)
    """
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    while True:
        main_ = WebDriverWait(driver_, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
        )
        search = main_.find_elements(By.CSS_SELECTOR, '.calculate-offer__calculation-value')[index]

        if search:
            active_sheet[f'{chr(col_f)}{row_f}'] = search.text
            break


def start(driver, active_sheet, row_f):
    """
    1)Заходим на mega-online, забираем инфу оттуда.
    2)Заходим на "Всё включено", забираем всю инфу.
    3)Ставим галочку "Хочу пополнять", забираем всю инфу.
    4)Ставим галочку "Хочу снимать", забираем всю инфу.
    *)Не забываем двигать слайдер влево
    :param driver: browser
    :param active_sheet: Лист эксель, куда записывать инфу (нужен пустой лист)
    :param row_f: с какого ряда начинать
    :return:
    """
    hrefs = ['https://mkb.ru/personal/deposits/mega-online',
             'https://mkb.ru/personal/deposits/allinclusive']
    slide_by = [[0, 50, 50, 50, 100, 200, 200],
                [0, 75, 150, 200, 200]]
    ammounts = ['500000', '1000000', '2000000', '3000000',
                '5000000', '10000000', '20000000']
    periods = [['3 мес.', '6 мес.', '9 мес.', '1 год', '18 мес.', '2 года', '3 года'],
               ['3 мес.', '6 мес.', '1 год', '18 мес.', '2 года']]
    print(f"    proccessing 1/4...")
    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)
    driver.get(hrefs[0])
    driver.execute_script(f"window.scrollBy(0, {1000})")
    time.sleep(1)  # Скролить нужно, потому что иначе невозможно нажимать на кнопки/слайдеры (их не видно)
    active_sheet[f'A{row_f - 1}'] = "МЕГА \"Онлайн\""
    row_f = mkb_write(driver, active_sheet=active_sheet, row_f=row_f, slide_by=slide_by[0], periods=periods[0], ind=0, ammounts=ammounts)

    driver.get(hrefs[1])
    driver.execute_script(f"window.scrollBy(0, {1000})")
    time.sleep(1)
    print(f"    proccessing 2/4...")
    active_sheet[f'A{row_f - 1}'] = "Без опций \"Всё включено\""
    row_f = mkb_write(driver, active_sheet=active_sheet, row_f=row_f, slide_by=slide_by[1], periods=periods[1], ind=1, ammounts=ammounts)

    main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
        EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
    )
    drag_slider(driver, main, -500)
    time.sleep(1)

    use_button(driver, 'Хочу пополнять')
    print(f"    proccessing 3/4...")
    active_sheet[f'A{row_f - 1}'] = "С пополнением \"Всё включено\""
    row_f = mkb_write(driver, active_sheet=active_sheet, row_f=row_f, slide_by=slide_by[1], periods=periods[1], ind=1, ammounts=ammounts)

    main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
        EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
    )
    drag_slider(driver, main, -500)
    time.sleep(1)

    use_button(driver, 'Хочу снимать')
    print(f"    proccessing 4/4...")
    active_sheet[f'A{row_f - 1}'] = "С пополнением и снятием \"Всё включено\""
    row_f = mkb_write(driver, active_sheet=active_sheet, row_f=row_f, slide_by=slide_by[1], periods=periods[1], ind=1, ammounts=ammounts)

    return row_f


def mkb_write(driver, active_sheet, row_f, slide_by, periods, ind, ammounts):
    """
    На конктреной странице находит слайдер и поле ввода. Сначала делает слайд, потом вводит все суммы, потому что
    мне так показалось лучше, чем постоянно дёргать не самый надёжный слайдер.
    :param driver: Browser
    :param active_sheet: Лист в ворде, куда записывается рез-т (Сделайте его пустым)
    :param row_f: С какого ряда идёт запись
    :param slide_by: Параметр, к сожалению, выставляется вручную. Список значений, на которые слайдер передвигается
    вправо.
    :param periods: Параметр, идущий совместно с slide_by. Слайдер передвинулся вправо, но какой стал период?
    К сожалению, на странице не всегда есть информация о периоде (например, когда период равен трём месяцам), поэтому
    решено было сделать так.
    :param ind: Иногда нужный процент является либо первым (0), либо вторым (1) представителем класса на странице. Для
    удобства было сделано так, что этот параметр выставляется вручную. Это позволяет также избежать некоторых багов и
    неточностей.
    :return:
    """

    ignored_exceptions = (NoSuchElementException, StaleElementReferenceException,)

    main = WebDriverWait(driver, 10, ignored_exceptions=ignored_exceptions).until(
            EC.presence_of_element_located((By.CLASS_NAME, "js-navigation"))
        )
    for i in range(len(ammounts)):
        active_sheet[f'A{row_f + i}'] = ammounts[i]
    col = 66
    for i in range(len(slide_by)):
        drag_slider(driver, main, slide_by[i])
        time.sleep(1)
        active_sheet[f'{chr(col)}{row_f - 1}'] = periods[i]
        for j in range(len(ammounts)):
            use_bar(driver, ammounts[j])
            parsing_mkb(driver, active_sheet, row_f + j, col, index=ind)
        col += 1
    row_f += len(ammounts) + 1

    return row_f


if __name__ == '__main__':
    PATH = "C:\\Program Files (x86)\\geckodriver.exe"
    warnings.filterwarnings("ignore")

    options = webdriver.FirefoxOptions()
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--incognito')
    options.add_argument('window-size=1920,1080')
    # options.add_argument('--headless')
    """
    Классическое предупреждение для будущих разработчиков: вроде как в селениуме собираются убрать параметр
        executable_path (см. DeprecationWarning при запуске), но сейчас же у меня, как бы я ни пытался, не получается
        использовать service=... . Возможно, в будущем это уже будет работать, так что вам придётся с этим запаритсья.
        Вместо chrome_options= можно использовать options= ... . Аргумент headless можно использовать по желанию, но
        иногда он приводит к тому, что программа не работает.
        """
    browser = webdriver.Firefox(executable_path=PATH, options=options)
    row = 3
    wb = openpyxl.open(date.today().strftime("%d.%m.%y") + '.xlsx')
    sheet = wb.worksheets[0]
    sheet_1 = wb.worksheets[1]
    start(browser, sheet, row)
    get_table(browser, sheet, row)
    start_curr(browser, sheet_1, row)
    mkb_get_table_curr(browser, sheet_1, row)
    browser.quit()
    wb.save(date.today().strftime("%d.%m.%y") + '.xlsx')
