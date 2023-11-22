import time
from base_driver.webdriver_project import BaseOptions
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, WebDriverException
from excel_writer.get_and_write_xlsx import worksheet

import re


def _parser(inns, config):
    sleep = int(config.get("program", "sleep"))
    if sleep < 18:
        sleep = 0
    else:
        sleep -= 18

    min_row = int(config.get("program", "write_row"))
    if min_row < 1:
        min_row = 1

    headless = bool(int(config.get("program", "headless")))
    for row in range(min_row - 1, len(inns)):
        print(f'Пройдено {row} из {len(inns)} организаций')

        inns[row].insert(13, None)
        [inns[row].insert(16, None) for _ in range(2)]

        inn = str(int(inns[row][4]))
        name = inns[row][3]
        print(f'Собираются данные о компании: {name}\nИНН: {inn}')

        with BaseOptions(headless).create_driver() as browser:
            browser.get('https://pb.nalog.ru/index.html')
            inn, company = _chek(inn, name)
            browser.implicitly_wait(3)
            browser.find_element(By.ID, 'queryAll').send_keys(inn)
            ActionChains(browser).send_keys(Keys.ENTER).perform()
            time.sleep(6)
            browser.implicitly_wait(5)
            new_row = inns[row].copy()

            try:
                if company == 'Человек':
                    href = browser.find_element(By.CLASS_NAME, 'pb-card.pb-card--clickable')
                    ActionChains(browser).move_to_element(href).click().perform()
                    time.sleep(6)

                try:
                    href = 'https://pb.nalog.ru/' + browser.find_element(By.CLASS_NAME,
                                                                         'pb-card.pb-card--clickable').get_attribute(
                        'data-href')
                except WebDriverException:
                    browser.refresh()
                    time.sleep(6)
                    try:
                        href = 'https://pb.nalog.ru/' + browser.find_element(By.CLASS_NAME,
                                                                             'pb-card.pb-card--clickable').get_attribute(
                            'data-href')
                    except WebDriverException:
                        worksheet.writer(inns[row])
                        continue
                browser.get(href)
                time.sleep(2)
                while True:
                    if browser.find_elements(By.ID, 'lnkBackToSearch'):
                        break
                    time.sleep(1)
                time.sleep(1)
                browser.implicitly_wait(5)
                if company == 'ИП':
                    sost_org = _sost_org(browser)
                    new_row[17] = sost_org
                    okved = _okved(browser, company)
                    new_row[18] = okved
                    msp = _msp(browser)
                    new_row[16] = msp
                    worksheet.writer(new_row)
                    time.sleep(sleep)

                    continue
                msp = _msp(browser)
                new_row[16] = msp
                okved = _okved(browser)
                new_row[18] = okved
                sost_org = _sost_org(browser)
                new_row[17] = sost_org
                nalog2021, nalog2020 = _nalog(browser)
                new_row[12] = nalog2021
                new_row[13] = nalog2020
                income2021, income2020 = _income(browser)
                new_row[10] = income2021
                new_row[11] = income2021

                worksheet.writer(new_row)
            except (WebDriverException, IndexError, AttributeError, ValueError) as ex:
                worksheet.writer(new_row)
                continue
            time.sleep(sleep)


def _sost_org(browser: WebDriver) -> str:
    sost_org = None
    try:
        sost_org = browser.find_element(By.CLASS_NAME, 'pb-subject-status').text.strip()
    except (NoSuchElementException, AttributeError):
        pass
    return sost_org


def _okved(browser: WebDriver, company=None) -> str:
    okved = None
    try:
        for i in browser.find_elements(By.CLASS_NAME, 'lnk-appeal'):
            if i.get_attribute('data-appeal-kind') == 'EGRUL_OKVED' or 'EGRIP_OKVED' == i.get_attribute('data-appeal-kind'):
                okved = i.text
    except (NoSuchElementException, AttributeError):
        pass
    return okved


def _msp(browser: WebDriver) -> str:
    msp = None
    try:
        for el in browser.find_elements(By.CLASS_NAME, 'pb-company-field'):
            if 'МСП' in el.text:
                msp = el.find_element(By.CLASS_NAME, 'pb-company-field-value').text
                break
    except (NoSuchElementException, AttributeError):
        msp = None
    return msp


def _income(browser: WebDriver) -> tuple[float | None, float | None]:
    income2021 = income2020 = None
    if 'Суммы доходов и расходов по данным бухгалтерской отчетности организации' in browser.page_source:
        try:
            incoms = {'2021': None, '2020': None}
            elem = browser.find_elements(By.CLASS_NAME, 'lnk-forward.lnk-detail')[-1]
            actions = ActionChains(browser)
            actions.move_to_element(elem)
            actions.click(elem)
            actions.perform()
            time.sleep(3)
            for tr in browser.find_element(By.ID, 'modalCompanyTbody').find_elements(By.TAG_NAME, 'tr'):
                td_elements = [td.text.strip() for td in tr.find_elements(By.TAG_NAME, 'td')]
                if '2021' in td_elements[0]:
                    incoms['2021'] = td_elements[1].replace(' ', '')
                elif '2020' in td_elements[0]:
                    incoms['2020'] = td_elements[1].replace(' ', '')

            income2021 = float(incoms['2021']) if incoms['2021'] else incoms['2021']
            income2020 = float(incoms['2020']) if incoms['2020'] else incoms['2020']
            browser.find_element(By.CLASS_NAME, 'close').click()
        except (IndexError, ValueError, WebDriverException):
            pass
    return income2021, income2020


def _nalog(browser: WebDriver) -> tuple[str | int, str | int]:
    nalog2021 = nalog2020 = 0
    if 'Уплаченные налоги и сборы' in browser.page_source:
        try:
            tax = {'2021': 0, '2020': 0}
            if 'Суммы доходов и расходов по данным бухгалтерской отчетности организации' in browser.page_source:
                elem = browser.find_elements(By.CLASS_NAME, 'lnk-forward.lnk-detail')[-2]
            else:
                elem = browser.find_elements(By.CLASS_NAME, 'lnk-forward.lnk-detail')[-1]
            actions = ActionChains(browser)
            actions.move_to_element(elem)
            actions.click(elem)
            actions.perform()
            time.sleep(3)
            years_elements = [i
                              for i in browser.find_element(By.CLASS_NAME,
                                                            'gamlet-period-selector-container-year').find_elements(
                    By.TAG_NAME, 'li')
                              if i.text.strip() in ['2021', '2020']]
            for year in years_elements:
                year_name = year.text.strip()
                year.click()
                for tr in browser.find_element(By.ID, 'modalCompanyTbody').find_elements(By.TAG_NAME, 'tr'):
                    n = tr.find_elements(By.TAG_NAME, 'td')[1].text.replace(' ', '')
                    if n and year_name == tr.get_attribute('data-year-code').strip():
                        tax[year_name] += float(n)
            nalog2021 = tax['2021']
            nalog2020 = tax['2020']
            browser.find_element(By.CLASS_NAME, 'close').click()
        except (IndexError, ValueError, WebDriverException):
            pass
    return nalog2021, nalog2020


def _chek(inn, name) -> tuple[str, str]:
    if re.fullmatch(r'[А-Яа-яA-Za-zЁё]+? [А-Яа-яA-Za-zЁё]+? [А-Яа-яA-Za-zЁё]+?', name):
        company = 'Человек'
        if len(inn) < 12:
            while len(inn) != 12:
                inn = '0' + inn
    elif re.match(r'ИП ', name):
        company = 'ИП'
        if len(inn) < 12:
            while len(inn) != 12:
                inn = '0' + inn
    else:
        company = 'Компания'
        if len(inn) < 10:
            while len(inn) != 10:
                inn = '0' + inn

    return inn, company
