import time
from webdriver_project import BaseOptions
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException, WebDriverException
from get_and_write_xlsx import _get_rows, _writer, _create_file
import configparser
import re


def _parser(inns, config):
    sleep = int(config.get("program", "sleep"))
    if sleep < 18:
        sleep = 0
    else:
        sleep -= 18
    headless = bool(int(config.get("program", "headless")))
    for row in range(101, len(inns)):
        print(f'Пройдено {row} из {len(inns)} организаций')
        inns[row].insert(13, None)
        inns[row].insert(16, None)
        inns[row].insert(16, None)
        inns[row].insert(16, None)
        inn = str(inns[row][4])
        name = inns[row][3]
        print(f'Собираются данные о компании: {name}\nИНН: {inn}')
        with BaseOptions(headless).create_driver() as browser:
            browser.get('https://pb.nalog.ru/index.html')
            inn, company = _chek(inn, name)
            browser.find_element(By.ID, 'queryAll').send_keys(inn)
            ActionChains(browser).send_keys(Keys.ENTER).perform()
            time.sleep(6)
            browser.implicitly_wait(5)

            try:
                if company == 'Человек':
                    okved = None
                    href = browser.find_element(By.CLASS_NAME, 'mass-group')
                    ActionChains(browser).move_to_element(href).click().perform()
                    time.sleep(6)

                try:
                    href = browser.find_element(By.CLASS_NAME, 'lnk.company-info').get_attribute('href')
                except WebDriverException:
                    browser.refresh()
                    time.sleep(6)
                    try:
                        href = browser.find_element(By.CLASS_NAME, 'lnk.company-info').get_attribute('href')
                    except WebDriverException:
                        _writer(inns[row])
                        continue
                browser.get(href)
                time.sleep(6)
                browser.implicitly_wait(5)

                if company == 'ИП':
                    try:
                        okved = browser.find_elements(By.CLASS_NAME, 'field.row.row__stretch')[-1].find_element(
                            By.CLASS_NAME, 'lnk-appeal').text
                    except NoSuchElementException:
                        okved = None

                    try:
                        msp = browser.find_element(By.CLASS_NAME, 'has-stickers').find_element(By.CLASS_NAME,
                                                                                               'has-stickers').text.split(
                            ':')[1].replace('"', '').strip()
                    except NoSuchElementException:
                        msp = None

                    new_row = inns[row].copy()
                    new_row[16] = msp
                    new_row[18] = okved
                    _writer(new_row)
                    time.sleep(sleep)
                    continue

                sost_org = _sost_org(browser)

                okved = _okved(browser)

                nalog2021, nalog2020 = _nalog(browser)

                income2022, income2021 = _income(browser)

                msp = _msp(browser)

                new_row = inns[row].copy()
                new_row[10] = income2021
                new_row[11] = income2022
                new_row[12] = nalog2021
                new_row[13] = nalog2020
                new_row[16] = msp
                new_row[17] = sost_org
                new_row[18] = okved
                _writer(new_row)
            except NoSuchElementException:
                _writer(inns[row])
                continue
            time.sleep(sleep)


def _sost_org(browser: WebDriver) -> str:
    sost_org = None
    try:
        for i in browser.find_elements(By.CLASS_NAME, 'field-group-name'):
            if 'Состояние организации:' in i.text:
                sost_org = i.text.split(':')[1].replace('"', '').strip()
                break
    except AttributeError:
        pass
    return sost_org


def _okved(browser: WebDriver) -> str:
    okved = None
    try:
        for i in browser.find_elements(By.CLASS_NAME, 'field.row.row__stretch'):
            if i.get_attribute('data-group') == 'okved':
                okved = i.find_element(By.CLASS_NAME, 'lnk-appeal').text
    except AttributeError:
        pass
    return okved


def _msp(browser: WebDriver) -> str:
    try:
        msp = browser.find_element(By.CLASS_NAME, 'has-stickers').find_element(By.CLASS_NAME,
                                                                               'has-stickers').text.split(':')[
            1].replace('"', '').strip()
    except (NoSuchElementException, AttributeError):
        msp = None
    return msp


def _income(browser: WebDriver) -> tuple[float | None, float | None]:
    income2022 = income2021 = None
    if 'Суммы доходов и расходов по данным бухгалтерской отчетности организации:' in browser.page_source:
        elem = browser.find_elements(By.CLASS_NAME, 'toggle')[-1]
        ActionChains(browser).move_to_element(elem).click().perform()
        for i in browser.find_elements(By.CLASS_NAME, 'field-group'):
            if i.get_attribute('data-group') == 'form1':
                elems = i.find_element(By.CLASS_NAME, 'wide').find_elements(By.TAG_NAME, 'td')
                break

        try:
            elems = [i.text for i in elems]
        except Exception as ex:
            print(ex, 'ELEMS')
            print(elems)
        for i in range(len(elems)):
            if elems[i] == '2022':
                income2022 = elems[i + 1].replace(' ', '').strip()
            if elems[i] == '2021':
                income2021 = elems[i + 1].replace(' ', '').strip()
    try:
        if income2022.replace('.', '', 1).isdigit():
            income2022 = float(income2022)
        if income2021.replace('.', '', 1).isdigit():
            income2021 = float(income2021)
    except (ValueError, AttributeError):
        pass
    return income2022, income2021


def _nalog(browser: WebDriver) -> tuple[str | int, str | int]:
    nalog2021 = nalog2020 = 0
    if 'Уплаченные налоги и сборы (без учета сумм налогов (сборов)' in browser.page_source:
        nalogs = \
            [i for i in browser.find_elements(By.CLASS_NAME, 'field-group') if i.get_attribute('data-group') == 'taxpay'][0] \
                .find_element(By.CLASS_NAME, 'gamlet-period-selector-container-year').find_elements(By.TAG_NAME, 'li')
        click2021 = [i for i in nalogs if i.get_attribute('data-value') == '2021'][0]
        click2020 = [i for i in nalogs if i.get_attribute('data-value') == '2020'][0]
        if 'Суммы доходов и расходов по данным бухгалтерской отчетности организации:' in browser.page_source:
            elem = browser.find_elements(By.CLASS_NAME, 'toggle')[-2]
        else:
            elem = browser.find_elements(By.CLASS_NAME, 'toggle')[-1]
        ActionChains(browser).move_to_element(elem).click().perform()

        for i in browser.find_element(By.ID, 'tbodyTaxpay').find_elements(By.TAG_NAME, 'tr'):
            try:
                if i.get_attribute('data-year-code') == '2021':
                    click2021.click()
                    nalog2021 += float(i.find_elements(By.TAG_NAME, 'td')[1].text.replace(' ', ''))
            except IndexError:
                nalog2021 = 0
            try:
                if i.get_attribute('data-year-code') == '2020':
                    click2020.click()
                    nalog2020 += float(i.find_elements(By.TAG_NAME, 'td')[1].text.replace(' ', ''))
            except IndexError:
                nalog2020 = 0
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


def main():
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')
    _create_file(config)
    inns = _get_rows(config)
    print(f'Найдено {len(inns)} организаций\n'
          f'Начинаю собирать информацию.\n')
    _parser(inns, config)


if __name__ == '__main__':
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')
    _create_file(config)
    inns = _get_rows(config)
    print(f'Найдено {len(inns)} организаций\n'
          f'Начинаю собирать информацию.\n')
    _parser(inns)
