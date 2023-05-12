import openpyxl
import configparser


def _create_file(config):
    road = config.get("program", "create_file")
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.append(['Тип', 'regNumber', 'Регистрационный номер поставщика', 'Полное название', 'ИНН',
                      'Телефон в контрактах, число раз',
                      'Телефон в контрактах', 'Телефон', 'Комментарий', 'КПП', 'Доход 2021', 'Доход 2022',
                      'Налог 2021', 'Налог 2020', 'ОГРН', 'largestTaxpayerKpp',
                      'МСП', 'Состояние организации', 'okved', 'supplierOkved', 'Адрес', 'Почтовый адрес',
                      'Дата регистрации (юрлица?)',
                      'Дата регистрации поставщика', 'isSMP', 'Клиент', 'Поставщик',
                      'Эксклюзивный поставщик', 'Уровень подчинения', 'Описание уровня подчинения',
                      'Недобросовестный ФЗ44 поставщик',
                      'Недобросовестный ФЗ233 поставщик', 'isSono', 'Самозанятый', 'operatorEdo', 'abonentId',
                      'Директор', 'email', None, 'Факс',
                      'Сайт', 'Запрещено подавать заявки', 'Причина запрета подачи заявок', 'Запрещено делать закупки',
                      'Причина запрета закупок',
                      'id', None, None, None, None])
    worksheet.auto_filter.ref = 'A1:AS1'
    workbook.save(road)


def _get_rows(config):
    road = config.get("program", "read_file")
    rows = []
    workbook = openpyxl.load_workbook(road)

    worksheet = workbook.active

    for row in worksheet.iter_rows():
        rows.append([i.value for i in row])
    return rows


def _writer(row):
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')
    road = config.get("program", "create_file")
    workbook = openpyxl.load_workbook(road)

    worksheet = workbook.active

    worksheet.append(row)

    # Сохраняем файл
    workbook.save(road)
