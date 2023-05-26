from excel_writer import worksheet
from handlers.handlers import _parser
import configparser

def main():
    config = configparser.ConfigParser()
    config.read('config.ini', encoding='utf-8')
    if not int(config.get("program", "write_type")):
        worksheet.create_file()
    inns = worksheet.get_rows()
    print(f'Найдено {len(inns)} организаций\n'
          f'Начинаю собирать информацию.\n')
    _parser(inns, config)