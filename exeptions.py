class CreateException(Exception):
    '''Ошибка создания объекта вебдрайвера'''


class ShopExeption(Exception):
    '''Ошибка запроса получения списка магазинов'''


class AgentsException(Exception):
    '''Ошибка считывания User Agent'''


class GetFileExeption(Exception):
    '''Ошибка считывания файла на входе(файл не найден)'''


class WriteFileExeption(Exception):
    '''Ошибка записи в файл(файл не найден)'''


class CreateFileExeption(Exception):
    '''Ошибка создания файла'''
