import time
from handlers.main_parser import main

if __name__ == '__main__':
    try:
        main()
    except Exception as ex:
        print(ex)
        time.sleep(50000)
    else:
        print('Сбор данных завершен')
        while True: time.sleep(50000)

