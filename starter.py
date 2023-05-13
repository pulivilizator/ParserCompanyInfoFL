import time
from main import main

if __name__ == '__main__':
    try:
        main()
    except Exception as ex:
        print(ex)
        time.sleep(50000)

