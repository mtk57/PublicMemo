
import traceback

from logging_test import Log
from logging_test2 import LogTest2
from logging_test3 import LogTest3

if __name__ == '__main__':
    try:
        myLog = Log('my.log')
        myLog.log.info('caller')

        test2 = LogTest2()
        test2.func()
        test3 = LogTest3()
        test3.func()
    except:
        print(traceback.format_exc())