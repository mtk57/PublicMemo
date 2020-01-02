
import traceback
import time

from logging_test import _log

class LogTest3():
    def __init__(self):
        _log.log.info('LogTest3.init')
    def func(self):
        _log.log.info('LogTest3.func')

if __name__ == '__main__':
    try:
        test = LogTest3()
        test.func()
    except:
        print(traceback.format_exc())