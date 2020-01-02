
import traceback
import time

from logging_test import _log

class LogTest2():
    def __init__(self):
        _log.log.info('LogTest2.init')
    def func(self):
        _log.log.info('LogTest2.func')

if __name__ == '__main__':
    try:
        test = LogTest2()
        test.func()
    except:
        print(traceback.format_exc())