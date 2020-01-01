#!/usr/bin/env python

import traceback
from functools import wraps
from time import sleep, time

class RetryOverError(Exception):
    pass

def retry(count:int, delay:float, errs:tuple):
    """ リトライデコレーター
        count:int:回数
        delay:float:インターバル(秒)
        errs:tuple:捕捉するエラータイプ
    """

    def _retry(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            _count = _delay = 3
            _errs = (Exception,)

            if isinstance(count, int):
                _count = abs(count)
            if isinstance(delay, float):
                _delay = abs(delay)
            if isinstance(errs, tuple):
                _errs = errs
 
            for c in range(_count):
                try:
                    return func(*args, **kwargs)
                except errs as e:
                    sleep(_delay)

            raise RetryOverError('Retry over!')

        return wrapper
    return _retry

RETRY_CNT = 3
RETRY_DELAY = 1.5
RETRY_ERRORS = (Exception,)

try_count = 0


@retry(count=RETRY_CNT, delay=RETRY_DELAY, errs=RETRY_ERRORS)
def retry_test(threshold):
    try:
        global try_count
        try_count += 1
        print(f'try count: {try_count}')
    except:
        print(traceback.format_exc())   # リトライオーバー時はここには来ない
    else:
        if try_count > threshold:
            return True
        else:
            raise ValueError()

if __name__ == '__main__':
    try:
        print(f'retry_test is {"succeeded" if retry_test(2) else "failed"}')
        print()
        try_count = 0
        print(f'retry_test is {"succeeded" if retry_test(3) else "failed"}')
    except RetryOverError as roe:
        print(f'{roe.args[0]}')
    except:
        print(traceback.format_exc())
