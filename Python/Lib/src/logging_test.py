import logging
import time
import traceback

"""
ログのテスト
https://docs.python.org/ja/3/library/logging.html

同じログファイルに対して別プロセスから書き込んでも
正常に書き込めた。
"""


class Log():
    def __init__(self, logfile):
        # ロガーを作成
        #self.log = logging.getLogger(__name__)
        self.log = logging.getLogger(logfile)

        # レベルを設定
        self.log.setLevel(logging.DEBUG)

        # ハンドラを作成
        handler = logging.FileHandler(logfile)

        # 書式を作成
        fmt = "%(asctime)s %(process)d %(thread)d %(levelname)s %(name)s :%(message)s"
        formatter = logging.Formatter(fmt)

        # 書式を設定
        handler.setFormatter(formatter)

        # ハンドラを設定
        self.log.addHandler(handler)


LOG = 'logtest.log'
_log = None


def init_log(logfile):
    try:
        global _log
        if _log is None:
            _log = Log(logfile)
    except Exception:
        print(traceback.format_exc())


init_log(LOG)


if __name__ == '__main__':
    log = Log('test.log')
    try:
        # log.log.info('test')
        for i in range(100):
            log.log.info(time.time())
            time.sleep(1)
    except Exception:
        print(traceback.format_exc())