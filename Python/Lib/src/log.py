import logging


class Log():
    def __init__(self, logfile):
        # ロガーを作成
        self.log = logging.getLogger(logfile)

        # レベルを設定
        self.log.setLevel(logging.DEBUG)

        # ハンドラを作成
        handler = logging.FileHandler(logfile)

        # 書式を作成
        fmt = "%(asctime)s:%(message)s"
        formatter = logging.Formatter(fmt)

        # 書式を設定
        handler.setFormatter(formatter)

        # ハンドラを設定
        self.log.addHandler(handler)
