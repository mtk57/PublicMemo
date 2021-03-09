import datetime
from logging import getLogger, FileHandler, Formatter, DEBUG, INFO, ERROR, WARN


class Logger():
    LOG_FILE_NAME = 'quizer.log'
    ENCODE = 'utf-8'
    FORMAT = '%(asctime)s:%(levelname)s:%(message)s'

    LEVEL_DEBUG = 'DEBUG'
    LEVEL_INFO = 'INFO'
    LEVEL_WARN = 'WARN'
    LEVEL_ERROR = 'ERROR'

    LEVEL_MAP = {
        LEVEL_DEBUG: DEBUG,
        LEVEL_INFO: INFO,
        LEVEL_WARN: WARN,
        LEVEL_ERROR: ERROR
    }

    def __init__(self, level: str = LEVEL_INFO,
                 log_file_name: str = LOG_FILE_NAME):
        try:
            self._level = Logger.LEVEL_MAP[level]
        except KeyError:
            self._level = Logger.LEVEL_MAP[Logger.LEVEL_INFO]

        self._log_file_name = log_file_name
        self._logger = self._get_logger()

    @classmethod
    def print_console(cls, message: str):
        print(f'{datetime.datetime.now()}:{message}')

    def DEBUG(self, message: str):
        self._logger.debug(message)

    def INFO(self, message: str):
        self._logger.info(message)
        Logger.print_console(message)

    def WARN(self, message: str):
        self._logger.warn(message)
        Logger.print_console(message)

    def ERROR(self, message: str):
        self._logger.error(message)
        Logger.print_console(message)

    def _get_logger(self) -> object:
        logger = getLogger(__name__)
        logger.setLevel(self._level)

        handler = FileHandler(self._log_file_name, encoding=Logger.ENCODE)
        handler.setLevel(self._level)

        handler.setFormatter(Formatter(Logger.FORMAT))
        logger.addHandler(handler)

        return logger
