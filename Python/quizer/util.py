import datetime
import os
import traceback
import random
import shutil

import openpyxl
from openpyxl import utils

import const as const


class Util():
    @classmethod
    def change_current_dir(cls, path: str = __file__):
        os.chdir(os.path.dirname(path))

    @classmethod
    def read_excel(cls, path: str) -> dict:
        return openpyxl.load_workbook(path, read_only=True, data_only=True)

    @classmethod
    def get_datetime_from_YYYYMMDD(cls, yyyymmdd: str) -> datetime:
        return datetime.datetime.strptime(yyyymmdd, r'%Y/%m/%d')

    @classmethod
    def get_datetime_from_general_values(cls, value: int) -> datetime:
        return utils.datetime.from_excel(value)

    @classmethod
    def get_hhmmss_str_from_sec(cls, seconds) -> str:
        return str(datetime.timedelta(seconds=seconds))

    @classmethod
    def get_exception_message(cls, ex: Exception) -> str:
        return f'Exception!! ex=[{ex}], trace=[{traceback.format_exc()}]'

    @classmethod
    def get_random_list(cls, src_list: list) -> list:
        return random.sample(src_list, len(src_list))

    @classmethod
    def create_tmp_file(cls, src_file_path: str) -> str:
        """
        引数に渡されたファイルをコピーして一時ファイルを作成する
        作成されたファイルのパスを返す
        """
        dst_file_path = os.path.join(os.path.dirname(src_file_path), const.TMP_FILE_NAME)
        if os.path.exists(dst_file_path):
            os.remove(dst_file_path)
        shutil.copy(src_file_path, dst_file_path)
        return dst_file_path
