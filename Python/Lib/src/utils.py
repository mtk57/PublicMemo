import os
import uuid
import glob
import time
import shutil
import re
import datetime
import base64
import json
import subprocess
import socket

# UUIDを検索する正規表現
REG_UUID = r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}'

# Prefix UUIDを検索する正規表現
# Prefix UUIDとは
# UUIDの1,2文字目 + "/" + UUIDの3,4文字目 + "/" + UUID
REG_PREFIX_UUID = r'[0-9a-f]{2}/[0-9a-f]{2}/' + REG_UUID


def UUID_to_PrefixUUID(uuid: str) -> str:
    """ UUIDをPrefix UUIDに変換する """
    return f'{uuid[0:2]}/{uuid[2:4]}/{uuid}'


def PrefixUUID_to_UUID(prefix_uuid: str) -> str:
    """ Prefix UUIDをUUIDに変換する """
    return prefix_uuid[6:]


def get_uuid() -> str:
    """ UUIDを返す """
    return str(uuid.uuid4())


def get_all_file_dir(dir_path: str, sub_dir: bool = True) -> list:
    """
    指定されたディレクトリ配下のすべてのファイル、ディレクトリのパスをリストで返す
    サブディレクトリ内も再帰的に探索する
    @param dir_path 対象ディレクトリの絶対パス
    @param sub_dir  サブディレクトリを再帰的に探索するか否か
                    Falseの場合、対象ディレクトリ直下のリストのみを返す
    """
    wild_card = "**" if sub_dir is True else "*"
    return glob.glob(os.path.join(dir_path, wild_card), recursive=True)


def make_dir(path: str) -> bool:
    """
    ディレクトリを中間ディレクトリも含めて作成する
    作成前に全て削除する
    """

    # ディレクトリが既に存在していれば削除する
    if os.path.exists(path):
        shutil.rmtree(path)

    del_ok = False

    # リトライは3回まで
    for retry in range(3):
        try:
            # ディレクトリを作成
            os.makedirs(path, exist_ok=True)
            del_ok = True
        except PermissionError:
            time.sleep(1)

    return del_ok


_reg_uuid = re.compile(REG_UUID)
_reg_prefix_uuid = re.compile(REG_PREFIX_UUID)


def check_format_UUID(uuid: str) -> bool:
    """ UUIDのフォーマットチェック """
    return _reg_uuid.match(uuid) is not None


def check_format_Prefix_UUID(uuid: str) -> bool:
    """ Prefix UUIDのフォーマットチェック """
    return _reg_prefix_uuid.match(uuid) is not None


def get_datetime_string(time) -> str:
    FMT = '%Y-%m-%d %H:%M:%S.%f'
    return datetime.datetime.fromtimestamp(time).strftime(FMT)


def is_bit_on(val1, val2) -> bool:
    return val1 & val2 == val2


def to_base64_encode(bytes: bytes) -> str:
    return base64.b64encode(bytes)


def to_base64_decode(string: str) -> str:
    return base64.b64decode(string)


def is_json_format(json_str: str) -> bool:
    try:
        json.loads(json_str)
    except json.JSONDecodeError:
        return False
    else:
        return True


def read_json(path: str) -> dict:
    ret = {}
    if os.path.exists(path):
        try:
            json_str = open(path, 'r')
            if is_json_format(json_str):
                ret = json.load(json_str)
        except Exception:
            pass
    return ret


def get_ip_address() -> str:
    return socket.gethostbyname(socket.gethostname())