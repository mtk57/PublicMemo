#!/usr/bin/python3

import os
import sys
import re

from utils import PrefixUUID_to_UUID, \
                  get_all_file_dir, \
                  check_format_UUID, \
                  check_format_Prefix_UUID, \
                  REG_UUID
from data_structure.tree.element import TreeElement

"""
s3_object_creator.py で出力したオブジェクトファイルからツリーを作成する

<使い方>
python3 s3_object_to_tree.py <出力先ディレクトリの絶対パス>
"""


# 対象ディレクトリ（ルート）のUUID
ROOT_DIR_UUID = '00000000-0000-0000-0000-000000000001'

# 種別(レギュラーファイル)
TYPE_REGULAR_FILE = 4

# 種別(ディレクトリ)
TYPE_DIRECTORY = 10


class FileModel():
    """
    ディレクトリオブジェクトのエントリリストの1行を
    表現したデータモデル
    """

    def __init__(self, key: str, name: str, file_type: int):
        self._key = key
        self._name = name
        self._file_type = file_type

    def __repr__(self):
        return f'_key={self._key}\n_name={self._name}\n' \
               f'_file_type={self._file_type}\n'

    @property
    def key(self) -> str:
        """ キー名(Prefix UUID) """
        return self._key

    @property
    def name(self) -> str:
        """ ファイル名 or ディレクトリ名 """
        return self._name

    @property
    def file_type(self) -> int:
        """ 種別 """
        return self._file_type


class ObjectFile():
    """
    オブジェクトファイルを表現したデータモデル
    """

    def __init__(self, contents: list):
        self._key = ''
        self._body = []
        self._read_contents(contents)

    def __repr__(self):
        return f'_key={self._key}\n_body={self._body}\n'

    def _read_contents(self, contents: list):
        """
        オブジェクトの内容を解析してメンバに設定する
        """

        # オブジェクトの全行をパース
        for line in contents:
            # 行末の改行コードは削除
            line = line.rstrip(os.linesep)

            if re.match(r'^key=*', line):
                # key=で開始していればKeyと見なす

                # デリミタで分割
                prefix_uuid = line.split('=')[1]
                # 念の為、Prefix UUIDのフォーマットチェック
                if check_format_Prefix_UUID(prefix_uuid):
                    # キーを取得
                    self._key = prefix_uuid

            elif re.match(r'^' + REG_UUID, line):
                # UUIDで開始していればBodyと見なす

                # デリミタで分割
                entry = line.split('/')

                if len(entry) > 2:
                    # 各フィールドを取得しファイルモデル化
                    key = entry[0]
                    name = entry[1]
                    file_type = int(entry[2])
                    model = FileModel(key, name, file_type)

                    # リストに追加
                    self._body.append(model)

            else:
                # 上記以外無視
                continue

    @property
    def key(self) -> str:
        return self._key

    @property
    def body(self) -> list:
        return self._body


def get_file_objects(targets: list) -> list:
    """
    オブジェクトファイルのリストからオブジェクトを作成して返す
    """

    objects = []
    for target in targets:
        basename = os.path.basename(target)

        # UUID以外は無視
        if check_format_UUID(basename) is False:
            continue

        # オブジェクト化
        with open(target, mode='r') as f:
            lines = f.readlines()
            obj = ObjectFile(contents=lines)
            objects.append(obj)
    return objects


def find_object(key: str, objects: list) -> ObjectFile:
    """
    オブジェクトリストから指定したキー(UUID)に
    一致するオブジェクトを探索する
    """

    for obj in objects:
        # オブジェクトリストのキーはPrefix UUIDなので
        # UUIDに変換する
        uuid = PrefixUUID_to_UUID(obj.key)
        if uuid == key:
            # 発見
            return obj
    return None


def create_node(node: TreeElement, obj: ObjectFile, objects: list):
    """
    ツリーのノードを作成する
    """

    for item in obj.body:
        if item.file_type == TYPE_REGULAR_FILE:
            # レギュラーファイル
            node.append(f'[{item.key}] | [{item.name}]')
        else:
            # 上記以外はディレクトリとする

            # ディレクトリなので子ノードを追加する
            child_node = node.append_child(f'[{item.key}] | [{item.name}]/')
            child_obj = find_object(item.key, objects)

            # 子ノードを作成するため再帰
            create_node(child_node, child_obj, objects)


def convert_object_to_tree(objects: list) -> TreeElement:
    """
    オブジェクトのリストからツリーに変換する
    """
    ret_tree = None

    # ルートを探す
    root_obj = None
    for obj in objects:
        uuid = PrefixUUID_to_UUID(obj.key)
        if uuid == ROOT_DIR_UUID:
            # ツリーのルートノードを作成
            ret_tree = TreeElement(id=uuid)
            root_obj = obj
            break

    if ret_tree is None:
        raise Exception('Root object is not exist!')

    # ツリーノード作成開始
    create_node(ret_tree, root_obj, objects)

    return ret_tree


def main(target_dir_path) -> int:

    # 対象ディレクトリの全ファイル・ディレクトリを取得
    targets = get_all_file_dir(target_dir_path)

    if len(targets) == 0:
        print("The target object does not exist.")
        return 0

    # まずはオブジェクト化
    objects = get_file_objects(targets)

    # オブジェクトのエントリリストをツリーノード化
    tree = convert_object_to_tree(objects)

    # ツリーを表示
    print(f'{tree.to_tree_string(show_id=True)}')

    return 0


if __name__ == '__main__':
    """
    第1引数：対象ディレクトリまでの絶対パス
             例：/home/kawa/result
                 c:\\kawa\\result
    """

    AGRG_CNT = 2

    args = sys.argv

    # for TEST >>>
    # args = []
    # args.append('dummy')
    # args.append(r"C:\_tmp\result")
    # for TEST <<<

    if len(args) < AGRG_CNT:
        # パラメータ不足
        print("Parameters missing.")
        sys.exit(0)

    target_dir_path = args[AGRG_CNT-1].rstrip(os.sep)

    if not os.path.exists(target_dir_path):
        # 対象ディレクトリが存在しない
        print("The target directory does not exist.")
        sys.exit(0)

    # 処理開始
    ret = main(target_dir_path)

    sys.exit(ret)
