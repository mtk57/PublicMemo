#!/usr/bin/python3

import os
import sys

from utils import UUID_to_PrefixUUID, \
                  get_uuid, \
                  get_all_file_dir, \
                  make_dir

"""
S3(ObjectStorage)向けのオブジェクトファイルを作成する

<使い方>
python3 s3_object_creator.py <対象ディレクトリの絶対パス> <出力先ディレクトリの絶対パス>
"""

# 対象ディレクトリ（ルート）のUUID
ROOT_DIR_UUID = '00000000-0000-0000-0000-000000000001'

# 種別(レギュラーファイル)
TYPE_REGULAR_FILE = 4

# 種別(ディレクトリ)
TYPE_DIRECTORY = 10


def create_object_file(
    output_dir_path: str,
    target: str,
    uuid_map: dict,
    uuid: str,
    is_dir=True
) -> dict:
    """
    オブジェクトファイルを作成する

    <ディレクトリの場合>
      ディレクトリエントリリストファイルを作成
       ファイル名：UUID
       ファイル内容：
         key=Prefix UUID
         body=対象ディレクトリ直下のファイル・ディレクトリの情報
            ファイルの場合    ：UUID + "/" + ファイル名 + "/" + "4"
            ディレクトリの場合：UUID + "/" + ディレクトリ名 + "/" + "10"

    <ファイルの場合>
       ファイル名：UUID
       ファイル内容：
         key=Prefix UUID
         body=ファイル名
    """

    ret_dict = uuid_map

    # ディレクトリのUUIDをマップに登録
    if uuid not in ret_dict:
        ret_dict[uuid] = target

    # UUIDでファイルを作成
    # TODO：本来であればPrefix UUIDでファイル名を作成すべきだが"/"があるので不可。
    parent_file = os.path.join(output_dir_path, uuid)
    with open(parent_file, mode='a') as f:

        f.write(f'key={UUID_to_PrefixUUID(uuid)}{os.linesep}')
        f.write(f'body={os.linesep}')

        body = ""

        if is_dir is True:
            # ディレクトリの場合
            # 親ディレクトリ直下のファイル・ディレクトリを取得
            sub = get_all_file_dir(target, sub_dir=False)
            for item in sub:
                # UUIDをマップに登録
                sub_uuid = get_uuid()
                ret_dict[sub_uuid] = item.rstrip(os.sep)

                name = os.path.basename(item)
                type = TYPE_REGULAR_FILE
                if os.path.isdir(item):
                    type = TYPE_DIRECTORY

                body = body + f'{sub_uuid}/{name}/{type}{os.linesep}'
        else:
            body = f'{os.path.basename(target)}'

        f.write(body)

    return ret_dict


def get_uuid_from_map(target: str, uuid_map: dict) -> str:
    """
    UUIDマップに対象が登録済みか調べる
    未登録の場合は新たなUUIDを生成して返す。
    登録済のみ場合は登録されているUUIDを返す
    """
    for key, value in uuid_map.items():
        if value == target:
            return key
    return get_uuid()


def main(target_dir_path, output_dir_path) -> int:

    # 対象ディレクトリの全ファイル・ディレクトリを取得
    targets = get_all_file_dir(target_dir_path)

    if len(targets) == 0:
        print("The target object does not exist.")
        return 0

    # UUIDとファイル・ディレクトリのマッピング
    uuid_map = {}

    # 全ファイル数分ループ
    for target in targets:

        target = target.rstrip(os.sep)

        if (target == target_dir_path):
            # オブジェクトファイルを作成（ルート用）
            uuid_map = create_object_file(
                        output_dir_path=output_dir_path,
                        target=target,
                        uuid_map=uuid_map,
                        uuid=ROOT_DIR_UUID
                       )
            continue

        # オブジェクトファイルを作成
        # 既に登録済の場合はマップからUUIDを取得
        uuid_map = create_object_file(
                    output_dir_path=output_dir_path,
                    target=target,
                    uuid_map=uuid_map,
                    uuid=get_uuid_from_map(target=target, uuid_map=uuid_map),
                    is_dir=os.path.isdir(target)
                   )

    return 0


if __name__ == '__main__':
    """
    第1引数：対象ディレクトリまでの絶対パス
             例：/home/kawa/test
                 C:\\kawa\\test
    第2引数：出力先ディレクトリまでの絶対パス
             例：/home/kawa/result
                 c:\\kawa\\result
    """

    AGRG_CNT = 3

    args = sys.argv

    # for TEST >>
    # args = []
    # args.append('dummy')
    # args.append(r"C:\_tmp\brick")
    # args.append(r"C:\_tmp\result")
    # for TEST <<

    if len(args) < AGRG_CNT:
        # パラメータ不足
        print("Parameters missing.")
        sys.exit(0)

    target_dir_path = args[AGRG_CNT-2].rstrip(os.sep)
    output_dir_path = args[AGRG_CNT-1].rstrip(os.sep)

    if target_dir_path == output_dir_path:
        # 同じディレクトリは指定不可
        print("You can't specify the same directory")
        sys.exit(0)

    if not os.path.exists(target_dir_path):
        # 対象ディレクトリが存在しない
        print("The target directory does not exist.")
        sys.exit(0)

    # 出力先ディレクトリを作成
    if make_dir(output_dir_path) is False:
        print("Output directory create failed.")
        sys.exit(0)

    # 処理開始
    ret = main(target_dir_path, output_dir_path)

    sys.exit(ret)
