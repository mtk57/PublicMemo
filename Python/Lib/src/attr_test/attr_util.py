#!/usr/bin/python3

import os
import sys
import re
from stat import S_ISDIR, S_ISCHR, S_ISBLK, S_ISREG, S_ISFIFO, S_ISLNK, \
                 S_ISSOCK, S_ISDOOR, S_ISPORT, S_ISWHT, S_IMODE, S_IFMT, \
                 S_ISUID, S_ISGID, S_ENFMT, S_ISVTX, S_IREAD, S_IWRITE, \
                 S_IEXEC, S_IRWXU, S_IRUSR, S_IWUSR, S_IXUSR, S_IRWXG, \
                 S_IRGRP, S_IWGRP, S_IXGRP, S_IRWXO, S_IROTH, S_IWOTH, \
                 S_IXOTH

from utils import get_datetime_string, \
                  is_bit_on, \
                  read_json

"""
Linuxの拡張属性の操作を行う

<扱う情報>
・inode
・拡張属性

[inode]
os.stat('abc.txt')        // https://docs.python.org/ja/3/library/stat.html
os.stat_result(
    st_mode=33188,        // inodeの保護モード

                             S_ISDIR(mode)  ディレクトリの場合：非零
                             S_ISCHR(mode)  キャラクタ型の特殊デバイスファイルの場合：非零
                             S_ISBLK(mode)  ブロック型の特殊デバイスファイルの場合：非零
                             S_ISREG(mode)  レギュラーファイルの場合：非零
                             S_ISFIFO(mode) FIFO (名前つきパイプ) の場合：非零
                             S_ISLNK(mode)  シンボリックリンクの場合：非零
                             S_ISSOCK(mode) ソケットの場合：非零
                             S_ISDOOR(mode) ドアの場合：非零
                             S_ISPORT(mode) イベントポートの場合：非零
                             S_ISWHT(mode)  ホワイトアウトの場合：非零
                             S_IMODE(mode)  os.chmod() で設定することのできる一部のファイル
                                            モード すなわち、ファイルの許可ビットに加え、
                                            スティッキービット、実行グループ ID 設定 および
                                            実行ユーザ ID 設定 ビットを返します。

                                            S_ISUID   UIDビット
                                            S_ISGID   GIDビット
                                            S_ENFMT   S_ISGID と共有
                                            S_ISVTX   スティッキービット
                                                      このビットがディレクトリに対して設定されているとき、
                                                      そのディレクトリ内のファイルは、そのファイルのオーナー、
                                                      あるいはそのディレクトリのオーナーか特権プロセスのみが、
                                                      リネームや削除をすることが出来ることを意味しています。
                                            S_IREAD   S_IRUSR のUnix V7 のシノニム。
                                            S_IWRITE  S_IWUSR のUnix V7 のシノニム。
                                            S_IEXEC   S_IXUSR のUnix V7 のシノニム。
                                            S_IRWXU   ファイルオーナーの権限に対するマスク。
                                            S_IRUSR   オーナーがリード権限を持っている。
                                            S_IWUSR   オーナーがライト権限を持っている。
                                            S_IXUSR   オーナーが実行権限を持っている。
                                            S_IRWXG   グループの権限に対するマスク。
                                            S_IRGRP   グループがリード権限を持っている。
                                            S_IWGRP   グループがライト権限を持っている。
                                            S_IXGRP   グループが実行権限を持っている。
                                            S_IRWXO   その他 (グループ外) の権限に対するマスク。
                                            S_IROTH   その他はリード権限を持っている。
                                            S_IWOTH   その他はライト権限を持っている。
                                            S_IXOTH   その他は実行権限を持っている。

                             S_IFMT(mode)   ファイルの形式を記述しているファイルモードの一部
                                            (上記の S_IS*() 関数で使われます) を返します。

    st_ino=1286145,       // inode番号
    st_dev=64768,         // inodeが存在するデバイス
    st_nlink=2,           // 該当するinodeへのリンク数
    st_uid=0,             // ファイルの所持者のユーザID
    st_gid=0,             // ファイルの所持者のグループID
    st_size=0,            // 通常ファイルではバイトサイズ
                          // いくつかの特殊ファイルでは処理待ちのデータ量
    st_atime=1587685838,  // 最後にアクセスした時刻
    st_mtime=1587685838,  // 最後に変更された時刻
    st_ctime=1587685838   // 最後にメタデータが更新された時間
)


[拡張属性]   https://docs.python.org/ja/3.5/library/os.html
os.getxattr
os.listxattr
os.removexattr
os.setxattr

拡張属性の値を16進数で表示するか、デコードして表示するかは外部ファイルに正規表現で定義している
拡張属性名は以下の名前空間に属している必要がある（名前空間をつけないとサポートされていない操作ですと怒られる）
system    カーネルが使用
security  主にSELinuxが使用
trusted   信頼されるプロセスが使用
user      ユーザー用（自由に定義して使用できる）

"""

# 対応コマンドと必須パラメータ数の辞書
CMD_STAT = 'stat'
CMD_GETXATTR = 'getxattr'
CMD_LISTXATTR = 'listxattr'
CMD_REMOVEXATTR = 'removexattr'
CMD_SETXATTR = 'setxattr'
CMD_ALLGET = 'allget'
CMD_REMOVE_FROM_FILE = 'remove_from_file'
CMD_SET_FROM_FILE = 'set_from_file'

# Short name
CMD_GETXATTR_S = 'get'
CMD_LISTXATTR_S = 'list'
CMD_REMOVEXATTR_S = 'remove'
CMD_SETXATTR_S = 'set'
CMD_REMOVE_FROM_FILE_S = 'removef'
CMD_SET_FROM_FILE_S = 'setf'

COMMANDS = {
    CMD_STAT: 3,
    CMD_GETXATTR: 4,
    CMD_GETXATTR_S: 4,
    CMD_LISTXATTR: 3,
    CMD_LISTXATTR_S: 3,
    CMD_REMOVEXATTR: 4,
    CMD_REMOVEXATTR_S: 4,
    CMD_SETXATTR: 5,
    CMD_SETXATTR_S: 5,
    CMD_ALLGET: 3,
    CMD_REMOVE_FROM_FILE: 3,
    CMD_REMOVE_FROM_FILE_S: 3,
    CMD_SET_FROM_FILE: 3,
    CMD_SET_FROM_FILE_S: 3
}

# 設定ファイル
conf = {}


def cmd_stat(target_path: str):
    """
    stat()    path で指定されたファイルの状態を取得して buf へ格納する。
    lstat()   stat ()と同じであるが、 path がシンボリックリンクの場合、リンクが参照しているファイルではなく、
              リンク自身の状態を取得する点が異なる。
    fstat()は stat ()と同じだが、 状態を取得するファイルをファイル・ディスクリプタ filedes で指定する。
    """
    # result = os.stat(target_path)
    result = os.lstat(target_path)
    # result = os.fstat(target_path)

    mode = result.st_mode
    imode = S_IMODE(mode)

    disp = f"""
<inode>  {target_path}
  st_mode  ={result.st_mode}
            S_ISDIR  ={S_ISDIR(mode)}
            S_ISCHR  ={S_ISCHR(mode)}
            S_ISBLK  ={S_ISBLK(mode)}
            S_ISREG  ={S_ISREG(mode)}
            S_ISFIFO ={S_ISFIFO(mode)}
            S_ISLNK  ={S_ISLNK(mode)}
            S_ISSOCK ={S_ISSOCK(mode)}
            S_ISDOOR ={S_ISDOOR(mode)}
            S_ISPORT ={S_ISPORT(mode)}
            S_ISWHT  ={S_ISWHT(mode)}
            S_IMODE  ={oct(S_IMODE(mode))}
                      [User or Group ID]
                      S_ISUID  ={is_bit_on(imode, S_ISUID)}
                      S_ISGID  ={is_bit_on(imode, S_ISGID)}

                      S_ENFMT  ={is_bit_on(imode, S_ENFMT)}

                      [Sticky bit]
                      S_ISVTX  ={is_bit_on(imode, S_ISVTX)}

                      S_IREAD  ={is_bit_on(imode, S_IREAD)}
                      S_IWRITE ={is_bit_on(imode, S_IWRITE)}
                      S_IEXEC  ={is_bit_on(imode, S_IEXEC)}

                      [Owner permission]
                      S_IRWXU  ={is_bit_on(imode, S_IRWXU)}
                      S_IRUSR  ={is_bit_on(imode, S_IRUSR)}
                      S_IWUSR  ={is_bit_on(imode, S_IWUSR)}
                      S_IXUSR  ={is_bit_on(imode, S_IXUSR)}

                      [Group permission]
                      S_IRWXG  ={is_bit_on(imode, S_IRWXG)}
                      S_IRGRP  ={is_bit_on(imode, S_IRGRP)}
                      S_IWGRP  ={is_bit_on(imode, S_IWGRP)}
                      S_IXGRP  ={is_bit_on(imode, S_IXGRP)}

                      [Other permission]
                      S_IRWXO  ={is_bit_on(imode, S_IRWXO)}
                      S_IROTH  ={is_bit_on(imode, S_IROTH)}
                      S_IWOTH  ={is_bit_on(imode, S_IWOTH)}
                      S_IXOTH  ={is_bit_on(imode, S_IXOTH)}
            S_IFMT   ={S_IFMT(mode)}
  st_ino   ={result.st_ino}
  st_dev   ={result.st_dev}
  st_nlink ={result.st_nlink}
  st_uid   ={result.st_uid}
  st_gid   ={result.st_gid}
  st_size  ={result.st_size}
  st_atime ={get_datetime_string(result.st_atime)}
  st_mtime ={get_datetime_string(result.st_mtime)}
  st_ctime ={get_datetime_string(result.st_ctime)}
"""

    print(disp)


def is_exist_attr_name(target_path: str, attr_name: str) -> bool:
    """ 拡張属性名が存在しているか否か """
    attr_list = cmd_listxattr(target_path, is_show=False)
    for attr in attr_list:
        if attr == attr_name:
            return True
    return False


def cmd_getxattr(target_path: str, attr_name: str, is_show: bool = True):
    """ 拡張属性名から値を取得する """
    result = os.getxattr(target_path, attr_name)
    # encode = to_base64_encode(result)
    if is_show:
        for hex in conf[CMD_GETXATTR]['to_hex']:
            if re.match(hex, attr_name):
                result = result.hex()
                break
        else:
            for decode in conf[CMD_GETXATTR]['to_decode']:
                if re.match(decode, attr_name):
                    result = result.decode()
                    break

        print(f'{attr_name}={result}')
    return result


def cmd_listxattr(target_path: str, is_show: bool = True) -> list:
    """ 拡張属性名を列挙する """
    result = os.listxattr(target_path)
    if is_show:
        print(f'<listxattr>  {target_path}')
        for attr in result:
            print(f'{attr}')
    return result


def cmd_removexattr(target_path: str, attr_name: str):
    """ 拡張属性名と値を削除する """
    os.removexattr(target_path, attr_name)


def cmd_setxattr(target_path: str, attr_name: str, attr_value: str):
    """ 拡張属性名と値を設定する """
    os.setxattr(target_path, attr_name, attr_value.encode('utf-8'))


def main(
    target_path: str,
    command: str,
    attr_name: str = None,
    attr_value: str = None
):
    # jsonを読み込む(無くてもエラーとはしない)
    global conf
    conf = read_json('attr_util.json')

    # コマンド毎に分岐（ダサい）
    if command == CMD_STAT:
        cmd_stat(target_path)
    elif command == CMD_GETXATTR or command == CMD_GETXATTR_S:
        if is_exist_attr_name(target_path, attr_name):
            cmd_getxattr(target_path, attr_name)
        else:
            print(f'Not exist attribute name.[{attr_name}]')
    elif command == CMD_LISTXATTR or command == CMD_LISTXATTR_S:
        cmd_listxattr(target_path)
    elif command == CMD_REMOVEXATTR or command == CMD_REMOVEXATTR_S:
        if is_exist_attr_name(target_path, attr_name):
            cmd_removexattr(target_path, attr_name)
        else:
            print(f'Not exist attribute name.[{attr_name}]')
    elif command == CMD_SETXATTR or command == CMD_SETXATTR_S:
        if is_exist_attr_name(target_path, attr_name):
            print(f'Already exist attribute name.[{attr_name}]')
        else:
            cmd_setxattr(target_path, attr_name, attr_value)
    elif command == CMD_ALLGET:
        cmd_stat(target_path)
        attr_list = cmd_listxattr(target_path, is_show=False)
        print('<attributes>')
        for attr in attr_list:
            cmd_getxattr(target_path, attr)
    elif command == CMD_REMOVE_FROM_FILE or command == CMD_REMOVE_FROM_FILE_S:
        for key in conf['from_file'].keys():
            if is_exist_attr_name(target_path, key):
                cmd_removexattr(target_path, key)
    elif command == CMD_SET_FROM_FILE or command == CMD_SET_FROM_FILE_S:
        for key, value in conf['from_file'].items():
            if is_exist_attr_name(target_path, key):
                cmd_removexattr(target_path, key)
            cmd_setxattr(target_path, key, value)


if __name__ == '__main__':

    HELP = """
ファイルまたはディレクトリの詳細情報(inode, 拡張属性)の操作を行う

attr_util.py PATH COMMAND [ATTR NAME] [ATTR VALUE]

  PATH：対象の絶対パス
        例：/home/yuko/test/abc.txt   ※ファイルの場合
            /home/yuko/test/          ※ディレクトリの場合

  COMMAND：コマンド
           stat
           getxattr | get
           listxattr | list
           removexattr | remove
           setxattr | set
           allget                       ※stat + listxattr + getxattr
           remove_from_file | removef   ※外部ファイルを使用
           set_from_file | setf         ※外部ファイルを使用

  ATTR NAME：属性名 (以下コマンドのみ必須)
             getxattr | get
             removexattr | remove
             setxattr | set

  ATTR VALUE：値  (以下コマンドのみ必須)
              setxattr | set
"""

    MIN_ARGS = 3

    args = sys.argv

    if len(args) <= 1:
        # ヘルプを表示
        print(HELP)
        sys.exit(0)

    if len(args) < MIN_ARGS:
        # パラメータ不足
        print("Parameters missing.")
        sys.exit(0)

    # パラメータ解析

    # 対象パス
    target_path = args[1]
    if not os.path.exists(target_path):
        # 対象が存在しない
        print("The target does not exist.")
        sys.exit(0)

    # コマンド
    command = args[2]
    for cmd in COMMANDS.keys():
        if cmd == command:
            break
    else:
        # 未知のコマンド
        print("It's an unknown command.")
        sys.exit(0)

    # コマンド毎の必須パラメータチェック
    min_args_cnt = COMMANDS[command]
    if len(args) < min_args_cnt:
        # パラメータ不足
        print("Parameters missing.")
        sys.exit(0)

    # 属性名と属性値
    attr_name = None
    attr_value = None

    if len(args) == MIN_ARGS + 1:
        attr_name = args[MIN_ARGS]
    if len(args) == MIN_ARGS + 2:
        attr_value = args[MIN_ARGS + 1]

    # 処理開始
    main(
            target_path=target_path,
            command=command,
            attr_name=attr_name,
            attr_value=attr_value
        )
