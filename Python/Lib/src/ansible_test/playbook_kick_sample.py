#!/usr/bin/python3
import os
import sys
import subprocess

from utils import is_json_format, \
                  get_ip_address
from log import Log


IS_DEBUG = True
_log = Log('/var/log/playbook_kick_sample.log')


def main(
    inventry_file_path: str,
    playbook_file_path: str,
    param: str,
    is_call_back: bool = False
):
    _log.log.debug('main() S')

    _log.log.debug(f'inventry_file_path={inventry_file_path}')
    _log.log.debug(f'playbook_file_path={playbook_file_path}')
    _log.log.debug('param=' + param)
    _log.log.debug(f'is_call_back={is_call_back}')

    if is_call_back is False:
        # PlayBookをコール
        try:
            args = f'ansible-playbook ' \
                   f'-i {inventry_file_path} ' \
                   f'{playbook_file_path} ' \
                   f"-e '{param}'"

            _log.log.info(f'exec command={args}')

            result = subprocess.run(
                # 何故かlistにすると上手くいかないのでstrにする
                # [
                #     'ansible-playbook',
                #     f'-i {inventry_file_path}',
                #     playbook_file_path
                # ]
                args=args,

                # strの場合はshell=Trueが必要
                shell=True,

                # 失敗時はCalledProcessErrorをスロー
                check=True,

                # 標準出力を抑止
                stdout=subprocess.PIPE
            )

            for line in result.stdout.decode().splitlines():
                if 'RESULT' in line:
                    print(line)

        except subprocess.CalledProcessError as e:
            _log.log.error(f'ansible-playbook is failed.[{e}]')

    else:
        # PlayBookからコールバック
        if IS_DEBUG is True:
            _log.log.debug('CallBack process:')
        print(f'RESULT!  from {get_ip_address()}')

    _log.log.debug('main() E')


if __name__ == '__main__':
    HELP = """
playbook_kick_sample.py INVENTRY_FILE_PATH PLAYBOOK_FILE_PATH PARAM [--callback]

INVENTRY_FILE_PATH:インベントリファイルの絶対パス
        例：/root/test_playbook/hosts
PLAYBOOK_FILE_PATH:プレイブックファイルの絶対パス
        例：/root/test_playbook/site.yml
PARAM:プレイブックに渡すパラメータ(JSON形式の文字列)
        例：'{ "foo":"FOO", "fruits":["apple", "cherry", "orange"] }'
--callback:コールバックされた場合に付けるオプション
"""
    _log.log.debug('------------------------------------')
    _log.log.debug('START')
    _log.log.info(f'Host={get_ip_address()}')
    _log.log.info(f'args_len={sys.argv}, args={sys.argv}')

    MIN_ARGS = 4
    MAX_ARGS = MIN_ARGS + 1
    args = sys.argv

    # for TEST >>>
    if IS_DEBUG is True:
        if len(args) < MIN_ARGS:
            args = []
            args.append('dummy')
            BASE_DIR = r'/root/test_playbook'
            args.append(os.path.join(BASE_DIR, 'hosts'))
            args.append(os.path.join(BASE_DIR, 'site.yml'))
            args.append('{ "foo":"FOO", "fruits":["apple", "cherry", "orange"] }')

            _log.log.debug(f'new args_len={args}, args={args}')
    # for TEST <<<

    if len(args) <= 1:
        # ヘルプを表示
        print(HELP)
        sys.exit(0)

    if len(args) < MIN_ARGS:
        # 必須パラメータ不足
        msg = 'Parameters missing.'
        _log.log.error(f'END  [{msg}]')
        sys.exit(0)

    # パラメータ解析

    # Inventryファイルパス
    inventry_file_path = args[MIN_ARGS-3]
    if inventry_file_path != 'dummy':
        if not os.path.exists(inventry_file_path):
            # 対象が存在しない
            msg = f'The Inventry file does not exist.[{inventry_file_path}]'
            _log.log.error(f'END  [{msg}]')
            sys.exit(0)

    # PlayBookファイルパス
    playbook_file_path = args[MIN_ARGS-2]
    if playbook_file_path != 'dummy':
        if not os.path.exists(playbook_file_path):
            # 対象が存在しない
            msg = f'The PlayBook file does not exist.[{playbook_file_path}]'
            _log.log.error(f'END  [{msg}]')
            sys.exit(0)

    # PlayBookに渡すパラメータ(json)
    param = args[MIN_ARGS-1]
    if is_json_format(param) is False:
        # jsonではない
        msg = f'It is not a JSON format.[{param}]'
        _log.log.error(f'END  [{msg}]')
        sys.exit(0)

    # 任意オプションチェック
    is_call_back = False
    if len(args) >= MAX_ARGS:
        if args[MAX_ARGS-1] != r'--callback':
            # オプション名が不正
            msg = f'The option name is incorrect.[{args[MAX_ARGS-1]}]'
            _log.log.error(f'END  [{msg}]')
            sys.exit(0)
        else:
            is_call_back = True

    # 処理開始
    try:
        main(
            inventry_file_path=inventry_file_path,
            playbook_file_path=playbook_file_path,
            param=param,
            is_call_back=is_call_back
            )
    except Exception as ex:
        _log.log.error(ex)

    _log.log.debug('END')
