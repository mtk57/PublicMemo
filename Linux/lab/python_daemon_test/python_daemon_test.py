#!/usr/bin/env python3
import time
import os
import sys
import signal
"""
https://qiita.com/croquisdukke/items/9c5d8933496ba6729c78

注意：このスクリプトはpython_daemon_test@.serviceから起動されることを想定している。
      Ex. systemctl start python_daemon_test@.service
      スクリプト単体で実行するとゾンビプロセス(親がすぐ死ぬので)となる。
"""


def handler(signum, frame):
    # signal 受信時の処理

    print(f'signal num=[{signum}]')

    if signum == signal.SIGTERM:
        print(f'SIGTERM received.')
    elif signum == signal.SIGINT:
        print(f'SIGINT received.')
    elif signum == signal.SIGKILL:
        print(f'SIGKILL received.')
    elif signum == signal.SIGHUP:
        print(f'SIGHUP received.')
    else:
        print(f'SIGxxx received.')

    sys.exit(0)


def main_unit():
    """
    10秒おきに時刻を書き込む
    """

    args = sys.argv
    instance = ''
    if len(args) > 1:
        instance = args[1]
    while True:
        filepath = '/opt/python_daemon_test.log'
        with open(filepath, 'a') as log_file:
            log_file.write(f'{time.ctime()}, {instance}\n')
        time.sleep(10)


def daemonize():
    pid = os.fork()

    print(f'pid={pid}')

    if pid > 0:
        # 親プロセスの場合(pidは子プロセスのプロセスID)
        with open('/var/run/python_daemon_test.pid', 'w') as pid_file:
            pid_file.write(str(pid)+"\n")
        sys.exit()
    elif pid == 0:
        # 子プロセスの場合

        # シグナルをハンドリング
        signal.signal(signal.SIGTERM, handler)
        signal.signal(signal.SIGINT, handler)
        # signal.signal(signal.SIGKILL, handler)  # KILLは捕捉不可なのでエラーとなる
        signal.signal(signal.SIGHUP, handler)

        main_unit()
    else:
        # メモリ不足でfork失敗
        sys.exit(-1)


if __name__ == '__main__':
    daemonize()
