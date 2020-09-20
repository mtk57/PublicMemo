#!/usr/bin/env python3
import time
import os
import sys
"""
https://qiita.com/croquisdukke/items/9c5d8933496ba6729c78

注意：このスクリプトはpython_daemon_test@.serviceから起動されることを想定している。
      Ex. systemctl start python_daemon_test@.service
      スクリプト単体で実行するとゾンビプロセス(親がすぐ死ぬので)となる。
"""


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
        main_unit()
    else:
        # メモリ不足でfork失敗
        sys.exit(-1)


if __name__ == '__main__':
    daemonize()
