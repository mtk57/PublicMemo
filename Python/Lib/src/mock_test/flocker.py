#!/usr/bin/env python3

import fcntl


def flocker():
    import os
    print(os.getcwd())

    with open('sample_db.sqlite') as oLockFile:
        try:
            fcntl.flock(oLockFile.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except IOError:
            print('ロックを獲得できませんでした')
            return

        try:
            print('ロックを獲得できました')
            print('ロックを獲得できた時の処理をここに書きます')
            print('例えば、ここでは 300秒間スリープします')
            # import time
            # time.sleep(300)
        finally:
            fcntl.flock(oLockFile.fileno(), fcntl.LOCK_UN)
