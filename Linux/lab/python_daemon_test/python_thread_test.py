#!/usr/bin/env python3
import os
import sys
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed
from queue import Queue, deque

"""
ThreadPoolExecutorを使って、サービスインスタンスを開始する実験
注意：このスクリプトはpython_thread_test.serviceから起動されることを想定している。
      Ex. systemctl start python_thread_test.service
      Ex. systemctl stop python_thread_test.service
      スクリプト単体で実行するとゾンビプロセス(親がすぐ死ぬので)となる。
"""

INSTANCE = 'python_daemon_test@{0}.service'
KILL = '^python_daemon_test$'
TARGETS = [123, 234, 345]


def run_command(cmd: list):
    result = subprocess.run(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL)
    return result


def execute_thread(method: callable, targets: list, max_workers: int) -> int:
    ret = 1

    queue = Queue()
    queue.queue = deque(targets)

    workers_cnt = max_workers
    queue_size = queue.qsize()
    if queue_size < max_workers:
        workers_cnt = queue_size

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []
        for _ in range(workers_cnt):
            futures.append(executor.submit(method, queue))

        for future in as_completed(futures):
            try:
                result = future.result()
                if not result:
                    continue
                ret = 0
            except Exception as e:
                print(f'{e}')
    return ret


def _start_service(queue: Queue) -> int:
    ret = 1

    while not queue.empty():
        try:
            num = queue.get_nowait()
        except Exception:
            break

        try:
            # サービス開始
            instance = INSTANCE.format(num)
            args = ['systemctl', 'start', instance]
            result = run_command(args)
            if result.returncode != 0:
                print(f'cmd failed!. [{args}]')
                return ret
            ret = 0
        except Exception as e:
            print(f'{e}')
            return ret

    return 0


def start() -> int:
    ret = 0

    pid = os.fork()

    if pid > 0:
        # 親プロセスの場合(pidは子プロセスのプロセスID)
        pass
    elif pid == 0:
        # 子プロセスの場合
        ret = execute_thread(method=_start_service,
                             targets=TARGETS, max_workers=2)
    else:
        # メモリ不足でfork失敗
        ret = -1

    return ret


def stop() -> int:
    ret = 1

    try:
        # サービス開始
        args = ['pkill', 'USR', KILL]
        result = run_command(args)
        if result.returncode != 0:
            print(f'cmd failed!. [{args}]')
            return ret
        ret = 0
    except Exception as e:
        print(f'{e}')
        return ret

    return ret


def main() -> int:
    ret = 1

    args = sys.argv
    if len(args) > 1:
        method = args[1]
        if method == 'start':
            ret = start()
        elif method == 'stop':
            ret = stop()
    return ret


if __name__ == '__main__':
    sys.exit(main())
