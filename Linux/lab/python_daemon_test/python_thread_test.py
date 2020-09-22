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

    failed_cnt = 0
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []
        for _ in range(workers_cnt):
            futures.append(executor.submit(method, queue))

        for future in as_completed(futures):
            try:
                result = future.result()
                if result != 0:
                    print(f'task failed.[{result}]')
                    failed_cnt = failed_cnt + 1
                else:
                    print(f'task success.[{result}]')
            except Exception as e:
                print(f'{e}')
    if failed_cnt == 0:
        ret = 0
    else:
        ret = 1
        print(f'failed count={failed_cnt}')
    return ret


def start_service(queue: Queue) -> int:
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


def start(max_instance=5, is_seq=False, worker_cnt=2) -> int:
    print(
        f'start() START  max_instance={max_instance}, is_sqe={is_seq}, worker_cnt={worker_cnt}')

    ret = 0

    # pid = os.fork()

    # if pid > 0:
    #     # 親プロセスの場合(pidは子プロセスのプロセスID)
    #     pass
    # elif pid == 0:
    #     # 子プロセスの場合

    TARGETS = list(range(max_instance))
    print(f'targets={TARGETS}')

    if is_seq:
        # シーケンシャル版
        print('Sequence start')
        queue = Queue()
        queue.queue = deque(TARGETS)
        ret = start_service(queue)
        print('Sequence end')
    else:
        # 非同期版
        print('Async start')
        ret = execute_thread(method=start_service,
                             targets=TARGETS, max_workers=worker_cnt)
        print('Async end')
    # else:
    #     # メモリ不足でfork失敗
    #     ret = -1

    print(f'start() END ({ret})')
    return ret


def stop() -> int:
    ret = 1

    print('stop() START')

    try:
        # サービス開始
        args = ['pkill', 'python3']
        print(f'exec cmd={args}')
        result = run_command(args)
        if result.returncode != 0:
            print(f'cmd failed!. [{args}]')
            return ret
        ret = 0
    except Exception as e:
        print(f'{e}')
        return ret

    print(f'stop() END ({ret})')
    return ret


def main() -> int:
    ret = 1

    args = sys.argv
    if len(args) > 1:
        method = args[1]

        max_instance = 5
        if len(args) > 2:
            max_instance = int(args[2])

        is_seq = False
        if len(args) > 3:
            is_seq = bool(args[3])

        worker_cnt = 2
        if len(args) > 4:
            worker_cnt = int(args[4])

        if method == 'start':
            ret = start(max_instance, is_seq, worker_cnt)
        elif method == 'stop':
            ret = stop()
    return ret


if __name__ == '__main__':
    sys.exit(main())
