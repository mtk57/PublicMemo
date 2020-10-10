#!/usr/bin/env python3
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from queue import Queue, deque
from enum import IntEnum

import util
"""
ThreadPoolExecutorを使って、ssh接続をする実験
"""

# キー
K_IP = 'ip'
K_USER = 'user'
K_PW = 'pw'
K_CMD = 'cmd'
K_CONN_TO = 'conn_to'
K_CMD_TO = 'cmd_to'

EXE_CMD_OK = 'python3 /tmp/target.py 0'
EXE_CMD_NG = 'python3 /tmp/target.py 1'
VAGRANT = 'vagrant'
CONN_TIMEOUT = 10
CMN_TIMEOUT = 10

# 実行対象
SSH_MODELS = [
    {
        K_IP: '10.0.0.10',
        K_USER: VAGRANT,
        K_PW: VAGRANT,
        K_CONN_TO: CONN_TIMEOUT,
        K_CMD_TO: CMN_TIMEOUT,
        K_CMD: EXE_CMD_NG
    },
    {
        K_IP: '10.0.0.11',
        K_USER: VAGRANT,
        K_PW: VAGRANT,
        K_CONN_TO: CONN_TIMEOUT,
        K_CMD_TO: CMN_TIMEOUT,
        K_CMD: EXE_CMD_OK
    }
]


class Result(IntEnum):
    SUCCESS = 0
    FAILED = 1


def execute_thread(method: callable, targets: list, max_workers: int) -> int:
    ret = Result.FAILED

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

        # as_completedは処理が終わったタスクから結果を返していくジェネレータ
        for future in as_completed(futures):
            try:
                # 処理結果を取得
                result = future.result()
                if result != Result.SUCCESS:
                    print(f'task failed.[{result}]')
                    failed_cnt = failed_cnt + 1
                else:
                    print(f'task success.[{result}]')
            except Exception as e:
                print(f'{e}')
                failed_cnt = failed_cnt + 1
    if failed_cnt == 0:
        ret = Result.SUCCESS
    else:
        ret = Result.FAILED
        print(f'failed count={failed_cnt}')
    return ret


def ssh_exec_command(queue: Queue) -> int:
    ret = Result.FAILED

    while not queue.empty():
        try:
            ssh_model = queue.get_nowait()
        except Exception:
            break

        try:
            cmd_result = util.ssh_run_command(ssh_model=ssh_model)

            ret = cmd_result.ret_code

            if ret != Result.SUCCESS:
                print(cmd_result.stderr)

        except util.SshConnectError as e:
            print(f'{e}')
            return ret
        except util.SshExecCommandError as e:
            print(f'{e}')
            return ret
        except util.SshTimeoutError as e:
            print(f'{e}')
            return ret
        except Exception as e:
            print(f'{e}')
            return ret

    return ret


def main() -> int:
    print(f'main() START')
    ret = 0
    is_seq = False
    # is_seq = True
    worker_cnt = 2

    ssh_models = []

    for t in SSH_MODELS:
        ssh_models.append(
            util.SshCommandModel(
                ip=t[K_IP],
                user=t[K_USER],
                password=t[K_PW],
                connect_timeout=t[K_CONN_TO],
                command_timeout=t[K_CMD_TO],
                command=t[K_CMD]
            )
        )

    if is_seq:
        # シーケンシャル版
        print('Sequence start')
        queue = Queue()
        queue.queue = deque(ssh_models)
        ret = ssh_exec_command(queue)
        print('Sequence end')
    else:
        # 非同期版
        print('Async start')
        ret = execute_thread(method=ssh_exec_command,
                             targets=ssh_models, max_workers=worker_cnt)
        print('Async end')

    print(f'main() END ({ret})')
    return ret


if __name__ == '__main__':
    sys.exit(main())
