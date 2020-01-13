import subprocess as sp
import sys
import time


class Command():
    """ 子プロセスでシェルコマンドを実行して結果を返す """

    def __init__(self, args):
        """ args:シェルコマンド """
        self._args = args

    def run(self) -> list:
        """ コマンドを同期実行して結果をリストで返す """
        try:
            result = sp.run(
                        self._args,
                        shell=True,
                        check=True,
                        capture_output=True,
                        text=True,
                        )
            return result.stdout.splitlines()

        except sp.CalledProcessError:
            print('外部プログラムの実行に失敗しました', file=sys.stderr)
            raise

    def run_async(self):
        """ コマンドを非同期実行して結果を文字列(改行コード付)で返す """
        try:
            result = sp.Popen(
                    self._args,
                    shell=True,
                    stdout=sp.PIPE,
                    stderr=sp.STDOUT,
                    universal_newlines=True,
                    )

            while True:
                line = result.stdout.readline()
                if line:
                    yield line
                if not line and result.poll() is not None:
                    break

        except Exception:
            print('外部プログラムの実行に失敗しました', file=sys.stderr)
            raise

    def validate_volinfo(self, isasync=False) -> dict:
        import re
        from copy import deepcopy
        dict_main = {}
        dict_sub = {}
        list_brk = []
        ptn_brk = re.compile(r'Brick\d')
        last_volname = ''

        if not isasync:
            func = self.run
        else:
            func = self.run_async

        for line in func():
            split_line = self._split_line(line)
            print(split_line)

            if split_line[0] == 'Volume Name':
                if not last_volname:
                    last_volname = split_line[1].strip()
                else:
                    dict_sub['Bricks'] = deepcopy(list_brk)
                    dict_main[last_volname] = deepcopy(dict_sub)
                    list_brk = []
                    dict_sub = {}
                    last_volname = split_line[1].strip()
            else:
                if len(split_line) == 0 or not split_line[0]:
                    continue
                if ptn_brk.match(split_line[0]):
                    info = BrickInfo(
                            split_line[0],
                            split_line[1].strip(),
                            split_line[2].strip()
                            )
                    list_brk.append(info)
                else:
                    dict_sub[split_line[0]] = split_line[1].strip()

        dict_sub['Bricks'] = deepcopy(list_brk)
        dict_main[last_volname] = deepcopy(dict_sub)

        return dict_main

    def _split_line(self, line: str) -> str:
        return line.replace('\n', '').strip().split(':')


class BrickInfo():
    def __init__(self, num: str, node: str, path: str):
        self._num = num
        self._node = node
        self._path = path

    @property
    def num(self) -> str:
        return self._num

    @property
    def node(self) -> str:
        return self._node

    @property
    def path(self) -> str:
        return self._path


if __name__ == '__main__':
    try:
        cmd = Command('dir')

        print('run() START')
        start_run = time.time()

        for line in cmd.run():
            print(line)

        elapsed_time_run = time.time() - start_run
        print('run() END')

        print('run_async() START')
        start_run_async = time.time()

        for line in cmd.run_async():
            print(line, end="")

        elapsed_time_run_async = time.time() - start_run_async
        print('run_async() END')

        print('Result:')
        print(f'run()      :{elapsed_time_run}')
        print(f'run_async():{elapsed_time_run_async}')

    except Exception as e:
        print(e)
