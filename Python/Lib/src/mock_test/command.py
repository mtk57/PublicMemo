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
