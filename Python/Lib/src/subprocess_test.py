import subprocess
import sys


try:
    """
    subprocess.run(
        args,                       ：指定されたコマンドを実行します。
                                      コマンドの完了を待って、CompletedProcessを返します。
        bufsize=-1,                 ：int。バッファサイズ。-1=システム依存。1以上。
        executable=None,            ：置換プログラムを指定する。
        stdin=None,                 ：int。標準入力。(*1)
        stdout=None,                ：int。標準出力。(*1)
        stderr=None,                ：int。標準エラー出力。(*1)
        preexec_fn=None,            ：呼び出し可能なオブジェクトを指定すると子プロセスで呼び出される
        close_fds=True,             ：bool。True=子プロセスが実行される前に標準入出力及びエラー以外のファイル記述子が閉じられる。
        shell=False,                ：bool。True=shell依存のコマンドを実行するときに使う。
        cwd=None,                   ：str。子プロセスの作業ディレクトリを指定する。
        env=None,                   ：子プロセスの環境変数を指定する。
        universal_newlines=None,    ：bool。True=標準入出力及びエラーに文字列が返る。False=バイト列が返る。
                                      互換性保持のために残してある。textを使う事。
        startupinfo=None,           ：Windowsのみ。STARTUPINFOオブジェクトを指定する。
        creationflags=0,            ：int。Windowsのみ。CreateProcess()に渡される。
        restore_signals=True,       ：bool。True=すべてのシグナルは子プロセス実行前にSIG_DFLに戻される。
        start_new_session=False,    ：bool。True=子プロセスでsetsid()システムコールが実行される。
        pass_fds=(),                ：親子プロセス間で使うファイル記述子のシーケンスを指定する。
                                      これを指定すると自動的にclose_fdsにはTrueが入る。
        encoding=None,              ：標準入出力及びエラーを文字列で返す場合の文字コード。
        errors=None,                ：標準入出力及びエラーを文字列で返す場合のデコードエラー部分を指定文字列で置換する。

        <ここより下はPopenにはない>
        input=None,                 ：子プロセスの標準入力に渡される。文字列もしくはバイト列で指定する。
        capture_output=False,       ：bool。True=stdoutとstderrがキャプチャされる。
        timeout=None,               ：int。子プロセスの実行タイムアウト値(秒)。TimeoutExpiredがスローされる。
        check=False,                ：bool。True=run()の戻り値(CompletedProcess.returncode)が0以外だとCalledProcessErrorがスローされる。
        text=None,                  ：bool。universal_newlinesのエイリアス。
        )

    *1
    有効な値は以下。(1と2はsubprocessの定数)
      1.PIPE
         →新しいパイプが子プロセスに向けて作られる
      2.DEVNULL
         →特殊ファイル os.devnull が使用される
      3.既存のファイル記述子 (正の整数) ※STDIN, STDOUT, STDERR
      4.既存のファイルオブジェクト(ファイルポインタ)
      5.None
         →リダイレクトは起こらない。
          子プロセスのファイルハンドルはすべて親から受け継がれます。
    stderr を STDOUT にすると、子プロセスの標準エラー出力からの出力は標準出力と同じファイルハンドルに出力されます。


    run()の全引数はPopen()に渡される。
    Popen()は非同期で実行したい場合に使う。

    subprocess.Popen(
        args,                       ：runと同じ
        bufsize=-1,                 ：runと同じ
        executable=None,            ：runと同じ
        stdin=None,                 ：runと同じ
        stdout=None,                ：runと同じ
        stderr=None,                ：runと同じ
        preexec_fn=None,            ：runと同じ
        close_fds=True,             ：runと同じ
        shell=False,                ：runと同じ
        cwd=None,                   ：runと同じ
        env=None,                   ：runと同じ
        universal_newlines=False,   ：runと同じ
        startupinfo=None,           ：runと同じ
        creationflags=0,            ：runと同じ
        restore_signals=True,       ：runと同じ
        start_new_session=False,    ：runと同じ
        pass_fds=(),                ：runと同じ
        encoding=None,              ：runと同じ
        errors=None                 ：runと同じ
        )

    参考：
    https://docs.python.org/ja/3/library/subprocess.html#subprocess.Popen
    https://masayoshi-9a7ee.hatenablog.com/entry/2018/12/11/130852
    https://qiita.com/megmogmog1965/items/5f95b35539ed6b3cfa17

    """

    result = subprocess.run(
                ['dir'],
                shell=True,
                check=True,
                capture_output=True,
                text=True,
                )

    a = result.stdout.splitlines()

    for line in result.stdout.splitlines():
        print('>>> ' + line)
except subprocess.CalledProcessError:
    # check=Trueにすると、この例外がスローされる
    print('外部プログラムの実行に失敗しました', file=sys.stderr)
