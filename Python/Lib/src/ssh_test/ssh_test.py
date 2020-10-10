#!/usr/bin/env python3
import paramiko
import socket
"""
API doc : http://docs.paramiko.org/en/stable/api/client.html

connect(
    hostname,                        (str)   - 接続先のサーバ。
    port=22,                         (int)   - 接続先のサーバポート。
    username=None,                   (str)   - 認証するユーザ名 (デフォルトは現在のローカルユーザ名)
    password=None,                   (str)   - パスワード認証に使用されます。
                                               パスフレーズが指定されていない場合は、秘密鍵の復号にも使用されます。
    pkey=None,                       (PKey)  - 認証に使用するオプションの秘密鍵。
    key_filename=None,               (str)   - 認証を試みるためのオプションの秘密鍵 (複数可) および/または証明書のファイル名、またはファイル名のリスト。
    timeout=None,                    (float) - TCP 接続のオプションのタイムアウト (秒単位)。
    allow_agent=True,                (bool)  - False に設定して SSH エージェントへの接続を無効にする
    look_for_keys=True,              (bool)  - False に設定すると、~/.ssh/ 内の検出可能な秘密鍵ファイルの検索を無効にします。
    compress=False,                  (bool)  - True に設定して圧縮を有効にします。
    sock=None,                       (socket)- ターゲットホストとの通信に使用するオープンソケットまたはソケットのようなオブジェクト (Channel など)。
    gss_auth=False,                  (bool)  - GSS-API 認証を使用する場合に true を指定する
    gss_kex=False,                   (bool)  - GSS-API 鍵交換およびユーザ認証を行う
    gss_deleg_creds=True,            (bool)  - GSS-API クライアントの認証情報を委任するかどうか。
    gss_host=None,                   (str)   - kerberos データベース内のターゲット名。
    banner_timeout=None,             (float) - SSH バナーが表示されるのを待つためのオプションのタイムアウト (秒単位)。
    auth_timeout=None,               (float) - オプションのタイムアウト (秒単位) で、認証応答を待ちます。 
    gss_trust_dns=True,              (bool)  - 接続先のホスト名を安全に正規化するために DNS を信頼するかどうかを示します (デフォルトは True)。
    passphrase=None,                 (str)   - 秘密鍵の復号化に使用します。
    disabled_algorithms=None         (dict)  - Transport に直接渡されるオプションの dict と、同じ名前のキーワード引数。
)

SSH サーバに接続して認証します。
サーバのホスト鍵はシステムのホスト鍵 (load_system_host_keys を参照) と
ローカルのホスト鍵 (load_host_keys) と照合されます。
サーバのホスト名がどちらのホスト鍵にも見つからない場合は、
missing host key ポリシーが使用されます (set_missing_host_key_policy を参照してください)。
デフォルトのポリシーは、鍵を拒否して SSHException を発生させることです。

認証は以下の優先順位で試行されます。

渡された pkey または key_filename (もしあれば)
key_filename には OpenSSH 公開証明書のパスと通常の秘密鍵のパスが含まれているかもしれません。
(秘密鍵自体が key_filename にある必要はありません - 証明書だけです)。

SSH エージェントを使って見つけた鍵で発見可能な任意の「id_rsa」、「id_dsa」、または「id_ecdsa」鍵。
OpenSSH スタイルの公開証明書が存在し、それが既存の秘密鍵と一致する場合 (たとえば id_rsa と 
id_rsa-cert.pub がある場合)、その証明書は秘密鍵と一緒に読み込まれ、認証に使われます。

明瞭なユーザ名/パスワード認証、パスワードが与えられている場合は秘密鍵のロックを解除するために
パスワードを必要とし、パスワードが渡された場合は、そのパスワードを使用して鍵のロックを解除しようとします。


exec_command(
    command,            (str)  - 実行するコマンド
    bufsize=-1,         (int)  - Python の組み込み file() 関数と同じように解釈される
    timeout=None,       (int)  - コマンドのチャンネルタイムアウトを設定します。Channel.settimeout を参照ください。
    get_pty=False,      (bool) - サーバに擬似ターミナルを要求する (デフォルトは False)。Channel.get_pty を参照ください。
    environment=None    (dict) - シェル環境変数のディクショナリで、リモートコマンドが実行するデフォルト環境にマージされます。
)
SSH サーバ上でコマンドを実行します。
新しいチャネルが開かれ、要求されたコマンドが実行されます。
コマンドの入出力ストリームは、標準入力、標準出力、標準エラーを表す 
Python ファイルのようなオブジェクトとして返されます。（※1）

警告
サーバーは、一部の環境変数を拒否することがあります。
詳細については、Channel.set_environment_variableの警告を参照してください。

※1
exec_command()のノンブロッキングで、すぐに関数が戻る。
タイムアウトは、戻ったあとの戻り値（タプルっぽいが異なる。Fitureみたいなオブジェクト）がsocket.errorをスローする
"""


IP = '10.0.0.10'
USER = 'vagrant'
PW = USER
CONN_TIMEOUT = 5
CMD_TIMEOUT = 10
CMD = 'python3 /tmp/mugen_loop.py'


class SshConnectError(Exception):
    def __inin__(self, message: str):
        self.message = message


class SshExecCommandError(Exception):
    def __inin__(self, message: str):
        self.message = message


def main():
    client = None
    try:
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.WarningPolicy())

        try:
            client.connect(IP, username=USER, password=PW,
                           timeout=CONN_TIMEOUT)
        except paramiko.BadHostKeyException as e:
            raise SshConnectError(f'BadHostKeyException! [{e}]')
        except paramiko.AuthenticationException as e:
            raise SshConnectError(f'AuthenticationException! [{e}]')
        except paramiko.SSHException as e:
            raise SshConnectError(f'SSHException! [{e}]')
        except socket.error as e:
            # タイムアウトも含む
            raise SshConnectError(f'socket.error! [{e}]')

        try:
            stdin, stdout, stderr = client.exec_command(
                command=CMD, timeout=CMD_TIMEOUT)

            for line in stdout:
                print(line)
            for line in stderr:
                print(line)
        except paramiko.SSHException as e:
            raise SshExecCommandError(f'SSHException! [{e}]')
        except socket.timeout as e:
            raise SshExecCommandError(f'Command timeout! [{e}]')

    except SshConnectError as e:
        print(f'The ssh connection failed. [{e}]')
    except SshExecCommandError as e:
        print(f'An error occurred when executing the command. [{e}]')
    except Exception as e:
        print(f'An unexpected error occurred. [{e}]')
    finally:
        if client:
            client.close()


main()
