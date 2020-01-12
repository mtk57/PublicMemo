
class DbMgr():
    """ sqlite3を用いたDBマネージャークラス """

    def run(self) -> bool:
        import sqlite3

        # データベースファイルのパス
        dbpath = 'sample_db.sqlite'

        # データベース接続とカーソル生成
        conn = sqlite3.connect(dbpath)
        # 自動コミットにする場合は下記を指定（コメントアウトを解除のこと）
        # connection.isolation_level = None
        cur = conn.cursor()

        # エラー処理（例外処理）
        try:
            # CREATE
            cur.execute("DROP TABLE IF EXISTS sample")
            cur.execute(
                "CREATE TABLE IF NOT EXISTS sample "
                "(id INTEGER PRIMARY KEY, name TEXT)"
                )

            # INSERT
            cur.execute("INSERT INTO sample VALUES (1, '佐藤')")
            # プレースホルダの使用例
            # プレースホルダには疑問符(qmark スタイル)と名前(named スタイル)の2つの方法がある
            # 1つの場合には最後に , がないとエラー。('鈴木') ではなく ('鈴木',)
            cur.execute("INSERT INTO sample VALUES (2, ?)", ('鈴木',))
            cur.execute("INSERT INTO sample VALUES (?, ?)", (3, '高橋'))
            cur.execute("INSERT INTO sample VALUES (:id, :name)",
                        {'id': 4, 'name': '田中'})
            # 複数レコードを一度に挿入 executemany メソッドを使用
            persons = [
                (5, '伊藤'),
                (6, '渡辺'),
            ]
            cur.executemany("INSERT INTO sample VALUES (?, ?)", persons)
            # わざと主キー重複エラーを起こして例外を発生させてみる
            cur.execute("INSERT INTO sample VALUES (1, '中村')")
        except sqlite3.Error as e:
            print('sqlite3.Error occurred:', e.args[0])
            return False

        # 保存を実行（忘れると保存されないので注意）
        conn.commit()

        # 接続を閉じる
        conn.close()

        return True
