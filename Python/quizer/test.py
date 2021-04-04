import sqlite3
import datetime

# DBを開く。適合関数・変換関数を有効にする。
conn = sqlite3.connect(
        ':memory:',
        detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)

# "TIMESTAMP"コンバータ関数 をそのまま ”DATETIME” にも使う
sqlite3.dbapi2.converters['DATETIME'] = sqlite3.dbapi2.converters['TIMESTAMP']

# カーソル生成
cur = conn.cursor()

# datetimeという型名のカラムを持つテーブルを作成
cur.execute("create table mytable(comment text, updated datetime);")

# datetimeなカラムに、文字列表現 と datetime でそれぞれ投入してみる。
cur.executemany("insert into mytable(comment, updated) value (?,?)",
        [["text_formated.", "2014-01-02 23:45:00"],
        ["datetime_class.", datetime.datetime(2014,3,4, 12,34,56)]])


ret = cur.execute("select * from mytable;")

for row in ret.fetchall():
    pass
    # print ("'%s'" % row[0], row[1], type(row[1])