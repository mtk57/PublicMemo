from importlib import import_module
import datetime


class DB_Util():
    # 統計情報のテーブル名
    TABLE_STATISTICS = 'statistics'
    # 統計情報のカラム名
    CLM_QNUM = 'q_num'                  # 問題番号(PK)
    CLM_CORRECT = 'correct'             # 正解数
    CLM_TOTAL = 'total'                 # 回答回数

    # 成績のテーブル名
    TABLE_RESULTS = 'results'
    # 統計情報のカラム名
    CLM_DATETIME = 'datetime'           # 実施日時(PK)
    CLM_QUESTIONS = 'questions'         # 問題数
    CLM_CORRECT_RATE = 'correct_rate'   # 正解率
    CLM_RESULT = 'result'               # 合否
    CLM_TIME = 'time'                   # 所要時間(秒)

    def __init__(self, path: str):
        self._path = path
        self._sqlite3 = None
        self._conn = None

    def is_exist(self) -> bool:
        """ DBが存在するか否か """
        if self._conn is None:
            return False
        return True

    def open(self) -> bool:
        try:
            self._sqlite3 = import_module("sqlite3")
        except ImportError:
            # 未インストールの場合は何もしない
            return True

        self._conn = self._sqlite3.connect(self._path)

        SQL = 'CREATE TABLE IF NOT EXISTS {0} (' \
              '  {1} INTEGER PRIMARY KEY,' \
              '  {2} INTEGER,' \
              '  {3} INTEGER' \
              ')'.format(
                  DB_Util.TABLE_STATISTICS,
                  DB_Util.CLM_QNUM,
                  DB_Util.CLM_CORRECT,
                  DB_Util.CLM_TOTAL
              )
        results = self._execute_query(SQL)
        if results[0] is False:
            return results[0]

        SQL = 'CREATE TABLE IF NOT EXISTS {0} (' \
              '  {1} TEXT PRIMARY KEY,' \
              '  {2} INTEGER,' \
              '  {3} INTEGER,' \
              '  {4} REAL,' \
              '  {5} INTEGER,' \
              '  {6} INTEGER' \
              ')'.format(
                  DB_Util.TABLE_RESULTS,
                  DB_Util.CLM_DATETIME,
                  DB_Util.CLM_CORRECT,
                  DB_Util.CLM_QUESTIONS,
                  DB_Util.CLM_CORRECT_RATE,
                  DB_Util.CLM_RESULT,
                  DB_Util.CLM_TIME
              )
        results = self._execute_query(SQL)
        if results[0] is False:
            return results[0]

    def close(self):
        if self._conn is None:
            return
        self._conn.close()

    def commit(self):
        if self._conn is None:
            return
        self._conn.commit()

    def rollback(self):
        if self._conn is None:
            return
        self._conn.rollback()

    def clear(self) -> bool:
        if self._conn is None:
            return True
        SQL = 'DELETE FROM {0}'.format(DB_Util.TABLE_STATISTICS)
        results = self._execute_query(SQL)
        if results[0] is False:
            return results[0]

        SQL = 'DELETE FROM {0}'.format(DB_Util.TABLE_RESULTS)
        results = self._execute_query(SQL)
        if results[0] is False:
            return results[0]

        self.commit()
        return True

    def update_statistics(self, question_num: int, is_correct: bool) -> bool:
        if self._conn is None:
            return True

        SQL = 'SELECT {0}, {1} FROM {2} WHERE {3}={4}'.format(
                DB_Util.CLM_CORRECT,
                DB_Util.CLM_TOTAL,
                DB_Util.TABLE_STATISTICS,
                DB_Util.CLM_QNUM,
                question_num
        )
        results = self._execute_query(SQL, is_update=False)
        if results[0] is False:
            return False

        countup = 1 if is_correct else 0

        result = results[1]
        if len(result) == 0:
            SQL = 'INSERT INTO {0}({1}, {2}, {3}) VALUES({4}, {5}, 1)'.format(
                    DB_Util.TABLE_STATISTICS,
                    DB_Util.CLM_QNUM,
                    DB_Util.CLM_CORRECT,
                    DB_Util.CLM_TOTAL,
                    question_num,
                    countup
            )
        elif len(result) == 1:
            SQL = 'UPDATE {0} SET {1}={2}, {3}={4} WHERE {5}={6}'.format(
                    DB_Util.TABLE_STATISTICS,
                    DB_Util.CLM_CORRECT,
                    result[0][0] + countup,
                    DB_Util.CLM_TOTAL,
                    result[0][1] + 1,
                    DB_Util.CLM_QNUM,
                    question_num
            )
        else:
            return False

        results = self._execute_query(SQL)
        if results[0] is False:
            return results[0]

        self.commit()
        return True

    def insert_result(self, corrects: int, questions: int, correct_rate: float,
                      result: bool, required_time: int) -> bool:
        if self._conn is None:
            return True

        SQL = 'INSERT INTO {0}({1}, {2}, {3}, {4}, {5}, {6}) VALUES("{7}", {8}, {9}, {10}, {11}, {12})'.format(
                DB_Util.TABLE_RESULTS,
                DB_Util.CLM_DATETIME,
                DB_Util.CLM_CORRECT,
                DB_Util.CLM_QUESTIONS,
                DB_Util.CLM_CORRECT_RATE,
                DB_Util.CLM_RESULT,
                DB_Util.CLM_TIME,

                str(datetime.datetime.now()),
                corrects,
                questions,
                correct_rate,
                1 if result else 0,
                int(required_time)
        )
        results = self._execute_query(SQL)
        if results[0] is False:
            return results[0]

        self.commit()
        return True

    def get_correct_rate(self, question_num: int) -> float:
        """ 正答率を返す """
        if self._conn is None:
            return 0.0

        SQL = 'SELECT {0}, {1} FROM {2} WHERE {3}={4}'.format(
                DB_Util.CLM_CORRECT,
                DB_Util.CLM_TOTAL,
                DB_Util.TABLE_STATISTICS,
                DB_Util.CLM_QNUM,
                question_num
        )
        results = self._execute_query(SQL, is_update=False)
        if results[0] is False:
            return 0.0

        result = results[1]
        if len(result) != 1:
            return 0.0

        correct = result[0][0]
        total = result[0][1]

        return round(correct / total, 2) * 100

    def _execute_query(self, sql: str, is_update: bool = True) -> tuple:
        if self._conn is None:
            return (True, [])
        results = (False, [])
        cur = self._conn.cursor()

        try:
            cur.execute(sql)
        except Exception as e:
            print(e)
            return results

        if is_update:
            return (True, [])
        return (True, cur.fetchall())
