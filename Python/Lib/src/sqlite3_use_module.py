import sqlite3


class DbMgr():
    def open(self):
        try:
            self._conn = sqlite3.connect('sample.db')
            if not self._conn:
                raise ValueError('connect ret is None')
        except Exception as e:
            print(f'open failed. ex={type(e)}')
            raise

    def close(self):
        try:
            self._conn.close()
        except Exception as e:
            print(f'close failed. ex={type(e)}')
            raise


if __name__ == '__main__':
    mgr = DbMgr()
    try:
        mgr.open()
    finally:
        mgr.close()
