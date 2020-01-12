import unittest
from unittest.mock import patch

from dbmgr import DbMgr


class Cursor():
    """ sqlite3.Cursorのモック """
    def execute(self, *args):
        pass

    def executemany(self, *args):
        pass


class Connection():
    """ sqlite3.Connectionのモック """
    def cursor(self):
        return Cursor()

    def commit(self):
        pass

    def close(self):
        pass


class TestDbMgr(unittest.TestCase):
    """ DbMgrのユニットテストクラス """

    @patch('sqlite3.connect', return_value=Connection())
    def test_dbmgr_mock_version(self, patched_object):
        """ DbMgrのテスト(モックバージョン)
            sqlite3.connectをモック(__main__.Connection)で置き換えてテストする。
        """
        ret = self._dbmgr_run()
        self.assertTrue(ret)
        self.assertTrue(patched_object.called)

    def test_dbmgr(self):
        """ DbMgrのテスト """
        ret = self._dbmgr_run()
        self.assertFalse(ret)

    def _dbmgr_run(self) -> bool:
        """ DbMgr.run()の実行 """
        db = DbMgr()
        return db.run()


if __name__ == '__main__':
    unittest.main(verbosity=2)
