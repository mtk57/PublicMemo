import unittest
from unittest.mock import patch
from unittest.mock import MagicMock
import fcntl_mock

from sqlite3_mock import Connection
from dbmgr import DbMgr
from command import Command
from subprocess_mock import CompletedProcess, Popen

# fcntlモジュールは、Windowsには存在しないため、
# モックに書き換える
import sys
sys.modules['fcntl'] = fcntl_mock
sys.modules['fcntl.flock'] = fcntl_mock.flock

# flake8に怒られるが,flockerでfcntlがimportされる前に↑の書き換えを
# 行う必要があるため直せない。
import flocker


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


class TestFlocker(unittest.TestCase):
    def test_flocker_mock_version(self):
        """ flockのテスト(正常系) """
        flocker.flocker()

    @patch('fcntl.flock', MagicMock(side_effect=IOError()))
    def test_flocker_failed_test_mock_version(self):
        """ flockのテスト(異常系) """
        flocker.flocker()


class TestSubprocess(unittest.TestCase):
    @patch('subprocess.run', return_value=CompletedProcess('test.log'))
    def test_subprocess_run_mock_version(self, patched_object):
        """ subprocess.run()の戻り値をテキストファイルの中身に差替 """
        cmd = Command('dir')
        for line in cmd.run():
            print(line, end="")
        self.assertTrue(patched_object.called)

    @patch('subprocess.run', return_value=CompletedProcess('volinfo_test.txt'))
    def test_subprocess_run_mock_version_volinfo(self, patched_object):
        """ subprocess.run()の戻り値をテキストファイルの中身に差替 """
        cmd = Command('gluster volume info')

        volinfo = cmd.validate_volinfo()
        self.assertIsInstance(volinfo, dict)

        self.assertTrue(patched_object.called)

    @patch('subprocess.Popen', return_value=Popen('test.log'))
    def test_subprocess_popen_mock_version(self, patched_object):
        """ subprocess.Popen()の戻り値をテキストファイルの中身に差替 """
        cmd = Command('dir')
        for line in cmd.run_async():
            print(line, end="")
        self.assertTrue(patched_object.called)


if __name__ == '__main__':
    unittest.main(verbosity=2)

"""
volinfo_test.txt

Volume Name: vol1
    Type: Distributed-Replicate
    Volume ID: 24a0437a-daa0-4044-8acf-7aa82efd76fd
    Status: Started
    Number of Bricks: 2 x 2 = 4
    Transport-type: tcp
    Bricks:
    Brick1: Server1:/home/gfs/r1_0
    Brick2: Server2:/home/gfs/r1_1
    Brick3: Server1:/home/gfs/r1_2
    Brick4: Server2:/home/gfs/r1_3

Volume Name: vol2
    Type: Distributed-Replicate
    Volume ID: 24a0437a-daa0-4044-8acf-7aa82efd76fe
    Status: Started
    Number of Bricks: 2 x 2 = 4
    Transport-type: tcp
    Bricks:
    Brick1: Server1:/home/gfs/r2_0
    Brick2: Server2:/home/gfs/r2_1
    Brick3: Server1:/home/gfs/r2_2
    Brick4: Server2:/home/gfs/r2_3
"""