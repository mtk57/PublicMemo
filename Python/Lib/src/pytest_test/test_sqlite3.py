import pytest
import src.sqlite3_use_module as db


class TestDbMgr():

    def test_open_connect_none_return(self, mocker):
        dbmgr = db.DbMgr()
        mocker.patch("sqlite3.connect", return_value=None)
        with pytest.raises(ValueError):
            dbmgr.open()

    def test_open_connect_throw_exception(self, mocker):
        dbmgr = db.DbMgr()
        mocker.patch("sqlite3.connect", side_effect=IOError())
        with pytest.raises(IOError):
            dbmgr.open()

    def test_close_close_throw_exception(self, mocker):
        dbmgr = db.DbMgr()
        dbmgr.open()
        mm = mocker.MagicMock()
        mm.close = mocker.Mock(side_effect=IOError())
        mocker.patch.object(dbmgr, "_conn", mm)
        with pytest.raises(IOError):
            dbmgr.close()

    def test_open(self):
        dbmgr = db.DbMgr()
        dbmgr.open()

    def test_close(self):
        dbmgr = db.DbMgr()
        dbmgr.open()
        dbmgr.close()
