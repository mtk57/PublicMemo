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
