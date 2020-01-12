
class Stdout():
    def __init__(self, path: str):
        self._path = path
        self._splitlines = self._read_file(path)
        self.is_end = False

    def splitlines(self) -> list:
        return self._splitlines

    def openfile(self):
        self._fp = open(self._path, 'r')

    def readline(self) -> str:
        line = self._fp.readline()
        if line:
            return line
        else:
            self.is_end = True
            self._fp.close()

    def _read_file(self, path: str) -> list:
        with open(path, 'r') as f:
            return list(f)


class CompletedProcess():
    def __init__(self, path: str) -> Stdout:
        self.stdout = Stdout(path)


class Popen():
    def __init__(self, path: str) -> Stdout:
        self.stdout = Stdout(path)
        self.stdout.openfile()

    def poll(self):
        if self.stdout.is_end:
            return True
        else:
            return None
