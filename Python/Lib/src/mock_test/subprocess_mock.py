class Stdout():
    def __init__(self, path: str):
        self._splitlines = self._read_file(path)

    def splitlines(self) -> list:
        return self._splitlines

    def _read_file(self, path: str) -> list:
        with open(path, 'r') as f:
            return list(f)


class CompletedProcess():
    def __init__(self, path: str) -> Stdout:
        self.stdout = Stdout(path)


def run(*args, path: str):
    return CompletedProcess(path)
