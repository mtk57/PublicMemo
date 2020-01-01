import traceback

class StrDict(dict):
    """
    キーの型が文字列であることを保証する辞書型
    """
    def __init__(self):
        pass
    
    def __setitem__(self, key, value):
        """ 辞書のセッターをオーバーライドして
            キーが文字列のみ許容する
        """
        if not isinstance(key, str):
            raise ValueError('Key must be str ir unicode.')

        #dict.__setitem__(self, key, value)
        super().__setitem__(key, value)

if __name__ == '__main__':
    try:
        d = StrDict()
        d["abc"] = 123
        print(d)
        d[1] = 456
    except:
        print(traceback.format_exc())