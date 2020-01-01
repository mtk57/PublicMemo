import traceback

class DistinctError(ValueError):
    pass

class DistinctDict(dict):
    """値の重複を許さない辞書"""
    def __setitem__(self, key, value):
        if value in self.values():      # 引数の値は既にあるか?

            cond1 = key in self
            cond3 = key not in self
            cond2 = self[key] != value
            #cond3 = key not in self

            if( (cond1 and cond2) or cond3 ):   # この判定がよくわからない。「if value in self.values():」がTrueならすぐ例外スローでよいのでは?
                raise DistinctError("この値はすでに別のキーで使用されています")

        super().__setitem__(key, value)

if __name__ == '__main__':  
    try:
        my = DistinctDict()

        my[ 'key1' ] = 123
        #my[ 'key1' ] = 123   # cond1 が True になるケース
        my[ 'key2' ] = 123
    except Exception as e:
        print(traceback.format_exc())
