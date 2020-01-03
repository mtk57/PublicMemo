#!/usr/bin/env python

from collections import UserList, defaultdict
import traceback


class DictList(UserList):
    """ 要素がdictであることを保証するlist
        <使用例>
        dl = DictList(keys=['key1', 'key2'])        # キー名はdictのキーに使える型と同じ
        dl.set_values(values=(123, 'apple'))
        dl.set_values(values=(456, 'banana'))
        values = dl.get_values(key='key2')          # ['apple', 'banana']
    """

    def __init__(self, keys: list):
        """ 初期化する
            keys:list:dictのキー名 (ex. DictList(['key1', 1]))
                      重複したキーがある場合はエラーとなる
                      キー名はlistで指定すること(setだと重複しないが順序も不定なので)
            ret:-
        """

        if not isinstance(keys, list):
            raise TypeError('keys is not list!')
        if None in keys:
            raise TypeError('None is not key!')
        if len(keys) == 0:
            raise TypeError('Key is nothing!')
        if len(set(keys)) != len(keys):
            raise TypeError('Duplicate Key Name!')

        self._keys = keys
        super().__init__([])

    @property
    def keys(self) -> list:
        """ dictのキー名(list) """

        return self._keys

    @property
    def values(self) -> list:
        """ dictの値(list) """

        vals = []
        for dic in self.data:
            vals.extend(dic.values())
        return vals

    def get_values(self, key) -> list:
        """ キー名に対応する全ての値をlistで返す (ex. dl.get_values('key1'))
            key:キー名
            ret:list:キーに対応する全ての値
        """

        if key not in self._keys:
            raise TypeError('key not exist!')

        return [dic[key] for dic in self.data]

    def set_values(self, values: tuple):
        """ dictの値をtupleで指定する
            values:tuple:dictの値 (ex. dl.set_values(('val1', 'val2')))
            ret:-
        """

        if not isinstance(values, tuple):
            raise TypeError('keys is not tuple!')
        if len(values) != len(self._keys):
            raise TypeError('Values count unmatch!')

        # 辞書を作成してリストに追加
        item = {k: v for k, v in zip(self._keys, values)}
        self.data.append(item)

    def _validate(self, item: dict) -> bool:
        """ listの要素を検証する
            item:dict:listの要素
            ret:bool:True=成功
        """

        if not isinstance(item, dict):
            raise TypeError('item is not dict!')
        if len(item.keys()) != len(self._keys):
            raise TypeError('Keys count unmatch!')
        if set(item.keys()) != set(self._keys):
            raise TypeError('Key name unmatch!')
        return True

    def __setitem__(self, index: int, item: dict):
        """ listに要素を代入する
            index:int:listのインデクス
            item:dict:代入する要素
            ret:-
        """

        self._validate(item)
        self.data[index] = item

    def __getitem__(self, index: int) -> dict:
        """ listから要素を取得する
            index:int:listのインデクス
            ret:dict:取得した要素
        """

        if isinstance(index, int) \
           and not isinstance(self.data[index], dict):
            raise TypeError('item is not dict!')
        return self.data[index]

    def append(self, item: dict):
        """ 要素を追加する
            item:dict:追加する要素
            ret:-
        """

        self._validate(item)
        super(DictList, self).append(item)

    def extend(self, seq: list):
        """ リストの要素を追加する
            seq:list:追加シーケンス
            ret:-
        """

        if not isinstance(seq, list):
            raise TypeError('seq is not list!')
        for item in seq:
            self._validate(item)
        super(DictList, self).extend(seq)

    def insert(self, index: int, item: dict):
        """ listに要素を挿入する
            index:int:listのインデクス
            item:dict:挿入する要素
            ret:-
        """

        self._validate(item)
        super(DictList, self).insert(index, item)


def _print(seq):
    for s in seq:
        print(s)


if __name__ == '__main__':
    try:
        # TEST >>
        # DictListオブジェクトを作成
        # キー名はlistで指定すること(setだと重複しないが順序も不定なので)
        keys = [1, 2, 3]                                # OK
        diclist = DictList(keys=keys)

        # キー名をプロパティから取得
        getKeys = diclist.keys
        _print(getKeys)

        # キーに対応する値を設定する(辞書が作成される)
        # 値がキーの数と一致していること
        diclist.set_values(values=(10, 20, 30))         # OK
        diclist.set_values(values=('A', 'B', 'C'))      # OK

        vals = diclist.values

        for dic in diclist:
            a, b, c = dic[1], dic[2], dic[3]
            print('{}, {}, {}'.format(a, b, c))

        import copy
        aa = copy.copy(diclist)
        bb = copy.deepcopy(diclist)

        # キー名を指定して、全ての値をリストで取得する
        v1 = diclist.get_values(key=1)                  # OK
        v2 = diclist.get_values(key=2)                  # OK
        v3 = diclist.get_values(key=3)                  # OK
        print('v1={}, v2={}, v3={}'.format(v1, v2, v3))

        dd = defaultdict(DictList)
        dd['aa'] = diclist                              # OK
        print('defaultdict={}'.format(dd))              # OK
        # TEST <<

    except:
        print(traceback.format_exc())

"""
_UserList__cast
__abstractmethods__
__add__
__class__
__contains__
__copy__
__delattr__
__delitem__
__dict__
__dir__
__doc__
__eq__
__format__
__ge__
__getattribute__
__getitem__
__gt__
__hash__
__iadd__
__imul__
__init__
__init_subclass__
__iter__
__le__
__len__
__lt__
__module__
__mul__
__ne__
__new__
__radd__
__reduce__
__reduce_ex__
__repr__
__reversed__
__rmul__
__setattr__
__setitem__
__sizeof__
__slots__
__str__
__subclasshook__
__weakref__
_abc_impl
append
clear
copy
count
extend
index
insert
pop
remove
reverse
sort
None
"""
