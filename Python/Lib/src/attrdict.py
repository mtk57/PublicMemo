#!/usr/bin/env python

import traceback
import json


class AttrDict():
    """ JavaScriptライクな属性アクセス可能なdict
        <使用例>
        data = json.loads(
                '{"v1":"xxxx",
                  "v2":"vvvv"}',
                object_hook=AttrDict)
        v1 = data.v1        # 'xxxx'

        data2 = json.loads(
                '{"arr":[
                          {"v":"xxxx"},
                          {"v":"vvvv"}
                        ]}',
                object_hook=AttrDict)
        v = data2.arr[0].v  # 'xxxx'
    """

    def __init__(self, dict_: dict):
        """ 初期化する
            obj:dict:辞書オブジェクト
            ret:-
        """

        if dict_ is None:
            raise TypeError('dict_ is None!')
        if not isinstance(dict_, dict):
            raise TypeError('dict_ is not dict!')

        self._dict = dict_

    def __getstate__(self):
        """ dictのitemsを返す """

        return self._dict.items()

    def __setstate__(self, items):
        """ dictのitemsを設定する """

        if items is None:
            raise TypeError('items is None!')
        if not hasattr(self, '_dict'):
            self._dict = {}
        for key, val in items:
            self._dict[key] = val

    def __getattr__(self, name: str):
        """ 未定義属性の取得
            name:str:属性名
            ret:属性値(存在しなければNone)
        """

        if name in self._dict:
            return self._dict.get(name)
        else:
            return None

    @property
    def data(self) -> dict:
        """ dictを返す """

        return self._dict

    @property
    def keys(self) -> list:
        """ dictのキー名を返す """

        return self._dict.keys()


if __name__ == '__main__':
    try:
        data = json.loads(
                '{"v1":"xxxx", "v2":"vvvv"}', object_hook=AttrDict)
        print("data.v1 = {0}".format(data.v1))

        data2 = json.loads(
                '{"arr":[{"v":"xxxx"}, {"v":"vvvv"}]}', object_hook=AttrDict)
        print("data2[0].v = {0}".format(data2.arr[0].v))
    except:
        print(traceback.format_exc())
