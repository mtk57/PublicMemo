#!/usr/bin/env python

import traceback
import json


class AttrDict():
    """ JavaScriptライクに属性アクセス可能なdict
        ただし、dictは継承していないので最低限の機能のみを提供する。

        <使用例>
        data = json.loads(
                '{"v1":"apple", "v2":"banana"}',
                object_hook=AttrDict)
        print("data.v1 = {0}".format(data.v1))

        data2 = json.loads(
                '{"arr":[{"v":"orange"}, {"v":"mango"}]}',
                object_hook=AttrDict)
        print("data2[0].v = {0}".format(data2.arr[0].v))

        data3 = json.loads(
                '{"group2":{"Eric":44, "ken":33, "John":44, "Mike":99},\
                "group1":{"Adam":40, "David":71, "Chris":60, "Bob":74}}',
                object_hook=AttrDict)
        print("data3.group2.Eric = {0}".format(data3.group2.Eric))

        # RESULT
        data.v1 = apple
        data2[0].v = orange
        data3.group2.Eric = 44
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

    def __getattr__(self, name: str):
        """ 未定義属性の取得
            name:str:属性名
            ret:属性値(存在しなければNone)
        """

        if name in self._dict:
            val = self._dict.get(name)

            # Valueがdictなら自身のオブジェクトを返す
            if isinstance(val, dict):
                return AttrDict(val)
            else:
                return val
        else:
            return None

    def __setattr__(self, name: str, value):
        """ 属性への代入 """

        if name == '_dict':
            object.__setattr__(self, name, value)
        else:
            self._dict[name] = value

    def __getitem__(self, key):
        """ [キー]の値を返す """
        return self._dict[key]

    def __setitem__(self, key, value):
        """ [キー]に値を代入する """
        self._dict[key] = value

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
        # from os import path
        # filename = 'test.json'
        # filepath = path.join(path.dirname(__file__), filename)

        # with open(filepath, 'r') as f:
        #     jsonData = json.load(f)

        # a = AttrDict(jsonData)
        # # a.group2 = 123
        # b = a.group2
        # # a.group2.Eric = 123
        # c = a.group2.Eric

        data = json.loads(
                '{"v1":"apple", "v2":"banana"}',
                object_hook=AttrDict)
        print("data.v1 = {0}".format(data.v1))

        data2 = json.loads(
                '{"arr":[{"v":"orange"}, {"v":"mango"}]}',
                object_hook=AttrDict)
        print("data2[0].v = {0}".format(data2.arr[0].v))

        data3 = json.loads(
                '{"group2":{"Eric":44, "ken":33, "John":44, "Mike":99},\
                "group1":{"Adam":40, "David":71, "Chris":60, "Bob":74}}',
                object_hook=AttrDict)
        print("data3.group2.Eric = {0}".format(data3.group2.Eric))

    except:
        print(traceback.format_exc())
