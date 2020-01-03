#!/usr/bin/env python

import unittest
import sys
import traceback


# Pythonの相対importは実行したファイルより上の階層には遡れないので、
# 強引にテスト対象モジュールのパスを追加する。
sys.path.append('C:\\_Dev\\Git\\GitHub\\Lab\\Python\\src')
from dictlist import DictList


class TestDictList(unittest.TestCase):

    def setUp(self):
        self.dict_list = DictList(keys=[1, 2, 3])

    def test_dictlist_success_simple(self):
        """正常系 シンプル"""

        dl = DictList(keys=['key1', 'key2'])
        dl.set_values(values=(123, 'apple'))
        dl.set_values(values=(456, 'banana'))
        values = dl.get_values(key='key2')
        vals = dl.values

        self.assertEqual(values, ['apple', 'banana'])
        self.assertEqual(vals, [123, 'apple', 456, 'banana'])

    def test_dictlist_success_init(self):
        """正常系 (__init__)"""

        self.assertIsInstance(DictList(keys=['a', 'b']), DictList)
        self.assertIsInstance(DictList(keys=[1]), DictList)
        self.assertIsInstance(DictList(keys=[1, ]), DictList)
        self.assertIsInstance(DictList(keys=[1, 2, 3]), DictList)

    def test_dictlist_success_keys(self):
        """正常系 (keys)"""

        self.assertEqual(self.dict_list.keys, [1, 2, 3])

    def test_dictlist_success_set_get_values(self):
        """正常系 (set_values, get_values)"""

        self.dict_list.set_values(values=(10, 20, 30))
        self.dict_list.set_values(values=('A', 'B', 'C'))

        v1 = self.dict_list.get_values(key=1)
        v2 = self.dict_list.get_values(key=2)
        v3 = self.dict_list.get_values(key=3)

        self.assertEqual(v1, [10, 'A'])
        self.assertEqual(v2, [20, 'B'])
        self.assertEqual(v3, [30, 'C'])

    def test_dictlist_success_base_method(self):
        """正常系 (基底メソッド)"""

        self.dict_list.set_values(values=(10, 20, 30))
        self.dict_list.set_values(values=('A', 'B', 'C'))

        self.assertIsInstance(
            list(self.dict_list) + list(self.dict_list), list)

        val1 = {1: 111, 2: 222, 3: 333}
        self.dict_list.append(item=val1)
        self.assertEqual(self.dict_list[2], val1)

        self.assertEqual(self.dict_list.count(1), 0)

        val2 = {1: 'a', 2: 'b', 3: 'c'}
        self.dict_list.extend([val2])
        self.assertEqual(self.dict_list[3], val2)

        val3 = {1: 'x', 2: 'y', 3: None}
        self.dict_list.insert(1, val3)
        self.assertEqual(self.dict_list[1], val3)

        val4 = self.dict_list.pop()
        self.assertEqual(val2, val4)
        self.assertEqual(len(self.dict_list), 4)

        self.dict_list.remove({1: 10, 2: 20, 3: 30})
        self.assertEqual(len(self.dict_list), 3)

        self.dict_list.reverse()
        self.assertEqual(self.dict_list[2], val3)

        val5 = {1: 'Hoge', 2: 123, 3: None}
        self.dict_list[2] = val5
        self.assertEqual(self.dict_list[2], val5)

        self.dict_list.clear()
        self.assertEqual(len(self.dict_list), 0)

    def test_dictlist_error_init(self):
        """異常系 (__init__)"""

        with self.assertRaises(Exception):
            DictList()
        with self.assertRaises(Exception):
            DictList(keys={})
        with self.assertRaises(Exception):
            DictList(keys=[])
        with self.assertRaises(Exception):
            DictList(keys=[None])
        with self.assertRaises(Exception):
            DictList(keys=(1, 2, 3))
        with self.assertRaises(Exception):
            DictList(keys={1, 2, 3})
        with self.assertRaises(Exception):
            DictList(keys='1,2,3')
        with self.assertRaises(Exception):
            DictList(keys=['a', 'a'])

    def test_dictlist_error_set_values(self):
        """異常系 (set_values)"""

        with self.assertRaises(Exception):
            self.dict_list.set_values()
        with self.assertRaises(Exception):
            self.dict_list.set_values((10,))
        with self.assertRaises(Exception):
            self.dict_list.set_values((10, 20, ))
        with self.assertRaises(Exception):
            self.dict_list.set_values((10, 20, 30, 40,))
        with self.assertRaises(Exception):
            self.dict_list.set_values({})
        with self.assertRaises(Exception):
            self.dict_list.set_values({1, 2, 3})
        with self.assertRaises(Exception):
            self.dict_list.set_values([])
        with self.assertRaises(Exception):
            self.dict_list.set_values([1, 2, 3])

    def test_dictlist_error_get_values(self):
        """異常系 (get_values)"""

        self.dict_list.set_values(values=(10, 20, 30))
        self.dict_list.set_values(values=('A', 'B', 'C'))

        with self.assertRaises(Exception):
            self.dict_list.get_values()
        with self.assertRaises(Exception):
            self.dict_list.get_values(None)
        with self.assertRaises(Exception):
            self.dict_list.get_values('NotExistKey')
        with self.assertRaises(Exception):
            self.dict_list.get_values(-1)

    def test_dictlist_error_base_method(self):
        """異常系 (基底メソッド)"""

        self.dict_list.set_values(values=(10, 20, 30))
        self.dict_list.set_values(values=('A', 'B', 'C'))

        with self.assertRaises(Exception):
            self.dict_list + self.dict_list
        with self.assertRaises(Exception):
            self.dict_list.append(item={5: 111, 2: 222, 3: 333})
        with self.assertRaises(Exception):
            self.dict_list.copy()
        with self.assertRaises(Exception):
            self.dict_list.extend([{5: 111, 2: 222, 3: 333}])
        with self.assertRaises(Exception):
            self.dict_list.index({1: None, 2: None, 3: None})
        with self.assertRaises(Exception):
            self.dict_list.insert(1, {5: 'a', 2: 'b', 3: None})
        with self.assertRaises(Exception):
            self.dict_list.sort()       # 元々Pythonでdictのlistのsortはエラーになる
        with self.assertRaises(Exception):
            self.dict_list.sort(reverse=True)


def suite():
    """ 指定したメソッド順に実行させるためにメソッドを登録する
    """

    suite = unittest.TestSuite()

    # 正常系
    suite.addTest(TestDictList('test_dictlist_success_simple'))

    suite.addTest(TestDictList('test_dictlist_success_init'))
    suite.addTest(TestDictList('test_dictlist_success_keys'))
    suite.addTest(TestDictList('test_dictlist_success_set_get_values'))
    suite.addTest(TestDictList('test_dictlist_success_base_method'))

    # 異常系
    suite.addTest(TestDictList('test_dictlist_error_init'))
    suite.addTest(TestDictList('test_dictlist_error_set_values'))
    suite.addTest(TestDictList('test_dictlist_error_get_values'))
    suite.addTest(TestDictList('test_dictlist_error_base_method'))

    return suite


if __name__ == '__main__':
    try:
        # unittest.main(verbosity=2, exit=False)
        runner = unittest.TextTestRunner(failfast=True, verbosity=2)
        runner.run(suite())
    except:
        print(traceback.format_exc())
