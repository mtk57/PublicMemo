#!/usr/bin/env python3

import pytest
# from unittest import mock
from ..src import base, child


def test_1():
    # pytest.set_trace()  # pdb kick

    # Childのクラスのインスタンス生成
    ins_c = child.Child()

    # Baseクラスのインスタンスメソッドを呼び出す
    assert ins_c.base_func() == base.Base.BASE_MEM

    # Baseクラスのクラスメソッドを呼び出す
    assert base.Base.base_class_func() == base.Base.BASE_DEF

    # Childクラスのインスタンスメソッドを呼び出す
    assert ins_c.child_func() == child.Child.CHILD_MEM + '&' + child.Child.DEF_PRM1

    # Childクラスのクラスメソッドを呼び出す
    assert child.Child.child_class_func() == child.Child.CHILD_DEF

    # Childクラスのクラスメソッドを呼び出す2
    assert child.Child.child_class_func(prm1=False) is None

    # chileモジュールのモジュールメソッドを呼び出す
    assert child.Child_module_func() == 1230


def test_2_child_init_mock(monkeypatch):
    """
    Childクラスの__init__をモックする
    """
    class Child():
        """ Childクラスのモック """
        def __init__(*args, **kwargs):
            raise RuntimeError('init failed!')

    def mock_child(*args, **kwargs):
        """ Childクラスのインスタンスを生成 """
        return Child()

    # Childクラスのインスタンス生成をモック
    monkeypatch.setattr(child, 'Child', mock_child)

    with pytest.raises(RuntimeError):
        # Childクラスのインスタンス生成
        child.Child()


def test_3_child_instance_method_mock(monkeypatch):
    """
    Childクラスのインスタンスメソッドをモックする
    """
    # Childのクラスのインスタンス生成
    ins_c = child.Child()

    def mock_child_func(*args, **kwargs):
        """ Childクラスのchild_funcメソッドのモック """
        return 'mock return'

    # child_funcメソッドのモックを設定
    monkeypatch.setattr(ins_c, 'child_func', mock_child_func)

    # pytest.set_trace()  # pdb kick

    # child_funcメソッドを呼び出す
    assert ins_c.child_func() == 'mock return'


def test_4_child_class_method_mock(monkeypatch):
    """
    Childクラスのクラスメソッドをモックする
    """
    def mock_child_class_func(*args, **kwargs):
        """ Childクラスのchild_class_funcメソッドのモック """
        return 'mock return'

    # child_class_funcメソッドのモックを設定
    monkeypatch.setattr(child.Child, 'child_class_func', mock_child_class_func)

    # pytest.set_trace()  # pdb kick

    # child_class_funcメソッドを呼び出す
    assert child.Child.child_class_func() == 'mock return'


def test_5_child_instance_member_mock(monkeypatch):
    """
    Childクラスのインスタンスメンバーをモックする
    """
    # Childのクラスのインスタンス生成
    ins_c = child.Child()

    # _child_memインスタンスメンバの値を設定
    monkeypatch.setitem(ins_c.__dict__, '_child_mem', 'mock value')

    # pytest.set_trace()  # pdb kick

    assert ins_c._child_mem == 'mock value'


def test_6_child_class_member_mock(monkeypatch):
    """
    Childクラスのクラスメンバーをモックする (無理かも...)
    """
    pass
    # CHILD_DEFインスタンスメンバの値を設定
    # monkeypatch.setitem(child.Child.__dict__, 'CHILD_DEF', 'mock value')

    # assert child.Child.CHILD_DEF == 'mock value'


def test_1000_module_method_mock(monkeypatch):
    """
    モジュールメソッドをモックする
    """
    def mock_Child_module_func(*args, **kwargs):
        """ Child_module_funcメソッドのモック """
        return 'mock return'

    # Child_module_funcメソッドのモックを設定
    monkeypatch.setattr(child, 'Child_module_func', mock_Child_module_func)
    assert child.Child_module_func() == 'mock return'


def test_1001_module_varialble_mock():
    """
    モジュール変数をモックする
    """
    pass
