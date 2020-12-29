#!/usr/bin/env python3

import pytest
from ..src import child


def test_1():
    ins_c = child.Child()

    assert ins_c.child_func() == 'child mem' + '&' + 'ABC'
    assert child.Child.child_class_func() == 456
    assert child.Child.child_class_func(prm1=False) is None
    assert child.Child_module_func() == 1230


def test_2_child_init_mock(monkeypatch):
    """
    Childクラスの__init__をモックする
    """
    class Child():
        """ Childクラスのモック """
        def __init__(*args, **kwargs):
            raise RuntimeError('init failed!')

    def mock_Child(*args, **kwargs):
        """ Childクラスのインスタンスを生成 """
        return Child()

    # Childのクラスのインスタンス生成をモック
    monkeypatch.setattr(child, 'Child', mock_Child)

    with pytest.raises(RuntimeError):
        # Childのクラスのインスタンス生成
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


def test_6_child_class_member_mock():
    """
    Childクラスのクラスメンバーをモックする
    """
    pass


def test_100_base_init_mock(monkeypatch):
    """
    Baseクラスの__init__をモックする
    """
    pass


def test_101_base_instance_method_mock():
    """
    Baseクラスのインスタンスメソッドをモックする
    """
    pass


def test_102_base_class_method_mock():
    """
    Baseクラスのクラスメソッドをモックする
    """
    pass


def test_103_base_instance_member_mock():
    """
    Baseクラスのインスタンスメンバーをモックする
    """
    pass


def test_104_base_class_member_mock():
    """
    Baseクラスのクラスメンバーをモックする
    """
    pass


def test_1000_module_method_mock():
    """
    モジュールメソッドをモックする
    """
    pass


def test_1001_module_varialble_mock():
    """
    モジュール変数をモックする
    """
    pass
