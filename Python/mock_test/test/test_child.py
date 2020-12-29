#!/usr/bin/env python3

import pytest
from ..src import child


def test_1():
    ins_c = child.Child()
    ins_c.child_func()


def test_2_init_mock(monkeypatch):
    """
    Childクラスの__init__をモックする
    """
    class Child():
        def __init__(*args, **kwargs):
            raise RuntimeError('init failed!')

    def mock_Child(*args, **kwargs):
        return Child()

    monkeypatch.setattr(child, 'Child', mock_Child)

    with pytest.raises(RuntimeError):
        child.Child()
