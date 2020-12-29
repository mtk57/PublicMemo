#!/usr/bin/env python3

from .base import Base


CHILD_MODULE_DEF = 9999


def Child_module_func(prm1=123):
    return prm1 * 10


class Child(Base):
    CHILD_DEF = 456

    def __init__(self, prm1=123):
        self._child_mem = 'child mem'
        self._child_mem_prm1 = None

        if prm1 == 123:
            self._child_mem_prm1 = prm1

    def child_func(self, prm1='ABC'):
        return self._child_mem + '&' + prm1

    @classmethod
    def child_class_func(cls, prm1=True):
        if prm1 is True:
            return Child.CHILD_DEF
        return None
