#!/usr/bin/env python3


class Base():
    BASE_MEM = 'base mem'
    BASE_DEF = 123

    def __init__(self):
        self._base_mem = Base.BASE_MEM

    def base_func(self):
        return self._base_mem

    @classmethod
    def base_class_func(cls):
        return Base.BASE_DEF
