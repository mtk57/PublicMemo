import traceback
from collections import UserList

"""
空飛ぶPython 即時開発指南書に載ってるコード
"""

class TypedUserList(UserList):
    def __init__(self, elm, init_list=[]):
        super().__init__(init_list)
        self._type = type(elm)
        if not isinstance(init_list, list):
            raise TypeError('2nd arg is not list!')
        for e in init_list:
            self._check(e)
    
    def _check(self, elm):
        if type(elm) != self._type:
            raise TypeError('!!')
    
    def __setitem__(self, i, elm):
        self._check(elm)
        self.data[i] = elm
    
    def __getitem__(self, i):
        return self.data[i]

if __name__ == '__main__':
    try:
        x = TypedUserList('', ['']*5)
        x[2] = 'Hello'
        x[3] = 'There'
        print(x[2] + ' ' + x[3])

        a, b, c, d, e = x
        print(a, b, c, d)

    except:
        print(traceback.format_exc())