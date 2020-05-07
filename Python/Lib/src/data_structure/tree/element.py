#!/usr/bin/env python

import uuid


class TreeElement():
    """
    ツリー構造の要素を表現するデータモデル
    要素（ノード）には、データ(object)または子ノードを格納できる
    """

    def __init__(self, id: str = None, parent=None, parent_ptr=None):
        """
        初期化
        @param  id  ID
        @param	parent  親ノードのオブジェクト
        @param	parent_ptr 親ノードのポインタ
        @return	なし
        """
        # データを格納するリスト
        self._items = []

        self._id = id if (id is not None) else self.create_id()
        self._pointer = 0
        self._parent = parent
        self._parent_ptr = parent_ptr

    def __repr__(self):
        return f'_items={self._items}\n_pointer={self._pointer}\n' \
               f'_parent={self._parent}\n_parent_ptr={self._parent_ptr}\n'

    @classmethod
    def search(cls, key: object, node: object):
        """
        検索キーを含むノードを探す
        @param	object key 検索キー
        @param	TreeElement node 探索対象ノード
        @return	TreeElement 検索キーが存在するノード
        """
        find_node = None

        for data in node.get_data():

            if (node.is_node(data)):
                # 子ノードの場合は再帰
                find_node = TreeElement.search(key, data)

                if (find_node is not None):
                    return find_node

            elif (data == key):
                return node

        return find_node

    def is_node(self, data: object) -> bool:
        """
        データはノードオブジェクトか否か
        @param  nothing
        @return bool
        """
        return isinstance(data, TreeElement)

    def get_data(self) -> list:
        """
        データをリストで返す
        @param nothing
        @return list
        """
        return self._items

    def exist(self) -> bool:
        """
        データが存在するか否か
        @param	なし
        @return	bool
        """
        i = self._pointer
        return True if ((i >= 0) and (self._items[i] is not None)) else False

    def get(self) -> object:
        """
        データを返す
        @param	なし
        @return	object
        """
        i = self._pointer
        return self._items[i] if self.exist() else None

    def append(self, data: object) -> int:
        """
        データを追加する
        @param	object data データ
        @return	int
        """
        self._items.append(data)
        self._pointer = len(self._items) - 1
        return self._pointer

    def delete(self):
        """
        データを削除する
        @param	なし
        @return	なし
        """
        i = self._pointer
        if (self._items[i] is not None):
            self._items.pop(i)

    def append_child(self, id: str = None):
        """
        子ノードを末尾に追加する
        @param	id  ID
        @return	TreeElement
        """
        child = TreeElement(id=id, parent=self, parent_ptr=self._pointer)
        pointer = self.append(child)
        child._parent_ptr = pointer
        return child

    def delete_child(self):
        """
        子ノードを削除する
        @param	なし
        @return	なし
        """
        if (self.is_node(self.get())):
            self.delete()

    def get_parent(self):
        """
        親ノードを返す
        @param	なし
        @return	TreeElement or None
        """
        ret = None
        if (self._parent is not None):
            self._parent._pointer = self._parent_ptr
            ret = self._parent
        return ret

    def create_id(self):
        """
        IDを生成する
        @param なし
        @return ID
        """
        return uuid.uuid4()

    def to_list_string(self) -> str:
        """
        ツリーのデータをリスト型の文字列に成形して返す
        @param	なし
        @return	str
        例：
            [ 1-A, 1-B, 1-C, [ 2-A, 2-B, 2-C ], 1-D ]
        """
        outstr = '[ '

        # ツリー末尾まで繰り返し
        for data in self.get_data():
            if outstr != '[ ':
                outstr += ', '

            if (self.is_node(data)):
                # 子ノードの場合は再帰
                outstr += data.to_list_string()
            else:
                outstr += str(data)

        return outstr + ' ]'

    def to_tree_string(self, show_id: bool = False, depth: int = 0) -> str:
        """
        ツリーのデータを階層型の文字列に成形して返す
        @param	int depth 階層の深さ
        @return	str
        例：
            root
            + 1-A
            + 1-B
            + 1-C
            child-1
            + 2-A
            + 2-B
                child-2
                + 3-A
                + 3-B
            + 2-C
            + 1-D
        """
        id = self._id if (show_id is True) else ''

        if (depth == 0):
            outstr = f'{id}\n'
        else:
            outstr = '  ' * depth
            outstr += f'|{id}\n'

        # ツリー末尾まで繰り返し
        for data in self.get_data():

            if (self.is_node(data)):
                # 子ノードの場合は再帰
                outstr += data.to_tree_string(
                            show_id=show_id, depth=depth + 1)
            else:
                outstr += '  ' * depth
                outstr += '+ ' + str(data) + '\n'

        return outstr
