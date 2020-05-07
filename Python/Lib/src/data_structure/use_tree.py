#!/usr/bin/python3
from data_structure.tree.element import TreeElement

# ---------------------------------------------------
# TreeElementの使用サンプル
# ---------------------------------------------------


# メイン

# ルートノードを作成
root = TreeElement()

# データを追加
root.append('1-A')
root.append('1-B')
root.append('1-C')

# 子ノード(第2階層)を追加
child_1 = root.append_child(id='child-1')

# データを追加
child_1.append('2-A')
child_1.append('2-B')

# 子ノード(第2階層)に子ノード(第3階層)を追加
child_2 = child_1.append_child(id='child-2')
# データを追加
child_2.append('3-A')
child_2.append('3-B')

# データを追加
child_1.append('2-C')

# データを追加
root.append('1-D')

# 現在のツリーをツリー表示
"""
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
print(f"\nツリー\n{root.to_tree_string()}\n")

# 検索キーでノードを検索
find_node = TreeElement.search('3-B', root)

# 検索キーを含むノードを削除
find_node.get_parent().delete_child()

# 現在のツリーをツリー表示
"""
root
+ 1-A
+ 1-B
+ 1-C
  child-1
  + 2-A
  + 2-B
  + 2-C
+ 1-D
"""
print(f"\n削除後\n{root.to_tree_string(show_id=True)}\n")

# 現在のツリーをリスト表示
"""
[ 1-A, 1-B, 1-C, [ 2-A, 2-B, 2-C ], 1-D ]
"""
print(f'\nリスト\n{root.to_list_string()}\n')


sample = TreeElement(id='brick')
sample.append_child(id='public').append_child(id='tmp')
sample.append(data='api.json')
print(f'{sample.to_tree_string(show_id=True)}')


print('SUCCESS!')
