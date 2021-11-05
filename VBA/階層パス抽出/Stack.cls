VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
'https://kazusa-pg.com/vba-stack/

Private ary() As Variant   '配列をスタックとして利用
Private size As Long       'スタックの現在のサイズ
 
Private Sub Class_Initialize()
  'コンストラクタ
  
  ReDim ary(0)
  size = 0
  
End Sub
Public Sub push(v As Variant)
  'スタックの一番上にデータを追加する
  
  size = UBound(ary) + 1
  ReDim Preserve ary(size)
  
  ary(size) = v
  
End Sub

Public Function peek() As Variant
  'スタックの一番上のデータを読み取る
  
  If size = 0 Then
    'スタックにデータが存在しないときは、エラー番号1002をRaiseする
    Err.Raise 1002, "Stack", "スタックにデータが存在しません"
  End If
  
  peek = ary(size)
  
End Function

Public Function pop() As Variant
  'スタックの一番上のデータを取り出す
  
  If size = 0 Then
    'スタックにデータが存在しないときは、エラー番号1001をRaiseする
    Err.Raise 1001, "Stack", "スタックにデータが存在しません"
  End If
  
  pop = ary(size)
  
  size = size - 1
  ReDim Preserve ary(size)
  
End Function
Public Function count() As Long
  'スタックのデータ数を返す
  
  count = size
 
End Function
Public Function getContents() As Variant
  'スタックの内容を配列として返す
  'インデックス1が一番下のデータ
  
  getContents = ary
 
End Function
