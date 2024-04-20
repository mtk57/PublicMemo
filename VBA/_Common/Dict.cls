VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const VERSION = "1.0.1"

Private dict_ As Object

'コンストラクタ
Private Sub Class_Initialize()
  
  Set dict_ = CreateObject("Scripting.Dictionary")
  
End Sub

'キーの存在有無を返す
Public Function IsExists(ByVal key As Variant) As Boolean
    IsExists = dict_.Exists(key)
End Function

'キーに紐づく値を追加する
Public Sub AppendValue(ByVal key As Variant, ByVal val As Variant)
    If IsNull(key) Then
        Err.Raise 53, , "キーが空です"
    End If
    
    If dict_.Exists(key) = False Then
        dict_.Add key, val
    Else
        dict_(key) = val
    End If
End Sub

'キーに紐づくString配列に値を追加する
Public Sub AppendStringArray(ByVal key As Variant, ByVal val As String)
    If IsNull(key) Then
        Err.Raise 53, , "キーが空です"
    End If

    Dim i As Long
    Dim wk_ary() As String
    
    If dict_.Exists(key) = False Then
        ReDim Preserve wk_ary(0)
        wk_ary(0) = val
    
        dict_.Add key, wk_ary
    Else
        Dim cnt As Long
        cnt = UBound(dict_(key)) + 1
        
        wk_ary = dict_(key)
        
        ReDim Preserve wk_ary(cnt)
        
        wk_ary(cnt) = val
        
        dict_(key) = wk_ary
    End If
End Sub

'キーに紐づく値を返す
Public Function GetValue(ByVal key As Variant) As Variant
    If IsNull(key) Then
        Err.Raise 53, , "キーが空です"
    End If
    
    GetValue = dict_(key)
End Function

'キー数を返す
Public Function GetKeyCount() As Long
    GetKeyCount = UBound(dict_.Keys)
End Function

'キーの一覧を返す
Public Function GetKeys() As Variant()
    GetKeys = dict_.Keys
End Function

'値の一覧を返す
Public Function GetValues() As Variant()
    GetValues = dict_.Items
End Function

'String配列のコピーを返す(ボツ)
'Public Function CopyStringArray(ByRef dst_dict As Dict) As Boolean
'    Set dst_dict = New Dict
'
'    Dim i As Long
'    Dim j As Long
'
'    For i = 0 To GetKeyCount()
'        Dim key As String: key = GetKeys()(i)
'
'        Dim values() As String
'        values = GetValue(key)
'
'        For j = 0 To UBound(values)
'            Dim value As String: value = values(j)
'            dst_dict.AppendStringArray key, value
'        Next j
'    Next i
'    CopyStringArray = True
'End Function

