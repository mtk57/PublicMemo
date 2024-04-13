VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CopyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'メンバー変数
Private key As String
Private value As String
Private is_copied As Boolean

'コンストラクタ
Private Sub Class_Initialize()
    key = ""
    value = ""
    is_copied = False
End Sub

'初期化
Public Sub Init(ByVal in_key As String, ByVal in_value As String, Optional ByVal copied As Boolean = False)
    key = in_key
    value = in_value
    is_copied = copied
End Sub

'=========================================
'Getter
'=========================================

'-----------------------------------------
'検索キー
'-----------------------------------------
Public Function GetKey() As String
    GetKey = key
End Function

'-----------------------------------------
'検索値
'-----------------------------------------
Public Function GetValue() As String
    GetValue = value
End Function

Public Function IsCopies() As Boolean
    IsCopies = is_copied
End Function


Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Copy Data] " & _
          "key=(" & key & "), " & _
          "value=(" & value & "), " & _
          "is_copied=(" & is_copied & ")"
    GetAllValue = ret
End Function



