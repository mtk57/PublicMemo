VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'キーワード
Private keyword As String

'コンストラクタ
Private Sub Class_Initialize()
    keyword = ""
End Sub

'初期化
Public Sub Init(ByVal row As Integer)
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("main")
    
    keyword = sheet.Range("B" & row).value
End Sub

'=========================================
'Getter
'=========================================
Public Function GetKeyword() As String
    GetKeyword = keyword
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[SubParam] " & _
          "keyword=(" & keyword & ")"

    GetAllValue = ret
End Function

