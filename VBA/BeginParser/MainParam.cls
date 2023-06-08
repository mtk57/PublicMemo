VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MainParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'メンバー変数

Private format_type As String

Private Const CLM_ = "L"
Private Const ROW_ = 7

'コンストラクタ
Private Sub Class_Initialize()
    format_type = ""
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    format_type = main_sheet.Range(CLM_ & ROW_).value
End Sub

'検証
Public Sub Validate()
    If format_type = "" Or format_type = "sakura" Then
        format_type = "sakura"
    Else
        Err.Raise 53, , "未サポートのGREP結果フォーマットです (" & format_type & ")"
    End If
End Sub

'=========================================
'Getter
'=========================================

Public Function GetFormatType() As String
    GetFormatType = format_type
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "format_type=(" & format_type & ")"

    GetAllValue = ret
End Function
