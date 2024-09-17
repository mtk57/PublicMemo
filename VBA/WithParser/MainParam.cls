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
Private check_len As Integer
Private check_len_s As String
Private is_confirm As Boolean
Private is_confirm_s As String

Private Const CLM_ = "L"
Private Const ROW_ = 7

'コンストラクタ
Private Sub Class_Initialize()
    format_type = ""
    check_len = 0
    is_confirm = True
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    format_type = main_sheet.Range(CLM_ & ROW_).value
    check_len_s = main_sheet.Range(CLM_ & ROW_ + 1).value
    is_confirm_s = main_sheet.Range(CLM_ & ROW_ + 2).value
End Sub

'検証
Public Sub Validate()
    If format_type = "" Or format_type = "sakura" Then
        format_type = "sakura"
    Else
        Err.Raise 53, , "未サポートのGREP結果フォーマットです (" & format_type & ")"
    End If
    
    If check_len_s = "" Or check_len_s = "なし" Then
        check_len = 0
    ElseIf check_len_s = "前後1文字" Then
        check_len = 1
    ElseIf check_len_s = "前後2文字" Then
        check_len = 2
    Else
        check_len = 0
    End If
    
    If is_confirm_s = "" Or is_confirm_s = "YES" Then
        is_confirm = True
    Else
        is_confirm = False
    End If
End Sub

'=========================================
'Getter
'=========================================

Public Function GetFormatType() As String
    GetFormatType = format_type
End Function

Public Function GetCheckLen() As Integer
    GetCheckLen = check_len
End Function

Public Function IsConfirm() As Boolean
    IsConfirm = is_confirm
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "format_type=(" & format_type & ")," & _
          "check_len=(" & check_len & "), " & _
          "is_confirm=(" & is_confirm_s & ")"
          

    GetAllValue = ret
End Function
