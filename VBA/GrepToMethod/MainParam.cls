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

Private format_type As GrepAppEnum
Private lang As LangEnum

Private format_type_s As String
Private lang_s As String

Private Const CLM_ = "L"
Private Const ROW_ = 8

'コンストラクタ
Private Sub Class_Initialize()
    format_type = GrepAppEnum.sakura
    lang = LangEnum.VB6
    
    format_type_s = ""
    lang_s = ""
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    format_type_s = main_sheet.Range(CLM_ & ROW_).value
    lang_s = main_sheet.Range(CLM_ & ROW_ + 1).value
End Sub

'検証
Public Sub Validate()
    If format_type_s = "" Or format_type_s = "sakura" Then
        format_type = GrepAppEnum.sakura
    Else
        Err.Raise 53, , "未サポートのGREP結果フォーマットです (" & format_type_s & ")"
    End If

    If lang_s = "" Or lang_s = "VB6" Then
        lang = LangEnum.VB6
    Else
        Err.Raise 53, , "未サポートの言語です (" & lang_s & ")"
    End If
End Sub

'=========================================
'Getter
'=========================================

Public Function GetFormatType() As GrepAppEnum
    GetFormatType = format_type
End Function

Public Function GetLang() As LangEnum
    GetLang = lang
End Function

Public Function GetAllValue() As String
    Dim Ret As String
    
    Ret = "[Main Param] " & _
          "format_type=(" & format_type_s & "), " & _
          "lang=(" & lang_s & ")"

    GetAllValue = Ret
End Function
