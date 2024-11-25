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
Private ignore_comment As Boolean
Private enabled_callstack As Boolean
Private max_depth_ As Long
Private vbp_ref_list_sheet As String

Private format_type_s As String
Private lang_s As String
Private ignore_comment_s As String
Private enabled_callstack_s As String
Private max_depth_s As String

Private Const CLM_ = "L"
Private Const ROW_ = 8
Private Const MAX_DEPTH = 10

'コンストラクタ
Private Sub Class_Initialize()
    format_type = GrepAppEnum.sakura
    lang = LangEnum.VB6
    ignore_comment = True
    enabled_callstack = False
    max_depth_ = 0
    vbp_ref_list_sheet = ""
    
    format_type_s = ""
    lang_s = ""
    enabled_callstack_s = ""
    max_depth_s = ""
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    format_type_s = main_sheet.Range(CLM_ & ROW_).value
    lang_s = main_sheet.Range(CLM_ & ROW_ + 1).value
    ignore_comment_s = main_sheet.Range(CLM_ & ROW_ + 2).value
    enabled_callstack_s = main_sheet.Range(CLM_ & ROW_ + 3).value
    max_depth_s = main_sheet.Range(CLM_ & ROW_ + 4).value
    vbp_ref_list_sheet = main_sheet.Range(CLM_ & ROW_ + 5).value
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
    
    If ignore_comment_s = "" Or ignore_comment_s = "YES" Then
        ignore_comment = True
    Else
        ignore_comment = False
    End If
    
    If enabled_callstack_s = "" Or enabled_callstack_s = "NO" Then
        enabled_callstack = False
    Else
        enabled_callstack = True
    End If
    
    If enabled_callstack = False Then
        'コールスタック出力が無しなのでこれ以降のチェックはしない
        Exit Sub
    End If
    
    If max_depth_s = "" Then
        max_depth_s = CStr(MAX_DEPTH)
    End If
    
    If IsNumeric(max_depth_s) = False Then
        Err.Raise 53, , "最大スタック数は1以上の整数値で指定してください。 (" & max_depth_s & ")"
    End If
    
    max_depth_ = CLng(max_depth_s)
    
    If max_depth_ < 1 Or max_depth_ > MAX_DEPTH Then
        Err.Raise 53, , "最大スタック数は1～" & MAX_DEPTH & "の範囲で指定してください。 (" & max_depth_s & ")"
    End If
    
    If vbp_ref_list_sheet = "" Then
        Err.Raise 53, , "VBP参照ファイル一覧のシート名を指定してください。"
    End If
    
    If Common.IsExistSheet(ActiveWorkbook, vbp_ref_list_sheet) = False Then
        Err.Raise 53, , "VBP参照ファイル一覧のシートが存在しません。(" & vbp_ref_list_sheet & ")"
    End If
    
    Dim target_sheet As Worksheet
    Set target_sheet = ThisWorkbook.Sheets(vbp_ref_list_sheet)
    
    If target_sheet.Range("A1").value = "" Or target_sheet.Range("B1").value = "" Then
        Err.Raise 53, , "VBP参照ファイル一覧のシートにデータが存在しません。(" & vbp_ref_list_sheet & ")"
    End If
    
    Set target_sheet = Nothing
End Sub

'=========================================
'Getter
'=========================================

Private Sub Sub1_bas()
    'Dim hoge● As String
End Sub


Public Function GetFormatType() As GrepAppEnum
    GetFormatType = format_type
End Function

Public Function GetLang() As LangEnum
    GetLang = lang
End Function

Public Function IsIgnoreComment() As Boolean
    IsIgnoreComment = ignore_comment
End Function

Public Function IsEnabledCallstack() As Boolean
    IsEnabledCallstack = enabled_callstack
End Function

Public Function GetMaxDepth() As Long
    GetMaxDepth = max_depth_
End Function

Public Function GetVbpRefListSheetName() As String
    GetVbpRefListSheetName = vbp_ref_list_sheet
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "format_type=(" & format_type_s & "), " & _
          "lang=(" & lang_s & "), " & _
          "ignore_comment=(" & ignore_comment_s & "), " & _
          "enabled_callstack=" & enabled_callstack & "), " & _
          "max_depth=(" & max_depth_ & "), " & _
          "vbp_ref_list_sheet=(" & vbp_ref_list_sheet & ")"

    GetAllValue = ret
End Function
