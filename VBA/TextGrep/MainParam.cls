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

Private Const INPUT_FILENAME = "input.tsv"
Private Const OUTPUT_FILENAME = "output.tsv"

'メンバー変数

'GREP.exe
Private grep_exe As String

'GREPフォルダパス
Private grep_dir_path As String

'出力シート名
Private output_sheet_name As String

'デバッグログ
Private debug_log As Boolean
Private debug_log_s As String

'サブフォルダも含む
Private contain_subdir As Boolean
Private contain_subdir_s As String

'対象ファイル
Private target_file As String

'正規表現
Private regular_expression As Boolean
Private regular_expression_s As String

'大文字小文字区別
Private ignorecase As Boolean
Private ignorecase_s As String

'0件も出力する
Private output_zero As Boolean
Private output_zero_s As String

Private Const CLM_ = "J"
Private Const ROW_ = 7

'コンストラクタ
Private Sub Class_Initialize()
    grep_exe = ""
    grep_dir_path = ""
    output_sheet_name = ""
    debug_log = False
    debug_log_s = ""
    contain_subdir = False
    contain_subdir_s = ""
    target_file = ""
    regular_expression = False
    regular_expression_s = ""
    ignorecase = False
    ignorecase_s = ""
    output_zero = False
    output_zero_s = ""
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    grep_exe = main_sheet.Range(CLM_ & ROW_).value
    grep_dir_path = main_sheet.Range(CLM_ & ROW_ + 1).value
    output_sheet_name = main_sheet.Range(CLM_ & ROW_ + 2).value
    debug_log_s = main_sheet.Range(CLM_ & ROW_ + 3).value
    contain_subdir_s = main_sheet.Range(CLM_ & ROW_ + 4).value
    target_file = main_sheet.Range(CLM_ & ROW_ + 5).value
    regular_expression_s = main_sheet.Range(CLM_ & ROW_ + 6).value
    ignorecase_s = main_sheet.Range(CLM_ & ROW_ + 7).value
    output_zero_s = main_sheet.Range(CLM_ & ROW_ + 8).value
End Sub

'検証
Public Sub Validate()
    If grep_exe = "" Then
        Err.Raise 53, , "[MainParam] GREP.exeが指定されていません"
    End If
    
    If Common.IsExistsFile(GetGrepExe()) = False Then
        Err.Raise 53, , "[MainParam] GREP.exeが存在しません"
    End If
    
    If output_sheet_name = "" Then
        output_sheet_name = Common.GetNowTimeString()
    End If
    
    If grep_dir_path = "" Then
        Err.Raise 53, , "[MainParam] GREPフォルダパスが指定されていません"
    End If
    
    If Common.IsExistsFolder(grep_dir_path) = False Then
        Err.Raise 53, , "[MainParam] GREPフォルダパスが存在しません"
    End If
    
    If debug_log_s = "" Or debug_log_s = "NO" Then
        debug_log = False
    Else
        debug_log = True
    End If
    
    If contain_subdir_s = "" Or contain_subdir_s = "NO" Then
        contain_subdir = False
    Else
        contain_subdir = True
    End If
    
    If target_file = "" Then
        Err.Raise 53, , "[MainParam] 対象ファイルが空です。"
    End If

    If regular_expression_s = "NO" Or regular_expression_s = "" Then
        regular_expression = False
    Else
        regular_expression = True
    End If
    
    If ignorecase_s = "NO" Or ignorecase_s = "" Then
        ignorecase = False
    Else
        ignorecase = True
    End If
    
    If output_zero_s = "NO" Or output_zero_s = "" Then
        output_zero = False
    Else
        output_zero = True
    End If
End Sub

Public Sub CreateInputFile(ByRef sub_params() As SubParam)
    Dim fso As Object
    Dim ts As Object
    Dim row As Range
    Dim cell As Range
    
    Dim path As String: path = GetInputFilePath()
    
    If Common.IsExistsFile(path) = True Then
        Common.DeleteFile path
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(path, False, False)
    
    'BOM
    'ts.Write Chr(239) & Chr(187) & Chr(191)
    
    ts.Write "GrepDirPath" & vbTab & GetGrepDirPath() & vbNewLine
    ts.Write "IsDebugLog" & vbTab & IsDebugLogInt() & vbNewLine
    ts.Write "IsSubDir" & vbTab & IsSubDirInt() & vbNewLine
    ts.Write "TargetFile" & vbTab & GetTargetFile() & vbNewLine
    ts.Write "IsRegEx" & vbTab & IsRegExInt() & vbNewLine
    ts.Write "IsIgnoreCase" & vbTab & IsIgnorecaseInt() & vbNewLine
    ts.Write "IsOutputZeroCount" & vbTab & IsOutputZeroCountInt() & vbNewLine
    
    Dim i As Long
    For i = 0 To UBound(sub_params)
        Dim sub_param As SubParam
        Set sub_param = sub_params(i)
        
        ts.Write "Keywords" & vbTab & sub_param.GetKeyword() & vbNewLine
    Next i
    
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    
    'UTF8(BOM)に変換
    Common.SJIStoUTF8 path, False
End Sub

'=========================================
'Getter
'=========================================

Public Function GetGrepExe() As String
    GetGrepExe = ActiveWorkbook.path & Application.PathSeparator & grep_exe
End Function

Public Function GetInputFilePath() As String
    GetInputFilePath = ActiveWorkbook.path & Application.PathSeparator & INPUT_FILENAME
End Function

Public Function GetOutputFilePath() As String
    GetOutputFilePath = ActiveWorkbook.path & Application.PathSeparator & OUTPUT_FILENAME
End Function

Public Function GetGrepDirPath() As String
    GetGrepDirPath = grep_dir_path
End Function

Public Function GetOutputSheetName() As String
    GetOutputSheetName = output_sheet_name
End Function

Public Function IsDebugLogInt() As Integer
    If debug_log = True Then
        IsDebugLogInt = 1
    Else
        IsDebugLogInt = 0
    End If
End Function

Public Function IsSubDir() As Boolean
    IsSubDir = contain_subdir
End Function

Public Function IsSubDirInt() As Integer
    If contain_subdir = True Then
        IsSubDirInt = 1
    Else
        IsSubDirInt = 0
    End If
End Function

Public Function GetTargetFile() As String
    GetTargetFile = target_file
End Function

Public Function IsRegEx() As Boolean
    IsRegularExpression = regular_expression
End Function

Public Function IsRegExInt() As Integer
    If regular_expression = True Then
        IsRegExInt = 1
    Else
        IsRegExInt = 0
    End If
End Function

Public Function IsIgnorecase() As Boolean
    IsIgnorecase = ignorecase
End Function

Public Function IsIgnorecaseInt() As Integer
    If ignorecase = True Then
        IsIgnorecaseInt = 1
    Else
        IsIgnorecaseInt = 0
    End If
End Function

Public Function IsOutputZeroCount() As Boolean
    IsOutputZeroCount = output_zero
End Function

Public Function IsOutputZeroCountInt() As Integer
    If output_zero = True Then
        IsOutputZeroCountInt = 1
    Else
        IsOutputZeroCountInt = 0
    End If
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "grep_exe=(" & grep_exe & "), " & _
          "grep_dir_path=(" & grep_dir_path & "), " & _
          "output_sheet_name=(" & output_sheet_name & "), " & _
          "debug_log=(" & debug_log_s & "), " & _
          "contain_subdir=(" & contain_subdir_s & "), " & _
          "target_file=(" & target_file & "), " & _
          "regular_expression=(" & regular_expression_s & "), " & _
          "ignorecase=(" & ignorecase_s & "), " & _
          "output_zero=(" & output_zero_s & ")"

    GetAllValue = ret
End Function


