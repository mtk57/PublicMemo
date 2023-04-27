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

'MUST
Private exe_file_path As String
Private src_dir_path As String
Private dst_dir_path As String
Private in_ext As String
Private addin_file_path As String
Private exe_ini_file_path As String

'WANT
Private is_contain_sub_dir As Boolean
Private ignore_files As String
Private tool_wk_dir_path As String
Private is_step_wk_dir As Boolean
Private is_del_wk_dir As Boolean
Private max_exec_cnt As Integer

Private is_contain_sub_dir_s As String
Private is_step_wk_dir_s As String
Private is_del_wk_dir_s As String
Private max_exec_cnt_s As String

Private ignore_file_path As String

'コンストラクタ
Private Sub Class_Initialize()
    'MUST
    exe_file_path = ""
    src_dir_path = ""
    dst_dir_path = ""
    in_ext = ""
    addin_file_path = ""
    exe_ini_file_path = ""
    
    'WANT
    is_contain_sub_dir = False
    ignore_files = ""
    tool_wk_dir_path = ""
    is_step_wk_dir = False
    is_del_wk_dir = True
    max_exec_cnt = 0
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Const CLM = "N"
    Const i = 6
    
    'MUST
    exe_file_path = main_sheet.Range(CLM & i).value
    src_dir_path = main_sheet.Range(CLM & i + 1).value
    dst_dir_path = main_sheet.Range(CLM & i + 2).value
    in_ext = main_sheet.Range(CLM & i + 3).value
    addin_file_path = main_sheet.Range(CLM & i + 4).value
    exe_ini_file_path = main_sheet.Range(CLM & i + 5).value
    
    'WANT
    is_contain_sub_dir_s = main_sheet.Range(CLM & i + 6).value
    ignore_files = main_sheet.Range(CLM & i + 7).value
    tool_wk_dir_path = main_sheet.Range(CLM & i + 8).value
    is_step_wk_dir_s = main_sheet.Range(CLM & i + 9).value
    is_del_wk_dir_s = main_sheet.Range(CLM & i + 10).value
    max_exec_cnt_s = main_sheet.Range(CLM & i + 11).value
End Sub

'検証
Public Sub Validate()
    'MUST
    If exe_file_path = "" Then
        Err.Raise 53, , "exeファイルパスが指定されていません"
    ElseIf Common.IsExistsFile(exe_file_path) = False Then
        Err.Raise 53, , "指定されたexeファイルが存在しません (" & exe_file_path & ")"
    End If

    If src_dir_path = "" Then
        Err.Raise 53, , "SRCフォルダパスが指定されていません"
    ElseIf Common.IsExistsFolder(src_dir_path) = False Then
        Err.Raise 53, , "指定されたSRCフォルダが存在しません (" & src_dir_path & ")"
    End If

    If dst_dir_path = "" Then
        Err.Raise 53, , "DSTフォルダパスが指定されていません"
    End If

    If in_ext = "" Then
        Err.Raise 53, , "拡張子が指定されていません"
    End If

    If addin_file_path = "" Then
        Err.Raise 53, , "アドインファイルパスが指定されていません"
    ElseIf Common.IsExistsFile(addin_file_path) = False Then
        Err.Raise 53, , "指定されたアドインファイルが存在しません (" & addin_file_path & ")"
    End If

    If exe_ini_file_path = "" Then
        Err.Raise 53, , "ExeIniファイルパスが指定されていません"
    ElseIf Common.IsExistsFile(exe_ini_file_path) = False Then
        Err.Raise 53, , "指定されたExeIniファイルが存在しません (" & exe_ini_file_path & ")"
    End If

    'WANT
    If is_contain_sub_dir_s = "" Or _
       is_contain_sub_dir_s = "NO" Then
       is_contain_sub_dir = False
    Else
        is_contain_sub_dir = True
    End If

    If is_step_wk_dir_s = "" Or _
       is_step_wk_dir_s = "NO" Then
       is_step_wk_dir = False
    Else
        is_step_wk_dir = True
    End If

    If is_del_wk_dir_s = "" Or _
       is_del_wk_dir_s = "YES" Then
       is_del_wk_dir = True
    Else
        is_del_wk_dir = False
    End If

    If max_exec_cnt_s = "" Then
        max_exec_cnt = 1
    ElseIf IsNumeric(max_exec_cnt) Then
        max_exec_cnt = Int(max_exec_cnt_s)
        
        If max_exec_cnt > 10 Then
            max_exec_cnt = 10
        ElseIf max_exec_cnt < 1 Then
            max_exec_cnt = 1
        End If
    Else
        Err.Raise 53, , "最大実行回数が数字ではありません (" & max_exec_cnt_s & ")"
    End If
End Sub

'=========================================
'Getter
'=========================================

'MUST

'-----------------------------------------
'exeファイルパス
'-----------------------------------------
Public Function GetExeFilePath() As String
    GetExeFilePath = exe_file_path
End Function

'-----------------------------------------
'SRCフォルダパス
'-----------------------------------------
Public Function GetSrcDirPath() As String
    GetSrcDirPath = src_dir_path
End Function

'-----------------------------------------
'DSTフォルダパス
'-----------------------------------------
Public Function GetDestDirPath() As String
    GetDestDirPath = dst_dir_path
End Function

'-----------------------------------------
'拡張子
'-----------------------------------------
Public Function GetInExtension() As String
    If IsVB6Extension() = True Then
        GetInExtension = "*.bas"
    Else
        GetInExtension = in_ext
    End If
End Function

'-----------------------------------------
'拡張子はVB6か?
'-----------------------------------------
Public Function IsVB6Extension() As Boolean
    If in_ext = "VB6" Then
        IsVB6Extension = True
    Else
        IsVB6Extension = False
    End If
End Function

'-----------------------------------------
'アドインファイルパス
'-----------------------------------------
Public Function GetAddinFilePath() As String
    GetAddinFilePath = addin_file_path
End Function

'-----------------------------------------
'ExeIniファイルパス
'-----------------------------------------
Public Function GetExeIniFilePath() As String
    GetExeIniFilePath = exe_ini_file_path
End Function



'WANT

'-----------------------------------------
'サブフォルダ含むか?
'-----------------------------------------
Public Function IsContainSubDir() As Boolean
    IsContainSubDir = is_contain_sub_dir
End Function

'-----------------------------------------
'除外ファイルリスト
'-----------------------------------------
Public Function GetIgnoreFiles() As String()
    GetIgnoreFiles = Split(ignore_files, ",")
End Function

Public Function GetIgnoreFilePath() As String
    GetIgnoreFilePath = ignore_file_path
End Function

'-----------------------------------------
'ツール作業用フォルダパス
'-----------------------------------------
Public Function GetToolWorkDirPath() As String
    GetToolWorkDirPath = tool_wk_dir_path
End Function

'-----------------------------------------
'作業用フォルダに途中経過を残すか?
'-----------------------------------------
Public Function IsStepWorkDir() As Boolean
    IsStepWorkDir = is_step_wk_dir
End Function

'-----------------------------------------
'実行後に作業用フォルダパスを削除するか?
'-----------------------------------------
Public Function IsDeleteWorkDir() As Boolean
    IsDeleteWorkDir = is_del_wk_dir
End Function

'-----------------------------------------
'最大実行回数
'-----------------------------------------
Public Function GetMaxExecCount() As Integer
    GetMaxExecCount = max_exec_cnt
End Function


'=========================================
'Setter
'=========================================

'-----------------------------------------
'除外ファイルリスト
'-----------------------------------------
Public Sub SetIgnoreFilePath(ByVal path As String)
    ignore_file_path = path
End Sub

'-----------------------------------------
'ツール作業用フォルダパス
'-----------------------------------------
Public Sub SetToolWorkDirPath(ByVal path As String)
    tool_wk_dir_path = path
End Sub



Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "exe_file_path=(" & exe_file_path & "), " & _
          "src_dir_path=(" & src_dir_path & "), " & _
          "dst_dir_path=(" & dst_dir_path & "), " & _
          "in_ext=(" & in_ext & "), " & _
          "addin_file_path=(" & addin_file_path & "), " & _
          "exe_ini_file_path=(" & exe_ini_file_path & "), " & _
          "is_contain_sub_dir=(" & is_contain_sub_dir_s & "), " & _
          "ignore_files=(" & ignore_files & "), " & _
          "tool_wk_dir_path=(" & tool_wk_dir_path & "), " & _
          "is_step_wk_dir=(" & is_step_wk_dir_s & "), " & _
          "is_del_wk_dir=(" & is_del_wk_dir_s & "), " & _
          "max_exec_cnt=(" & max_exec_cnt_s & ")"

    GetAllValue = ret
End Function
