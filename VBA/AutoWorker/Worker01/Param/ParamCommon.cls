VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private HEADER As String

'メンバー変数
Private proc_type As PROCESS_TYPE

Private external_path As String
Private dst_dir_path As String
Private ignore_files As String
Private is_debug_log As Boolean
Private is_debug_log_s As String
Private sakura_path As String
Private sakura_args As String

'コンストラクタ
Private Sub Class_Initialize()
    HEADER = "[Common Param] "

    proc_type = UNKNOWN
    external_path = ""
    dst_dir_path = ""
    ignore_files = ""
    is_debug_log = False
    sakura_path = ""
    sakura_args = ""
End Sub

'初期化
Public Sub Init()
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Sheets("params")
    Const clm = Define.COMMON_PARAM_CLM
    Const i = Define.COMMON_PARAM_ROW
    
    external_path = sheet.Range(clm & i).value
    dst_dir_path = sheet.Range(clm & i + 1).value
    ignore_files = sheet.Range(clm & i + 2).value
    is_debug_log_s = sheet.Range(clm & i + 3).value
    sakura_path = sheet.Range(clm & i + 4).value
    sakura_args = sheet.Range(clm & i + 5).value
End Sub

'検証
Public Sub Validate()

    If proc_type = PROCESS_TYPE.PROC_001 Then
        '①
        If external_path = "" Then
            Err.Raise 53, , HEADER & "外部ツールファイルパスが指定されていません"
        End If
        
        If Common.IsExistsFile(external_path) = False Then
            Err.Raise 53, , HEADER & "外部ツールファイルが見つかりません (" & external_path & ")"
        End If
    End If
    
    If proc_type = PROCESS_TYPE.PROC_001 Or _
       proc_type = PROCESS_TYPE.PROC_002 Then
        '②
        If dst_dir_path = "" Then
            Err.Raise 53, , HEADER & "コピー先フォルダパスが指定されていません"
        End If
    End If
    
    If is_debug_log_s = "" Or is_debug_log_s = "NO" Then
        is_debug_log = False
    Else
        is_debug_log = True
    End If
    
    If proc_type = PROCESS_TYPE.PROC_005 Then
        '④
        If sakura_path = "" Then
            Err.Raise 53, , HEADER & "sakuraエディタのパスが指定されていません"
        End If
        
        If Common.IsExistsFile(Common.RemoveQuotes(sakura_path)) = False Then
            Err.Raise 53, , HEADER & "sakuraエディタが見つかりません (" & external_path & ")"
        End If
    End If
End Sub

'=========================================
'Getter
'=========================================

Public Function GetProcessType() As PROCESS_TYPE
    GetProcessType = proc_type
End Function

Public Function GetExternalPath() As String
    GetExternalPath = external_path
End Function

Public Function GetDestDirPath() As String
    GetDestDirPath = dst_dir_path
End Function

Public Function GetIgnoreFiles() As String
    GetIgnoreFiles = ignore_files
End Function

Public Function IsDebugLog() As Boolean
    IsDebugLog = is_debug_log
End Function

Public Function GetSakuraPath() As String
    GetSakuraPath = sakura_path
End Function

Public Function GetSakuraArgs() As String
    GetSakuraArgs = sakura_args
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = HEADER & _
          "external_path=(" & external_path & "), " & _
          "dst_dir_path=(" & dst_dir_path & "), " & _
          "ignore_files=(" & ignore_files & "), " & _
          "is_debug_log=(" & is_debug_log_s & "), " & _
          "sakura_path=(" & sakura_path & "), " & _
          "sakura_args=(" & sakura_args & ")"
          
    GetAllValue = ret
End Function

'---------------------------
' Setter
'---------------------------
Public Sub SetProcessType(ByVal type_ As PROCESS_TYPE)
    proc_type = type_
End Sub
