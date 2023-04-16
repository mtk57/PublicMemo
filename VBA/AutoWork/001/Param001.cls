VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Param001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SHEET = "001"
Private Const HEADER = "[Param " & SHEET & "] "

'メンバー変数
Private external_path As String
Private dst_dir_path As String
Private vbprj_file_paths() As String
Private is_debug_log As Boolean
Private is_debug_log_s As String

'コンストラクタ
Private Sub Class_Initialize()
    external_path = ""
    dst_dir_path = ""
    is_debug_log = False
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(SHEET)
    Const CLM = "O"
    Const i = 8
    
    external_path = main_sheet.Range(CLM & i).value
    dst_dir_path = main_sheet.Range(CLM & i + 1).value
    is_debug_log_s = main_sheet.Range(CLM & i + 2).value

    CollectParam
End Sub

Private Sub CollectParam()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets(SHEET)

    Const CLM = "C"
    Dim row As Integer: row = 13
    Dim i As Integer: i = 0
    Dim path As String
    
    Do
        path = main_sheet.Range(CLM & row + i).value
        
        If path = "" Then
            Exit Do
        End If
        
        ReDim Preserve vbprj_file_paths(i)
        vbprj_file_paths(i) = path
        
        i = i + 1
    Loop
End Sub

'検証
Public Function Validate() As String
    If external_path = "" Then
        Validate = HEADER & "外部ツールファイルパスが指定されていません"
        Exit Function
    End If

    If dst_dir_path = "" Then
        Validate = HEADER & "コピー先フォルダパスが指定されていません"
        Exit Function
    End If
    
    If is_debug_log_s = "" Or is_debug_log_s = "NO" Then
        is_debug_log = False
    Else
        is_debug_log = True
    End If
    
    Dim msg As String: msg = ""
    Dim i As Integer
    Dim path As String
    Dim ext As String
    
    If GetVBProjFilePathListCount() <= 0 Then
        Validate = HEADER & "VBプロジェクトファイルが指定されていません"
        Exit Function
    End If
    
    For i = LBound(vbprj_file_paths) To UBound(vbprj_file_paths)
        path = vbprj_file_paths(i)
        If Common.IsExistsFile(path) = False Then
            msg = HEADER & "VBプロジェクトファイルが見つかりません。(" & path & ")"
            Exit For
        End If
        
        ext = Common.GetFileExtension(path)
        If ext <> "vbp" And ext <> "vbproj" Then
            msg = HEADER & "VBプロジェクトファイル名が未対応の拡張子です。(" & path & ")"
            Exit For
        End If
    Next i

    Validate = msg
End Function

'=========================================
'Getter
'=========================================

'-----------------------------------------
'VBプロジェクトファイル名
'-----------------------------------------
Public Function GetExternalPath() As String
    GetExternalPath = external_path
End Function

'-----------------------------------------
'コピー先フォルダパス
'-----------------------------------------
Public Function GetDestDirPath() As String
    GetDestDirPath = dst_dir_path
End Function

'-----------------------------------------
'デバッグログ出力
'-----------------------------------------
Public Function IsDebugLog() As String
    IsDebugLog = is_debug_log
End Function

Public Function GetVBProjFilePathList() As String()
    GetVBProjFilePathList = vbprj_file_paths()
End Function

Public Function GetVBProjFilePathListCount() As Integer
    If Common.IsEmptyArray(vbprj_file_paths) = True Then
        GetVBProjFilePathListCount = 0
        Exit Function
    End If
    
    Dim cnt As Integer: cnt = UBound(vbprj_file_paths())
    cnt = cnt + 1
    GetVBProjFilePathListCount = cnt
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = HEADER & _
          "external_path=(" & external_path & "), " & _
          "dst_dir_path=(" & dst_dir_path & "), " & _
          "is_debug_log=(" & is_debug_log_s & ")"

    GetAllValue = ret
End Function

