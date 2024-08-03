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

Private src_dir_path As String
Private target_type As String
Private ignore_files As String
Private target_exts As String
Private target_keys As String

Private Const CLM_ = "O"
Private Const ROW_ = 6

'コンストラクタ
Private Sub Class_Initialize()
    src_dir_path = ""
    target_type = "vbp"
    ignore_files = ""
    target_exts = ""
    target_keys = ""
End Sub

'初期化(外部実行用)
Public Sub InitForExternal( _
    ByVal arg_src_dir_path As String, _
    ByVal arg_target_type As String, _
    ByVal arg_ignore_files As String, _
    ByVal arg_target_exts As String, _
    ByVal arg_target_keys As String _
)
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    
    src_dir_path = arg_src_dir_path
    target_type = arg_target_type
    ignore_files = arg_ignore_files
    target_exts = arg_target_exts
    target_keys = arg_target_keys
End Sub


'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    src_dir_path = main_sheet.Range(CLM_ & ROW_).value
    target_type = main_sheet.Range(CLM_ & ROW_ + 2).value
    ignore_files = main_sheet.Range(CLM_ & ROW_ + 3).value
    target_exts = main_sheet.Range(CLM_ & ROW_ + 4).value
    target_keys = main_sheet.Range(CLM_ & ROW_ + 5).value
End Sub

'検証
Public Sub Validate()
    If src_dir_path = "" Then
        Err.Raise 53, , "検索するフォルダパスが指定されていません"
    ElseIf Common.IsExistsFolder(src_dir_path) = False Then
        Err.Raise 53, , "検索するフォルダが存在しません (" & src_dir_path & ")"
    End If
    
    If target_type = "" Or target_type = "vbp" Then
        target_type = "vbp"
    Else
        target_type = "vbproj"
    End If
End Sub

'=========================================
'Getter
'=========================================

Public Function GetSrcDirPath() As String
    GetSrcDirPath = src_dir_path
End Function

Public Function GetTargetType() As String
    GetTargetType = target_type
End Function

Public Function GetIgnoreFiles() As String
    GetIgnoreFiles = ignore_files
End Function

Public Function GetTargetExtensions() As String
    GetTargetExtensions = target_exts
End Function

Public Function GetTargetKeys() As String
    GetTargetKeys = target_keys
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "src_dir_path=(" & src_dir_path & "), " & _
          "target_type=(" & target_type & "), " & _
          "ignore_files=(" & ignore_files & "), " & _
          "target_exts=(" & target_exts & ")," & _
          "target_keys=(" & target_keys & ")"

    GetAllValue = ret
End Function
