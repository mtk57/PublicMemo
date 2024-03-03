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

Private Const CLM_ = "O"
Private Const ROW_ = 6

'コンストラクタ
Private Sub Class_Initialize()
    src_dir_path = ""
    target_type = "vbp"
    ignore_files = ""
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    src_dir_path = main_sheet.Range(CLM_ & ROW_).value
    target_type = main_sheet.Range(CLM_ & ROW_ + 2).value
    ignore_files = main_sheet.Range(CLM_ & ROW_ + 3).value
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

'-----------------------------------------
'検索するフォルダパス
'-----------------------------------------
Public Function GetSrcDirPath() As String
    GetSrcDirPath = src_dir_path
End Function

Public Function GetTargetType() As String
    GetTargetType = target_type
End Function

Public Function GetIgnoreFiles() As String
    GetIgnoreFiles = ignore_files
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "src_dir_path=(" & src_dir_path & "), " & _
          "target_type=(" & target_type & "), " & _
          "ignore_files=(" & ignore_files & ")"

    GetAllValue = ret
End Function
