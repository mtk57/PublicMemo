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
Private dst_dir_path As String    'コピー先フォルダパス(絶対パス)
Private is_copy_dir_s As String   'フォルダもコピーする
Private is_continue_s As String   'コピー元が存在しない場合は無視して続行する
Private is_overwrite_s As String  'コピー先に同名ファイルがある場合は上書きする

Private is_copy_dir As Boolean
Private is_continue As Boolean
Private is_overwrite As Boolean

Private Const CLM_ = "O"
Private Const ROW_ = 6

'コンストラクタ
Private Sub Class_Initialize()
    dst_dir_path = ""
    is_copy_dir = True
    is_continue = True
    is_overwrite = True
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    
    dst_dir_path = main_sheet.Range(CLM_ & ROW_ + 0).value
    is_copy_dir_s = main_sheet.Range(CLM_ & ROW_ + 1).value
    is_continue_s = main_sheet.Range(CLM_ & ROW_ + 2).value
    is_overwrite_s = main_sheet.Range(CLM_ & ROW_ + 3).value
End Sub

'検証
Public Sub Validate()
    Dim ext As String

    If dst_dir_path = "" Then
        Err.Raise 53, , "コピー先フォルダパスが指定されていません"
    End If
    
    If is_copy_dir_s = "" Or is_copy_dir_s = "YES" Then
        is_copy_dir = True
    Else
        is_copy_dir = False
    End If
    
    If is_continue_s = "" Or is_continue_s = "YES" Then
        is_continue = True
    Else
        is_continue = False
    End If
    
    If is_overwrite_s = "" Or is_overwrite_s = "YES" Then
        is_overwrite = True
    Else
        is_overwrite = False
    End If
End Sub

'=========================================
'Getter
'=========================================

'-----------------------------------------
'コピー先フォルダパス
'-----------------------------------------
Public Function GetDestDirPath() As String
    GetDestDirPath = dst_dir_path
End Function

'-----------------------------------------
'フォルダもコピーする
'-----------------------------------------
Public Function IsCopyDir() As Boolean
    IsCopyDir = is_copy_dir
End Function

'-----------------------------------------
'コピー元が存在しない場合は無視して続行する
'-----------------------------------------
Public Function IsContinue() As Boolean
    IsContinue = is_continue
End Function

'-----------------------------------------
'コピー先に同名ファイルがある場合は上書きする
'-----------------------------------------
Public Function IsOverWrite() As Boolean
    IsOverWrite = is_overwrite
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "dst_dir_path=(" & dst_dir_path & "), " & _
          "is_copy_dir=(" & is_copy_dir_s & "), " & _
          "is_continue=(" & is_continue_s & "), " & _
          "is_overwrite=(" & is_overwrite_s & ")"

    GetAllValue = ret
End Function
