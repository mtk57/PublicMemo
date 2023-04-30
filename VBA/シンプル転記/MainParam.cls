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

Private dst_not_found_action As String
Private is_copy_format As Boolean
Private is_copy_format_s As String
Private is_backup As Boolean
Private is_backup_s As String

'コンストラクタ
Private Sub Class_Initialize()
    dst_not_found_action = ""
    is_copy_format = False
    is_backup = False
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Const clm = "G"
    Const i = 6
    
    dst_not_found_action = main_sheet.Range(clm & i).value
    is_copy_format_s = main_sheet.Range(clm & i + 1).value
    is_backup_s = main_sheet.Range(clm & i + 1).value
End Sub

'検証
Public Sub Validate()
    If dst_not_found_action = "" Or _
       dst_not_found_action = "中断" Then
       dst_not_found_action = "中断"
    Else
        dst_not_found_action = "無視"
    End If

    If is_copy_format_s = "" Or _
       is_copy_format_s = "NO" Then
       is_copy_format = False
    Else
        is_copy_format = True
    End If

    If is_backup_s = "" Or _
       is_backup_s = "NO" Then
       is_backup = False
    Else
        is_backup = True
    End If
End Sub

'=========================================
'Getter
'=========================================

'-----------------------------------------
'DST検索列が見つからない場合
'-----------------------------------------
Public Function GetDestNotFoundAction() As String
    GetDestNotFoundAction = dst_not_found_action
End Function

'-----------------------------------------
'書式もコピーするか?
'-----------------------------------------
Public Function IsCopyFormat() As Boolean
    IsCopyFormat = is_copy_format
End Function

'-----------------------------------------
'バックアップ作成
'-----------------------------------------
Public Function IsBackup() As Boolean
    IsBackup = is_backup
End Function




Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "dst_not_found_action=(" & dst_not_found_action & "), " & _
          "is_copy_format=(" & is_copy_format_s & "), " & _
          "is_backup=(" & is_backup_s & ")"

    GetAllValue = ret
End Function
