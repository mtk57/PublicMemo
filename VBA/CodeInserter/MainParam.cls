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

Private target_dir_path As String
Private target_ext As String
Private is_subdir As Boolean
Private is_subdir_s As String
Private insert_word As String
Private dst_dir_path As String
Private ignore_files As String
Private ignore_keywords As String
Private base_dir As String
Private insert_seq As String

Private Const CLM_ = "L"
Private Const ROW_ = 6

'コンストラクタ
Private Sub Class_Initialize()
    target_dir_path = ""
    target_ext = ""
    is_subdir = False
    insert_word = ""
    dst_dir_path = ""
    ignore_files = ""
    ignore_keywords = ""
    base_dir = ""
    insert_seq = False
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    target_dir_path = main_sheet.Range(CLM_ & ROW_).value
    target_ext = main_sheet.Range(CLM_ & ROW_ + 2).value
    is_subdir_s = main_sheet.Range(CLM_ & ROW_ + 3).value
    insert_word = main_sheet.Range(CLM_ & ROW_ + 4).value
    dst_dir_path = main_sheet.Range(CLM_ & ROW_ + 5).value
    ignore_files = main_sheet.Range(CLM_ & ROW_ + 6).value
    ignore_keywords = main_sheet.Range(CLM_ & ROW_ + 7).value
    base_dir = main_sheet.Range(CLM_ & ROW_ + 8).value
    insert_seq = main_sheet.Range(CLM_ & ROW_ + 9).value
End Sub

'検証
Public Sub Validate()
    If target_dir_path = "" Then
        Err.Raise 53, , "ファイル検索フォルダパスが指定されていません"
    ElseIf Common.IsExistsFolder(target_dir_path) = False Then
        Err.Raise 53, , "ファイル検索フォルダが存在しません (" & target_dir_path & ")"
    End If
    
    If target_ext = "" Then
        target_ext = "VB6系"
    ElseIf target_ext <> "VB6系" And _
         target_ext <> "vb" Then
        Err.Raise 53, , "対象ファイル拡張子が不正です。(" & target_ext & ")"
    End If

    If is_subdir_s = "" Or is_subdir_s = "YES" Then
        is_subdir = True
    Else
        is_subdir = False
    End If
    
    If insert_word = "" Then
        Err.Raise 53, , "挿入文字列が空です。"
    End If

    If dst_dir_path = "" Then
        Err.Raise 53, , "出力先フォルダパスが指定されていません"
    End If
    
    If insert_seq = "" Then
        insert_seq = "しない"
    End If
End Sub

'=========================================
'Getter
'=========================================

Public Function GetTargetDirPath() As String
    GetTargetDirPath = target_dir_path
End Function

Public Function GetTargetExtension() As String
    GetTargetExtension = target_ext
End Function

Public Function IsSubDir() As Boolean
    IsSubDir = is_subdir
End Function

Public Function GetInsertWord() As String
    GetInsertWord = insert_word
End Function

Public Function GetDestDirPath() As String
    GetDestDirPath = dst_dir_path
End Function

Public Function GetIgnoreFiles() As String
    GetIgnoreFiles = ignore_files
End Function

Public Function GetIgnoreKeywords() As String
    GetIgnoreKeywords = ignore_keywords
End Function

Public Function GetBaseDir() As String
    GetBaseDir = base_dir
End Function

Public Function GetInsertSeq() As Boolean
    GetInsertSeq = insert_seq
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "target_dir_path=(" & target_dir_path & "), " & _
          "target_ext=(" & target_ext & "), " & _
          "is_subdir=(" & is_subdir_s & "), " & _
          "insert_word=(" & insert_word & "), " & _
          "dst_dir_path=(" & dst_dir_path & "), " & _
          "ignore_files=(" & ignore_files & "), " & _
          "ignore_keywords=(" & ignore_keywords & "), " & _
          "base_dir=(" & base_dir & ")," & _
          "insert_seq=(" & insert_seq & ")"

    GetAllValue = ret
End Function
