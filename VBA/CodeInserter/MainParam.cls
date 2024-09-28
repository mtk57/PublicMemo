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
Private ignore_method_keywords As String
Private is_use_reg_ex_for_ignore_method As Boolean
Private is_use_reg_ex_for_ignore_method_s As String
Private is_use_letter_case_for_ignore_method As Boolean
Private is_use_letter_case_for_ignore_method_s As String
Private is_use_exact_match_for_ignore_method As Boolean
Private is_use_exact_match_for_ignore_method_s As String

Private ignore_files_list() As String
Private ignore_files_is_empty As Boolean

Private ignore_keywords_list() As String
Private ignore_keywords_is_empty As Boolean

Private ignore_method_keywords_list() As String
Private ignore_method_keywords_is_empty As Boolean

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
    insert_seq = ""
    ignore_method_keywords = ""
    is_use_reg_ex_for_ignore_method = False
    is_use_letter_case_for_ignore_method = False
    is_use_exact_match_for_ignore_method = False
    
    ignore_files_is_empty = False
    ignore_keywords_is_empty = False
    ignore_method_keywords_is_empty = False
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
    ignore_method_keywords = main_sheet.Range(CLM_ & ROW_ + 10).value
    is_use_reg_ex_for_ignore_method_s = main_sheet.Range(CLM_ & ROW_ + 11).value
    is_use_letter_case_for_ignore_method_s = main_sheet.Range(CLM_ & ROW_ + 12).value
    is_use_exact_match_for_ignore_method_s = main_sheet.Range(CLM_ & ROW_ + 13).value
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
    
    If is_use_reg_ex_for_ignore_method_s = "" Or is_use_reg_ex_for_ignore_method_s = "NO" Then
        is_use_reg_ex_for_ignore_method = False
    Else
        is_use_reg_ex_for_ignore_method = True
    End If
    
    If is_use_letter_case_for_ignore_method_s = "" Or is_use_letter_case_for_ignore_method_s = "NO" Then
        is_use_letter_case_for_ignore_method = False
    Else
        is_use_letter_case_for_ignore_method = True
    End If
    
    If is_use_exact_match_for_ignore_method_s = "" Or is_use_exact_match_for_ignore_method_s = "NO" Then
        is_use_exact_match_for_ignore_method = False
    Else
        is_use_exact_match_for_ignore_method = True
    End If
    
    '除外ファイルリストを作成
    ignore_files_list = Split(ignore_files, ",")
    If Common.IsEmptyArray(ignore_files_list) = True Then
        ignore_files_is_empty = True
    End If
    
    '除外キーワードリストを作成
    ignore_keywords_list = Split(ignore_keywords, ",")
    If Common.IsEmptyArray(ignore_keywords_list) = True Then
        ignore_keywords_is_empty = True
    End If
    
    '除外メソッドキーワードリストを作成
    ignore_method_keywords_list = Split(ignore_method_keywords, ",")
    If Common.IsEmptyArray(ignore_method_keywords_list) = True Then
        ignore_method_keywords_is_empty = True
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

Public Function GetInsertSeq() As String
    GetInsertSeq = insert_seq
End Function

Public Function GetIgnoreMethodKeywords() As String
    GetIgnoreMethodKeywords = ignore_method_keywords
End Function

Public Function IsUseRegExForIgnoreMethod() As Boolean
    IsUseRegExForIgnoreMethod = is_use_reg_ex_for_ignore_method
End Function

Public Function IsUseLetterCaseForIgnoreMethod() As Boolean
    IsUseLetterCaseForIgnoreMethod = is_use_letter_case_for_ignore_method
End Function

Public Function IsUseExactMatchForIgnoreMethod() As Boolean
    IsUseExactMatchForIgnoreMethod = is_use_exact_match_for_ignore_method
End Function

Public Function GetIgnoreMethodKeywordsList() As String()
    GetIgnoreMethodKeywordsList = ignore_method_keywords_list
End Function

Public Function IsIgnoreMethodKeywordsEmpty() As Boolean
    IsIgnoreMethodKeywordsEmpty = ignore_method_keywords_is_empty
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
          "insert_seq=(" & insert_seq & "), " & _
          "ignore_method_keywords=(" & ignore_method_keywords & "), " & _
          "is_use_reg_ex_for_ignore_method=(" & is_use_reg_ex_for_ignore_method_s & "), " & _
          "is_use_letter_case_for_ignore_method=(" & is_use_letter_case_for_ignore_method_s & "), " & _
          "is_use_exact_match_for_ignore_method=(" & is_use_exact_match_for_ignore_method_s & ")"

    GetAllValue = ret
End Function

'除外ファイルかを返す
Public Function IsIgnoreFile(ByVal path As String) As Boolean
    If ignore_files_is_empty = True Then
        IsIgnoreFile = False
        Exit Function
    End If

    Dim i As Long
    For i = LBound(ignore_files_list) To UBound(ignore_files_list)
        If Common.GetFileName(path) = ignore_files_list(i) Then
            IsIgnoreFile = True
            Exit Function
        End If
    Next i
    
    IsIgnoreFile = False
End Function

'除外キーワードを含むかを返す
Public Function IsIgnoreKeyword(ByVal path As String) As Boolean
    If ignore_keywords_is_empty = True Then
        IsIgnoreKeyword = False
        Exit Function
    End If

    Dim i As Long
    For i = LBound(ignore_keywords_list) To UBound(ignore_keywords_list)
        If InStr(Common.GetFileName(path), ignore_keywords_list(i)) > 0 Then
            IsIgnoreKeyword = True
            Exit Function
        End If
    Next i
    
    IsIgnoreKeyword = False
End Function

'除外メソッドキーワードを含むかを返す
Public Function IsIgnoreMethodKeyword(ByVal methodName As String) As Boolean
    If ignore_method_keywords_is_empty = True Then
        IsIgnoreMethodKeyword = False
        Exit Function
    End If

    Dim i As Long
    For i = LBound(ignore_method_keywords_list) To UBound(ignore_method_keywords_list)
        If InStr(methodName, ignore_method_keywords_list(i)) > 0 Then
            IsIgnoreMethodKeyword = True
            Exit Function
        End If
    Next i
    
    IsIgnoreMethodKeyword = False
End Function
