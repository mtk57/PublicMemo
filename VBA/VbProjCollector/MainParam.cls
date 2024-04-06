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

Private vbprj_file_name As String
Private src_dir_path As String

Private dst_dir_path As String
Private is_out_sheet As Boolean
Private is_create_copy_bat As Boolean
Private move_base_dir As String
Private is_create_build_bat As Boolean
Private is_delete_path32 As Boolean
Private ignore_files As String

Private is_out_sheet_s As String
Private is_create_copy_bat_s As String
Private is_create_build_bat_s As String
Private is_delete_path32_s As String
Private merge_dir_path As String

Private is_continue_s As String
Private is_continue As Boolean

Private Const CLM_ = "O"
Private Const ROW_ = 6

'コンストラクタ
Private Sub Class_Initialize()
    vbprj_file_name = ""
    src_dir_path = ""
    
    dst_dir_path = ""
    is_out_sheet = False
    is_create_copy_bat = False
    move_base_dir = ""
    is_create_build_bat = False
    is_delete_path32 = False
    ignore_files = ""
    
    is_out_sheet_s = ""
    is_create_copy_bat_s = ""
    is_create_build_bat_s = ""
    is_delete_path32_s = ""
    
    merge_dir_path = ""
    
    is_continue = True
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    vbprj_file_name = main_sheet.Range(CLM_ & ROW_).value
    src_dir_path = main_sheet.Range(CLM_ & ROW_ + 1).value
    
    dst_dir_path = main_sheet.Range(CLM_ & ROW_ + 5).value
    is_out_sheet_s = main_sheet.Range(CLM_ & ROW_ + 6).value
    is_create_copy_bat_s = main_sheet.Range(CLM_ & ROW_ + 7).value
    move_base_dir = main_sheet.Range(CLM_ & ROW_ + 8).value
    is_create_build_bat_s = main_sheet.Range(CLM_ & ROW_ + 9).value
    is_delete_path32_s = main_sheet.Range(CLM_ & ROW_ + 10).value
    ignore_files = main_sheet.Range(CLM_ & ROW_ + 11).value
    merge_dir_path = main_sheet.Range(CLM_ & ROW_ + 12).value
    is_continue_s = main_sheet.Range(CLM_ & ROW_ + 13).value
End Sub

'初期化(外部実行用)
Public Sub InitForExternal( _
    ByVal dst_dir As String, _
    ByVal base_dir As String, _
    ByVal build_bat As Boolean _
)
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    
    vbprj_file_name = ""
    src_dir_path = ""
    
    dst_dir_path = dst_dir
    
    is_out_sheet_s = "NO"
    is_out_sheet = False
    
    is_create_copy_bat_s = "NO"
    is_create_copy_bat = False
    
    move_base_dir = base_dir
    
    is_create_build_bat = build_bat
    If is_create_build_bat = True Then
        is_create_build_bat_s = "YES"
    Else
        is_create_build_bat_s = "NO"
    End If
    
    is_delete_path32_s = "YES"
    is_delete_path32 = True
    
    ignore_files = main_sheet.Range(CLM_ & ROW_ + 11).value
    merge_dir_path = main_sheet.Range(CLM_ & ROW_ + 12).value
    
    is_continue_s = "YES"
    is_continue = True
End Sub

'検証
Public Sub Validate()
    Dim Ext As String

    If vbprj_file_name <> "" Then
        Ext = Common.GetFileExtension(vbprj_file_name)
    
        If Ext <> "vbp" And Ext <> "vbproj" Then
            Err.Raise 53, , "VBプロジェクトファイル名が未対応の拡張子です。(" & vbprj_file_name & ")"
        End If
    
        If src_dir_path = "" Then
            Err.Raise 53, , "VBプロジェクトファイルを検索するフォルダパスが指定されていません"
        ElseIf Common.IsExistsFolder(src_dir_path) = False Then
            Err.Raise 53, , "VBプロジェクトファイルを検索するフォルダが存在しません (" & src_dir_path & ")"
        End If
    End If

    If dst_dir_path = "" Then
        Err.Raise 53, , "コピー先フォルダパスが指定されていません"
    End If
    
    If is_out_sheet_s = "" Or is_out_sheet_s = "NO" Then
        is_out_sheet = False
    Else
        is_out_sheet = True
    End If

    If is_create_copy_bat_s = "" Or is_create_copy_bat_s = "NO" Then
        is_create_copy_bat = False
    Else
        is_create_copy_bat = True
    End If
    
    If is_create_build_bat_s = "" Or is_create_build_bat_s = "NO" Then
        is_create_build_bat = False
    Else
        is_create_build_bat = True
    End If

    If is_delete_path32_s = "" Or is_delete_path32_s = "NO" Then
        is_delete_path32 = False
    Else
        is_delete_path32 = True
    End If

    If is_continue_s = "" Or is_continue_s = "YES" Then
        is_continue = True
    Else
        is_continue = False
    End If
End Sub

'=========================================
'Getter
'=========================================

'-----------------------------------------
'VBプロジェクトファイル名
'-----------------------------------------
Public Function GetVBPrjFileName() As String
    GetVBPrjFileName = vbprj_file_name
End Function

'-----------------------------------------
'検索するフォルダパス
'-----------------------------------------
Public Function GetSrcDirPath() As String
    GetSrcDirPath = src_dir_path
End Function

'-----------------------------------------
'コピー先フォルダパス
'-----------------------------------------
Public Function GetDestDirPath() As String
    GetDestDirPath = dst_dir_path
End Function

'-----------------------------------------
'シート出力有無
'-----------------------------------------
Public Function IsOutSheet() As Boolean
    IsOutSheet = is_out_sheet
End Function

'-----------------------------------------
'コピーBATファイル出力有無
'-----------------------------------------
Public Function IsCreateCopyBat() As Boolean
    IsCreateCopyBat = is_create_copy_bat
End Function

'-----------------------------------------
'移動起点フォルダ名
'-----------------------------------------
Public Function GetMoveBaseDirName() As String
    GetMoveBaseDirName = move_base_dir
End Function

'-----------------------------------------
'ビルドBATファイル出力有無
'-----------------------------------------
Public Function IsCreateBuildBat() As Boolean
    IsCreateBuildBat = is_create_build_bat
End Function

'-----------------------------------------
'VBPファイルのPath32を削除有無
'-----------------------------------------
Public Function IsDeletePath32() As Boolean
    IsDeletePath32 = is_delete_path32
End Function

'-----------------------------------------
'マージフォルダパス
'-----------------------------------------
Public Function GetMergeDirPath() As String
    GetMergeDirPath = merge_dir_path
End Function

'-----------------------------------------
'コピー元が存在しない場合は無視して続行する
'-----------------------------------------
Public Function IsContinue() As Boolean
    IsContinue = is_continue
End Function

'-----------------------------------------
'収集除外ファイル(vbproj)
'-----------------------------------------
Public Function GetIgnoreFiles() As String
    GetIgnoreFiles = ignore_files
End Function

Public Function GetAllValue() As String
    Dim Ret As String
    
    Ret = "[Main Param] " & _
          "vbprj_file_name=(" & vbprj_file_name & "), " & _
          "src_dir_path=(" & src_dir_path & "), " & _
          "dst_dir_path=(" & dst_dir_path & "), " & _
          "is_out_sheet=(" & is_out_sheet_s & "), " & _
          "is_create_copy_bat=(" & is_create_copy_bat_s & "), " & _
          "move_base_dir=(" & move_base_dir & "), " & _
          "is_create_build_bat=(" & is_create_build_bat_s & "), " & _
          "is_delete_path32=(" & is_delete_path32_s & "), " & _
          "ignore_files=(" & ignore_files & "), " & _
          "merge_dir_path=(" & merge_dir_path & "), " & _
          "is_continue=(" & is_continue_s & ")"

    GetAllValue = Ret
End Function
