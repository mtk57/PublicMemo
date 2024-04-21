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

Private is_debug As Boolean
Private is_debug_s As String
Private vbproj_dir_path As String
Private vbp_base_dir_path As String
Private vbproj_base_dir_path As String
Private vb_file_list_creator_path As String
Private vbp_mapping_sheet_name As String
Private rename_mapping_sheet_name As String

Private Const CLM_ = "O"
Private Const ROW_ = 7

'コンストラクタ
Private Sub Class_Initialize()
    is_debug = False
    is_debug_s = ""
    vbproj_dir_path = ""
    vbp_base_dir_path = ""
    vbproj_base_dir_path = ""
    vb_file_list_creator_path = ""
    vbp_mapping_sheet_name = ""
    rename_mapping_sheet_name = ""
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    is_debug_s = main_sheet.Range(CLM_ & ROW_ - 1).value
    vbproj_dir_path = main_sheet.Range(CLM_ & ROW_).value
    vbp_base_dir_path = main_sheet.Range(CLM_ & ROW_ + 1).value
    vbproj_base_dir_path = main_sheet.Range(CLM_ & ROW_ + 2).value
    vb_file_list_creator_path = main_sheet.Range(CLM_ & ROW_ + 3).value
    vbp_mapping_sheet_name = main_sheet.Range(CLM_ & ROW_ + 4).value
    rename_mapping_sheet_name = main_sheet.Range(CLM_ & ROW_ + 5).value
End Sub

'検証
Public Sub Validate()
    If is_debug_s = "" Or is_debug_s = "NO" Then
        is_debug = False
    Else
        is_debug = True
    End If

    If vbproj_dir_path = "" Then
        Err.Raise 53, , "vbprojを検索するフォルダパスが指定されていません"
    ElseIf Common.IsExistsFolder(vbproj_dir_path) = False Then
        Err.Raise 53, , "vbprojを検索するフォルダが存在しません (" & vbproj_dir_path & ")"
    End If
    
    If vbp_base_dir_path = "" Then
        Err.Raise 53, , "vbpのベースフォルダパスが指定されていません"
    End If
    
    If vbproj_base_dir_path = "" Then
        Err.Raise 53, , "vbprojのベースフォルダパスが指定されていません"
    End If
    
    If vb_file_list_creator_path = "" Then
        Err.Raise 53, , "VbFileListCreator.xlsmパスが指定されていません"
    ElseIf Common.IsExistsFile(vb_file_list_creator_path) = False Then
        Err.Raise 53, , "VbFileListCreator.xlsmが存在しません (" & vb_file_list_creator_path & ")"
    End If
    
    If vbp_mapping_sheet_name = "" Then
        Err.Raise 53, , "VBPマッピングのシート名が指定されていません"
    ElseIf Common.IsExistSheet(ActiveWorkbook, vbp_mapping_sheet_name) = False Then
        Err.Raise 53, , "VBPマッピングのシート名が存在しません (" & vbp_mapping_sheet_name & ")"
    End If
    
    'リネームマッピングのシート名は任意
    If rename_mapping_sheet_name <> "" Then
        If Common.IsExistSheet(ActiveWorkbook, rename_mapping_sheet_name) = False Then
            Err.Raise 53, , "リネームマッピングのシート名が存在しません (" & vbp_mapping_sheet_name & ")"
        End If
    End If
End Sub

'=========================================
'Getter
'=========================================
Public Function IsDebug() As Boolean
    IsDebug = is_debug
End Function

Public Function GetVbprojDirPath() As String
    GetVbprojDirPath = vbproj_dir_path
End Function

Public Function GetVbpBaseDirPath() As String
    GetVbpBaseDirPath = vbp_base_dir_path
End Function

Public Function GetVbprojBaseDirPath() As String
    GetVbprojBaseDirPath = vbproj_base_dir_path
End Function

Public Function GetVbFileListCreatorPath() As String
    GetVbFileListCreatorPath = vb_file_list_creator_path
End Function

Public Function GetVbpMappingSheetName() As String
    GetVbpMappingSheetName = vbp_mapping_sheet_name
End Function

Public Function GetRenameMappingSheetName() As String
    GetRenameMappingSheetName = rename_mapping_sheet_name
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "is_debug=(" & is_debug_s & "), " & _
          "vbproj_dir_path=(" & vbproj_dir_path & "), " & _
          "vbp_base_dir_path=(" & vbp_base_dir_path & "), " & _
          "vbproj_base_dir_path=(" & vbproj_base_dir_path & "), " & _
          "vb_file_list_creator_path=(" & vb_file_list_creator_path & "), " & _
          "vbp_mapping_sheet_name=(" & vbp_mapping_sheet_name & "), " & _
          "rename_mapping_sheet_name=(" & rename_mapping_sheet_name & ")"

    GetAllValue = ret
End Function
