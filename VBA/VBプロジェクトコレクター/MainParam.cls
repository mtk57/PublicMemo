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
Private is_create_bat As Boolean

Private is_out_sheet_s As String
Private is_create_bat_s As String

'コンストラクタ
Private Sub Class_Initialize()
    vbprj_file_name = ""
    src_dir_path = ""
    
    dst_dir_path = ""
    is_out_sheet = False
    is_create_bat = False
    
    is_out_sheet_s = ""
    is_create_bat_s = ""
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")
    Const CLM = "O"
    Const i = 6
    
    vbprj_file_name = main_sheet.Range(CLM & i).value
    src_dir_path = main_sheet.Range(CLM & i + 1).value
    
    dst_dir_path = main_sheet.Range(CLM & i + 5).value
    is_out_sheet_s = main_sheet.Range(CLM & i + 6).value
    is_create_bat_s = main_sheet.Range(CLM & i + 7).value
End Sub

'初期化(外部実行用)
Public Sub InitForExternal(ByVal path As String)
    dst_dir_path = path
End Sub

'検証
Public Function Validate() As String
    Dim ext As String

    If vbprj_file_name <> "" Then
        ext = Common.GetFileExtension(vbprj_file_name)
    
        If ext <> "vbp" And ext <> "vbproj" Then
            Validate = "VBプロジェクトファイル名が未対応の拡張子です。(" & vbprj_file_name & ")"
            Exit Function
        End If
    
        If src_dir_path = "" Then
            Validate = "VBプロジェクトファイルを検索するフォルダパスが指定されていません"
            Exit Function
        ElseIf Common.IsExistsFolder(src_dir_path) = False Then
            Validate = "VBプロジェクトファイルを検索するフォルダが存在しません (" & src_dir_path & ")"
            Exit Function
        End If
    End If

    If dst_dir_path = "" Then
        Validate = "コピー先フォルダパスが指定されていません"
        Exit Function
    End If
    
    If is_out_sheet_s = "" Or is_out_sheet_s = "NO" Then
        is_out_sheet = False
    Else
        is_out_sheet = True
    End If

    If is_create_bat_s = "" Or is_create_bat_s = "NO" Then
        is_create_bat = False
    Else
        is_create_bat = True
    End If

    Validate = ""
End Function

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
'BATファイル出力有無
'-----------------------------------------
Public Function IsCreateBat() As Boolean
    IsCreateBat = is_create_bat
End Function

Public Function GetAllValue() As String
    Dim ret As String
    
    ret = "[Main Param] " & _
          "vbprj_file_name=(" & vbprj_file_name & "), " & _
          "src_dir_path=(" & src_dir_path & "), " & _
          "dst_dir_path=(" & dst_dir_path & "), " & _
          "is_out_sheet=(" & is_out_sheet_s & "), " & _
          "is_create_bat=(" & is_create_bat_s & ")"

    GetAllValue = ret
End Function
