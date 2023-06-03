VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'メンバー変数

Private src_frm_path As String
Private src_vb_path As String
Private dst_dir_path As String
Private twips_per_pixel As Long
Private twips_per_pixel_s As String

Private Const CLM_ = "L"
Private Const ROW_ = 7

'コンストラクタ
Private Sub Class_Initialize()
    src_frm_path = ""
    src_vb_path = ""
    dst_dir_path = ""
    twips_per_pixel = 0
End Sub

'初期化
Public Sub Init()
    Dim main_sheet As Worksheet
    Set main_sheet = ThisWorkbook.Sheets("main")

    src_frm_path = main_sheet.Range(CLM_ & ROW_).value
    src_vb_path = main_sheet.Range(CLM_ & ROW_ + 1).value
    dst_dir_path = main_sheet.Range(CLM_ & ROW_ + 2).value
    twips_per_pixel_s = main_sheet.Range(CLM_ & ROW_ + 3).value
End Sub

'検証
Public Sub Validate()
    If src_frm_path = "" Then
        Err.Raise 53, , "入力frmファイルパスが指定されていません"
    ElseIf Common.IsExistsFile(src_frm_path) = False Then
        Err.Raise 53, , "入力frmファイルが存在しません (" & src_frm_path & ")"
    End If

    If src_vb_path = "" Then
        Err.Raise 53, , "入力Designer.vbファイルパスが指定されていません"
    ElseIf Common.IsExistsFile(src_vb_path) = False Then
        Err.Raise 53, , "入力Designer.vbファイルが存在しません (" & src_vb_path & ")"
    End If

    If dst_dir_path = "" Then
        Err.Raise 53, , "出力先フォルダパスが指定されていません"
    End If
    
    If twips_per_pixel_s = "" Then
        twips_per_pixel = 1
    ElseIf IsNumeric(twips_per_pixel_s) Then
        twips_per_pixel = Int(twips_per_pixel_s)
        
        If twips_per_pixel < 1 Then
            Err.Raise 53, , "TwipsPerPixelが1未満です (" & twips_per_pixel_s & ")"
        End If
    Else
        Err.Raise 53, , "TwipsPerPixelが数字ではありません (" & twips_per_pixel_s & ")"
    End If
End Sub

'=========================================
'Getter
'=========================================

Public Function GetSrcFrmFilePath() As String
    GetSrcFrmFilePath = src_frm_path
End Function

Public Function GetSrcDesignerVbFilePath() As String
    GetSrcDesignerVbFilePath = src_vb_path
End Function

Public Function GetDestDirPath() As String
    GetDestDirPath = dst_dir_path
End Function

Public Function GetTwipsPerPixel() As Long
    GetTwipsPerPixel = twips_per_pixel
End Function


